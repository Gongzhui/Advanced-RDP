// Package rdp hosts the Remote Desktop ActiveX control of Windows
// (mstscax.dll) in a child window, through ATL's control host, without cgo.
package rdp

import (
	"errors"
	"fmt"
	"strings"
	"sync"
	"unsafe"

	"golang.org/x/sys/windows"
)

var (
	user32 = windows.NewLazySystemDLL("user32.dll")
	atl    = windows.NewLazySystemDLL("atl.dll")

	procCreateWindowExW = user32.NewProc("CreateWindowExW")
	procDestroyWindow   = user32.NewProc("DestroyWindow")
	procSetWindowPos    = user32.NewProc("SetWindowPos")
	procShowWindow      = user32.NewProc("ShowWindow")
	procGetWindow       = user32.NewProc("GetWindow")
	procSetFocus        = user32.NewProc("SetFocus")
	procGetFocus        = user32.NewProc("GetFocus")
	procIsChild         = user32.NewProc("IsChild")
	procIsWindowVisible = user32.NewProc("IsWindowVisible")

	procAtlAxWinInit       = atl.NewProc("AtlAxWinInit")
	procAtlAxAttachControl = atl.NewProc("AtlAxAttachControl")
)

const (
	wsChild        = 0x40000000
	wsClipChildren = 0x02000000
	wsClipSiblings = 0x04000000
	swHide         = 0
	swShowNA       = 8
	swpNoActivate  = 0x0010
	swpNoZOrder    = 0x0004
	gwChild        = 5
	clsctxInproc   = 1
)

// Controls to try, newest first: MsRdpClient12 to 9, not safe for
// scripting, which lets the container set the password.
var controls = []struct{ name, clsid string }{
	{"MsRdpClient12NotSafeForScripting", "{3F859AA3-C2D4-4FAA-B0E4-FD0C9C4E5E3A}"},
	{"MsRdpClient11NotSafeForScripting", "{1DF7C823-B2D4-4B54-975A-F2AC5D7CF8B8}"},
	{"MsRdpClient10NotSafeForScripting", "{A0C63C30-F08D-4AB4-907C-34905D770C7D}"},
	{"MsRdpClient9NotSafeForScripting", "{8B918B82-7985-4C24-89DF-C33AD2BBFBCD}"},
}

var iidExtendedSettings = mustGUID("{302D8188-0052-4807-806A-362B628F9AC5}")

var initOnce = sync.OnceValue(func() error {
	procOleInitialize.Call(0) // S_FALSE when the thread already is
	if err := procAtlAxWinInit.Find(); err != nil {
		return fmt.Errorf("系统缺少 atl.dll: %w", err)
	}
	if r, _, _ := procAtlAxWinInit.Call(); r == 0 {
		return errors.New("AtlAxWinInit 失败")
	}
	return nil
})

// Settings describe a connection.
type Settings struct {
	Server            string
	Port              int
	Username          string
	Domain            string
	Password          string
	DesktopWidth      int
	DesktopHeight     int
	ScaleFactor       int // the display's, in percent: 100, 125, 150…
	RedirectClipboard bool
	FullScreen        bool
}

// Client is one Remote Desktop control in a child window. Its methods must
// be called on the thread that created it, which must run a message loop:
// the control is apartment threaded and reports its events there.
type Client struct {
	hwnd    uintptr // the ATL host window
	unk     uintptr
	ctl     *Dispatch
	adv     *Dispatch
	cp      uintptr // IConnectionPoint of the events
	cookie  uint32
	Control string // the name of the control in use
	// Log receives what the client did not manage to do, for diagnosis.
	Log func(string)
}

// New creates the control in a hidden child window of parent and reports
// its events to onEvent.
func New(parent uintptr, onEvent func(Event)) (*Client, error) {
	if err := initOnce(); err != nil {
		return nil, err
	}
	c := &Client{}
	var errs []string
	for _, ctl := range controls {
		clsid := mustGUID(ctl.clsid)
		var unk uintptr
		hr, _, _ := procCoCreateInstance.Call(uintptr(unsafe.Pointer(&clsid)), 0, clsctxInproc,
			uintptr(unsafe.Pointer(&iidIUnknown)), uintptr(unsafe.Pointer(&unk)))
		if !failed(hr) && unk != 0 {
			c.unk, c.Control = unk, ctl.name
			break
		}
		errs = append(errs, fmt.Sprintf("%s: 0x%08X", ctl.name, uint32(hr)))
	}
	if c.unk == 0 {
		return nil, fmt.Errorf("无法创建远程桌面控件（请确认系统启用了“远程桌面连接”）: %s", strings.Join(errs, "; "))
	}

	class, _ := windows.UTF16PtrFromString("AtlAxWin")
	empty, _ := windows.UTF16PtrFromString("")
	hwnd, _, e := procCreateWindowExW.Call(0, uintptr(unsafe.Pointer(class)), uintptr(unsafe.Pointer(empty)),
		wsChild|wsClipChildren|wsClipSiblings, 0, 0, 640, 480, parent, 0, 0, 0)
	if hwnd == 0 {
		release(c.unk)
		return nil, fmt.Errorf("创建控件窗口失败: %v", e)
	}
	c.hwnd = hwnd
	if hr, _, _ := procAtlAxAttachControl.Call(c.unk, hwnd, 0); failed(hr) {
		c.Close()
		return nil, hrError("AtlAxAttachControl", hr)
	}
	if p := queryInterface(c.unk, &iidIDispatch); p != 0 {
		c.ctl = newDispatch(p)
	} else {
		c.Close()
		return nil, errors.New("控件不支持 IDispatch")
	}
	for _, name := range []string{"AdvancedSettings9", "AdvancedSettings8", "AdvancedSettings2"} {
		if adv, err := c.ctl.GetDispatch(name); err == nil {
			c.adv = adv
			break
		}
	}
	if c.adv == nil {
		c.Close()
		return nil, errors.New("控件没有高级设置接口")
	}
	if err := c.advise(onEvent); err != nil {
		c.Close()
		return nil, err
	}
	return c, nil
}

func (c *Client) advise(onEvent func(Event)) error {
	cpc := queryInterface(c.unk, &iidIConnectionPointContainer)
	if cpc == 0 {
		return errors.New("控件不支持事件")
	}
	defer release(cpc)
	if hr := comCall(cpc, 4, uintptr(unsafe.Pointer(&diidEvents)), uintptr(unsafe.Pointer(&c.cp))); failed(hr) {
		return hrError("FindConnectionPoint", hr)
	}
	s := newSink(onEvent)
	defer release(s) // the connection point keeps its own reference
	if hr := comCall(c.cp, 5, s, uintptr(unsafe.Pointer(&c.cookie))); failed(hr) {
		return hrError("Advise", hr)
	}
	return nil
}

func (c *Client) log(format string, args ...any) {
	if c.Log != nil {
		c.Log(fmt.Sprintf(format, args...))
	}
}

// try sets a property that only some versions or servers support, and
// logs when it cannot.
func (c *Client) try(d *Dispatch, name string, value any) {
	if err := d.Put(name, value); err != nil {
		c.log("设置 %s 失败: %v", name, err)
	}
}

// Configure applies the settings of a connection. Call it before Connect.
func (c *Client) Configure(s Settings) error {
	for _, p := range []struct {
		name  string
		value any
	}{
		{"Server", s.Server},
		{"UserName", s.Username},
		{"Domain", s.Domain},
		{"DesktopWidth", evenSize(s.DesktopWidth)},
		{"DesktopHeight", clampSize(s.DesktopHeight)},
	} {
		if err := c.ctl.Put(p.name, p.value); err != nil {
			return err
		}
	}
	c.try(c.ctl, "ColorDepth", 32)
	if err := c.adv.Put("RDPPort", s.Port); err != nil {
		return err
	}
	if err := c.adv.Put("ClearTextPassword", s.Password); err != nil {
		return fmt.Errorf("设置密码失败: %w", err)
	}
	a := c.adv
	c.try(a, "RedirectClipboard", s.RedirectClipboard)
	c.try(a, "SmartSizing", true)
	c.try(a, "EnableCredSspSupport", true)
	c.try(a, "NegotiateSecurityLayer", true)
	c.try(a, "AuthenticationLevel", 0) // connect without asking about the certificate, as before
	c.try(a, "EnableAutoReconnect", true)
	c.try(a, "MaxReconnectAttempts", 20)
	c.try(a, "keepAliveInterval", 20000) // keeps NAT and port forwards open
	// The window goes full screen for the control, which shows its
	// connection bar there: minimize, restore and close come back as events.
	c.try(a, "ContainerHandledFullScreen", 1)
	c.try(a, "DisplayConnectionBar", true)
	c.try(a, "PinConnectionBar", false)
	c.try(a, "ConnectionBarShowMinimizeButton", true)
	c.try(a, "ConnectionBarShowRestoreButton", true)
	c.try(a, "GrabFocusOnConnect", true)
	c.try(a, "BitmapPersistence", 1)
	c.try(a, "Compress", 1)
	c.try(a, "NetworkConnectionType", 7) // detect the connection's quality
	c.try(a, "BandwidthDetection", true)
	c.try(a, "EnableWindowsKey", 1)
	c.try(a, "ConnectToServerConsole", false)
	if sec, err := c.ctl.GetDispatch("SecuredSettings3"); err == nil {
		c.try(sec, "KeyboardHookMode", 2) // Windows keys go to the remote desktop in full screen
		sec.Release()
	}
	if s.ScaleFactor > 100 {
		c.setExtended("DesktopScaleFactor", uint32(s.ScaleFactor))
		c.setExtended("DeviceScaleFactor", uint32(deviceScale(s.ScaleFactor)))
	}
	c.try(c.ctl, "FullScreen", s.FullScreen)
	return nil
}

// setExtended sets a property of IMsRdpExtendedSettings, which has no
// IDispatch.
func (c *Client) setExtended(name string, value uint32) {
	ext := queryInterface(c.unk, &iidExtendedSettings)
	if ext == 0 {
		c.log("控件不支持扩展设置")
		return
	}
	defer release(ext)
	v, _ := toVariant(value)
	b := bstr(name)
	defer freeBSTR(b)
	if hr := comCall(ext, 3, b, uintptr(unsafe.Pointer(&v))); failed(hr) {
		c.log("设置 %s 失败: 0x%08X", name, uint32(hr))
	}
}

// deviceScale is the device scale factor the protocol allows for a
// desktop scale factor: 100, 140 or 180.
func deviceScale(desktop int) int {
	switch {
	case desktop >= 175:
		return 180
	case desktop >= 125:
		return 140
	}
	return 100
}

func clampSize(v int) int { return min(max(v, 200), 8192) }

// evenSize is a desktop width the protocol accepts.
func evenSize(v int) int { return clampSize(v) &^ 1 }

// Connect starts connecting; the events tell how it goes.
func (c *Client) Connect() error {
	_, err := c.ctl.Call("Connect")
	return err
}

// Disconnect starts disconnecting; EvDisconnected follows.
func (c *Client) Disconnect() error {
	_, err := c.ctl.Call("Disconnect")
	return err
}

// State is 0 when disconnected, 1 when connected and 2 while connecting.
func (c *Client) State() int {
	if c == nil || c.ctl == nil {
		return 0
	}
	v, _ := c.ctl.GetInt("Connected")
	return int(v)
}

// SetFullScreen tells the control whether its container shows it in full
// screen, which decides where the Windows keys go.
func (c *Client) SetFullScreen(on bool) {
	if v, err := c.ctl.GetInt("FullScreen"); err == nil && (v != 0) == on {
		return
	}
	c.try(c.ctl, "FullScreen", on)
}

// UpdateDisplay changes the remote desktop's resolution to fit the window,
// once logged in. Servers before Windows 8.1 cannot; SmartSizing then
// scales the desktop to the window.
func (c *Client) UpdateDisplay(width, height, scale int) error {
	if scale < 100 {
		scale = 100
	}
	w, h := uint32(evenSize(width)), uint32(clampSize(height))
	_, err := c.ctl.Call("UpdateSessionDisplaySettings", w, h, w, h, uint32(0), uint32(scale), uint32(deviceScale(scale)))
	return err
}

// ErrorDescription returns the control's description of a disconnect reason.
func (c *Client) ErrorDescription(reason int64) string {
	if c == nil || c.ctl == nil {
		return ""
	}
	ext, _ := c.ctl.GetInt("ExtendedDisconnectReason")
	s, err := c.ctl.Call("GetErrorDescription", uint32(reason), uint32(ext))
	if err != nil {
		return ""
	}
	return strings.TrimSpace(s)
}

// SetBounds places the control in its parent, in pixels, above the
// parent's other children.
func (c *Client) SetBounds(x, y, w, h int) {
	procSetWindowPos.Call(c.hwnd, 0 /* HWND_TOP */, uintptr(x), uintptr(y), uintptr(w), uintptr(h), swpNoActivate)
}

// Show shows or hides the control.
func (c *Client) Show(show bool) {
	if show {
		procShowWindow.Call(c.hwnd, swShowNA)
	} else {
		procShowWindow.Call(c.hwnd, swHide)
	}
}

// Visible reports whether the control shows.
func (c *Client) Visible() bool {
	r, _, _ := procIsWindowVisible.Call(c.hwnd)
	return r != 0
}

// Focus gives the keyboard to the control.
func (c *Client) Focus() {
	target := c.hwnd
	if child, _, _ := procGetWindow.Call(c.hwnd, gwChild); child != 0 {
		target = child
	}
	if f, _, _ := procGetFocus.Call(); f != 0 && (f == target || c.contains(f)) {
		return // already inside the control
	}
	procSetFocus.Call(target)
}

// HasFocus reports whether the keyboard is in the control.
func (c *Client) HasFocus() bool {
	f, _, _ := procGetFocus.Call()
	return f != 0 && (f == c.hwnd || c.contains(f))
}

func (c *Client) contains(w uintptr) bool {
	r, _, _ := procIsChild.Call(c.hwnd, w)
	return r != 0
}

// HWND is the window that hosts the control.
func (c *Client) HWND() uintptr { return c.hwnd }

// Close releases the control and destroys its window. Disconnect first:
// a control destroyed while connected ends the connection abruptly.
func (c *Client) Close() {
	if c.cp != 0 {
		if c.cookie != 0 {
			comCall(c.cp, 6, uintptr(c.cookie)) // Unadvise
			c.cookie = 0
		}
		release(c.cp)
		c.cp = 0
	}
	c.adv.Release()
	c.adv = nil
	c.ctl.Release()
	c.ctl = nil
	if c.hwnd != 0 {
		procDestroyWindow.Call(c.hwnd)
		c.hwnd = 0
	}
	release(c.unk)
	c.unk = 0
}
