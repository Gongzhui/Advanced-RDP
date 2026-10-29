package rdp

import (
	"image"
	"image/png"
	"os"
	"path/filepath"
	"runtime"
	"syscall"
	"testing"
	"time"
	"unsafe"

	"golang.org/x/sys/windows"
)

var (
	procRegisterClassExW = user32.NewProc("RegisterClassExW")
	procDefWindowProcW   = user32.NewProc("DefWindowProcW")
	procPeekMessageW     = user32.NewProc("PeekMessageW")
	procTranslateMessage = user32.NewProc("TranslateMessage")
	procDispatchMessageW = user32.NewProc("DispatchMessageW")
	procGetWindowRect    = user32.NewProc("GetWindowRect")
	procGetDC            = user32.NewProc("GetDC")
	procReleaseDC        = user32.NewProc("ReleaseDC")
	procSetForeground    = user32.NewProc("SetForegroundWindow")
	gdi32                = windows.NewLazySystemDLL("gdi32.dll")
	procCreateCompatDC   = gdi32.NewProc("CreateCompatibleDC")
	procCreateDIBSection = gdi32.NewProc("CreateDIBSection")
	procSelectObject     = gdi32.NewProc("SelectObject")
	procBitBlt           = gdi32.NewProc("BitBlt")
	procDeleteObject     = gdi32.NewProc("DeleteObject")
	procDeleteDC         = gdi32.NewProc("DeleteDC")
)

type wndClassEx struct {
	Size       uint32
	Style      uint32
	WndProc    uintptr
	ClsExtra   int32
	WndExtra   int32
	Instance   uintptr
	Icon       uintptr
	Cursor     uintptr
	Background uintptr
	MenuName   *uint16
	ClassName  *uint16
	IconSm     uintptr
}

type msg struct {
	Hwnd    uintptr
	Message uint32
	WParam  uintptr
	LParam  uintptr
	Time    uint32
	Pt      [2]int32
	_       uint32
}

func testWindow(t *testing.T, w, h int) uintptr {
	t.Helper()
	name, _ := windows.UTF16PtrFromString("ArdpTestWindow")
	wc := wndClassEx{WndProc: syscall.NewCallback(func(hwnd, m, wp, lp uintptr) uintptr {
		r, _, _ := procDefWindowProcW.Call(hwnd, m, wp, lp)
		return r
	}), ClassName: name, Background: 6 /* COLOR_WINDOW+1 */}
	wc.Size = uint32(unsafe.Sizeof(wc))
	procRegisterClassExW.Call(uintptr(unsafe.Pointer(&wc)))
	title, _ := windows.UTF16PtrFromString("Advanced RDP test")
	const wsOverlappedWindow, wsVisible, wsClipChildren = 0x00CF0000, 0x10000000, 0x02000000
	hwnd, _, err := procCreateWindowExW.Call(0, uintptr(unsafe.Pointer(name)), uintptr(unsafe.Pointer(title)),
		wsOverlappedWindow|wsVisible|wsClipChildren, 40, 40, uintptr(w), uintptr(h), 0, 0, 0, 0)
	if hwnd == 0 {
		t.Fatal(err)
	}
	t.Cleanup(func() { procDestroyWindow.Call(hwnd) })
	procSetForeground.Call(hwnd)
	return hwnd
}

// pump runs the message loop until done returns true or d passes.
func pump(d time.Duration, done func() bool) bool {
	deadline := time.Now().Add(d)
	var m msg
	for time.Now().Before(deadline) {
		for {
			r, _, _ := procPeekMessageW.Call(uintptr(unsafe.Pointer(&m)), 0, 0, 0, 1)
			if r == 0 {
				break
			}
			procTranslateMessage.Call(uintptr(unsafe.Pointer(&m)))
			procDispatchMessageW.Call(uintptr(unsafe.Pointer(&m)))
		}
		if done != nil && done() {
			return true
		}
		time.Sleep(10 * time.Millisecond)
	}
	return false
}

// capture saves what the screen shows of hwnd, as DWM composed it.
func capture(t *testing.T, hwnd uintptr, path string) image.Image {
	var r [4]int32
	procGetWindowRect.Call(hwnd, uintptr(unsafe.Pointer(&r)))
	w, h := int(r[2]-r[0]), int(r[3]-r[1])
	screen, _, _ := procGetDC.Call(0)
	defer procReleaseDC.Call(0, screen)
	mem, _, _ := procCreateCompatDC.Call(screen)
	defer procDeleteDC.Call(mem)
	bi := struct {
		Size                         uint32
		Width, Height                int32
		Planes, BitCount             uint16
		Compression, SizeImage       uint32
		XPels, YPels, ClrUsed, ClrIm int32
	}{Width: int32(w), Height: int32(-h), Planes: 1, BitCount: 32}
	bi.Size = uint32(unsafe.Sizeof(bi))
	var bits uintptr
	bmp, _, _ := procCreateDIBSection.Call(mem, uintptr(unsafe.Pointer(&bi)), 0, uintptr(unsafe.Pointer(&bits)), 0, 0)
	defer procDeleteObject.Call(bmp)
	procSelectObject.Call(mem, bmp)
	procBitBlt.Call(mem, 0, 0, uintptr(w), uintptr(h), screen, uintptr(r[0]), uintptr(r[1]), 0x00CC0020)
	src := unsafe.Slice((*byte)(native(bits)), w*h*4)
	img := image.NewRGBA(image.Rect(0, 0, w, h))
	for i := 0; i < w*h; i++ {
		img.Pix[4*i], img.Pix[4*i+1], img.Pix[4*i+2], img.Pix[4*i+3] = src[4*i+2], src[4*i+1], src[4*i], 255
	}
	f, err := os.Create(path)
	if err == nil {
		png.Encode(f, img)
		f.Close()
		t.Logf("screen capture: %s", path)
	}
	return img
}

// TestConnectLocalLogonScreen connects to this computer's RDP service with
// a user that does not exist: the control is created and configured, the
// connection is made and its events arrive, without signing anyone in.
// Run with ARDP_E2E=1.
func TestConnectLocalLogonScreen(t *testing.T) {
	if os.Getenv("ARDP_E2E") == "" {
		t.Skip("set ARDP_E2E=1")
	}
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()

	parent := testWindow(t, 1040, 700)
	var events []Event
	c, err := New(parent, func(e Event) { events = append(events, e) })
	if err != nil {
		t.Fatal(err)
	}
	defer c.Close()
	c.Log = func(s string) { t.Log("client:", s) }
	t.Logf("control: %s", c.Control)
	c.SetBounds(0, 0, 1024, 660)
	c.Show(true)
	if err := c.Configure(Settings{
		Server: "127.0.0.2", Port: 3389, Username: "ardp-e2e-no-such-user", Password: "not-a-password",
		DesktopWidth: 1024, DesktopHeight: 660, RedirectClipboard: true,
	}); err != nil {
		t.Fatal(err)
	}
	if err := c.Connect(); err != nil {
		t.Fatal(err)
	}
	has := func(id int) bool {
		for _, e := range events {
			if e.ID == id {
				return true
			}
		}
		return false
	}
	pump(25*time.Second, func() bool { return has(EvConnected) || has(EvDisconnected) })
	pump(4*time.Second, nil) // let the logon screen draw
	t.Logf("events: %+v, state %d", events, c.State())
	if !has(EvConnecting) {
		t.Error("no OnConnecting event")
	}
	if !has(EvConnected) && !has(EvDisconnected) {
		t.Fatal("neither connected nor disconnected")
	}
	for _, e := range events {
		if e.ID == EvDisconnected {
			t.Logf("disconnected %d: %s", e.Args[0], c.ErrorDescription(e.Args[0]))
		}
	}
	c.Focus()
	t.Logf("focus in control: %v", c.HasFocus())
	capture(t, parent, filepath.Join(os.TempDir(), "ardp-e2e-logon.png"))

	if c.State() != 0 {
		if err := c.Disconnect(); err != nil {
			t.Fatal(err)
		}
		if !pump(10*time.Second, func() bool { return has(EvDisconnected) }) {
			t.Fatal("no OnDisconnected after Disconnect")
		}
		last := events[len(events)-1]
		t.Logf("after Disconnect: %+v (%s)", last, c.ErrorDescription(last.Args[0]))
	}
}

// TestConnectRefused connects to a port nobody listens on and expects a
// disconnect event with a reason the control describes.
func TestConnectRefused(t *testing.T) {
	if os.Getenv("ARDP_E2E") == "" {
		t.Skip("set ARDP_E2E=1")
	}
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()
	parent := testWindow(t, 640, 480)
	var events []Event
	c, err := New(parent, func(e Event) { events = append(events, e) })
	if err != nil {
		t.Fatal(err)
	}
	defer c.Close()
	c.Configure(Settings{Server: "127.0.0.1", Port: 1, Username: "x", Password: "y", DesktopWidth: 800, DesktopHeight: 600})
	if err := c.Connect(); err != nil {
		t.Fatal(err)
	}
	var reason int64 = -1
	pump(30*time.Second, func() bool {
		for _, e := range events {
			if e.ID == EvDisconnected {
				reason = e.Args[0]
				return true
			}
		}
		return false
	})
	if reason < 0 {
		t.Fatalf("no disconnect event: %+v", events)
	}
	t.Logf("reason %d: %q", reason, c.ErrorDescription(reason))
	if reason == 1 {
		t.Error("a refused connection reported as a local disconnect")
	}
}
