package main

import (
	"context"
	"fmt"
	"strings"
	"time"

	"github.com/egoist/mygo"
	"github.com/egoist/mygo/ui"

	"github.com/Gongzhui/Advanced-RDP/internal/netprobe"
	"github.com/Gongzhui/Advanced-RDP/internal/rdp"
	"github.com/Gongzhui/Advanced-RDP/internal/store"
	"github.com/Gongzhui/Advanced-RDP/internal/winutil"
)

// The states of a session.
const (
	stConnecting = iota
	stConnected  // the transport is up: the remote screen shows
	stLoggedIn
	stReconnecting
	stDisconnected
)

// session is one remote desktop, in a window of its own.
type session struct {
	app      *app
	host     store.Host
	password string

	win    *mygo.Window
	client *rdp.Client

	state      int
	status     string   // what the status page says
	detail     string   // the reason of a disconnect
	probe      []string // what a probe found after a failure
	probing    bool
	started    time.Time
	fullScreen bool
	closing    bool // the user asked to close: close once disconnected
	closed     bool
	remoteW    int // the remote desktop's size
	remoteH    int
	resizeGen  int
}

// openSession opens a window for h and connects it.
func openSession(a *app, h store.Host, password string) *session {
	s := &session{app: a, host: h, password: password, fullScreen: h.FullScreen}
	opts := mygo.WindowOptions{
		Title:           fmt.Sprintf("%s — %s", h.Name, h.Address),
		BackgroundColor: "#0b0d10",
		MinWidth:        480,
		MinHeight:       320,
		FullScreen:      h.FullScreen,
		Content:         ui.View(s.view),
	}
	// Open on the display of the main window.
	if a.win != nil {
		d := mygo.Screen.DisplayMatching(a.win.Bounds())
		scale := max(d.ScaleFactor, 1)
		w, hh := int(float64(h.DesktopWidth)/scale), int(float64(h.DesktopHeight)/scale)
		w, hh = min(w, d.WorkArea.Width), min(hh, d.WorkArea.Height)
		opts.Width, opts.Height, opts.UseContentSize = w, hh, true
		opts.X = d.WorkArea.X + (d.WorkArea.Width-w)/2
		opts.Y = d.WorkArea.Y + max((d.WorkArea.Height-hh)/2-20, 0)
	}
	s.win = mygo.NewWindow(opts)
	s.win.OnResize(s.layout)
	s.win.OnMaximize(s.layout)
	s.win.OnUnmaximize(s.layout)
	s.win.OnRestore(s.layout)
	s.win.OnEnterFullScreen(func() { s.fullScreen = true; s.layout() })
	s.win.OnLeaveFullScreen(func() { s.fullScreen = false; s.layout() })
	s.win.OnFocus(func() {
		// After MyGo gives the keyboard to its own view.
		s.win.Update(func() {
			if s.client != nil && s.client.Visible() {
				s.client.Focus()
			}
		})
	})
	s.win.OnClose(func(e *mygo.CloseEvent) {
		if s.client != nil && s.client.State() != 0 {
			e.PreventDefault()
			s.close()
		}
	})
	s.win.OnClosed(s.cleanup)
	s.connect()
	return s
}

func (s *session) logf(format string, args ...any) {
	s.app.log.Printf("[%s] "+format, append([]any{s.host.Name}, args...)...)
}

// connect creates the control and connects it.
func (s *session) connect() {
	if s.client != nil {
		s.client.Close()
		s.client = nil
	}
	s.state, s.status, s.detail, s.probe = stConnecting, "正在连接…", "", nil
	s.started = time.Now()
	s.closing = false
	if s.host.FullScreen && !s.fullScreen {
		s.fullScreen = true
		s.win.SetFullScreen(true)
	}
	hwnd := s.win.NativeHandle()
	c, err := rdp.New(hwnd, s.onEvent)
	if err != nil {
		s.fail("无法启动远程桌面组件", err.Error())
		return
	}
	c.Log = func(m string) { s.logf("%s", m) }
	s.client = c
	s.logf("control %s", c.Control)

	host, port := netprobe.SplitHostPort(s.host.Address, netprobe.DefaultPort)
	scale := winutil.Scale(hwnd)
	w, h := s.host.DesktopWidth, s.host.DesktopHeight
	if s.fullScreen {
		w, h = winutil.MonitorSize(hwnd)
	} else if cw, ch := winutil.ClientSize(hwnd); cw > 0 && ch > 0 {
		w, h = cw, ch
	}
	s.remoteW, s.remoteH = w, h
	err = c.Configure(rdp.Settings{
		Server: host, Port: port,
		Username: s.host.Username, Domain: s.host.Domain, Password: s.password,
		DesktopWidth: w, DesktopHeight: h, ScaleFactor: scale,
		RedirectClipboard: s.host.RedirectClipboard,
		FullScreen:        s.fullScreen,
	})
	if err != nil {
		s.fail("连接参数无效", err.Error())
		return
	}
	s.layout()
	// The control shows once connected; until then the view shows how
	// connecting goes.
	if err := c.Connect(); err != nil {
		s.fail("无法开始连接", err.Error())
		return
	}
	s.logf("connecting to %s:%d as %q, %dx%d at %d%%", host, port, s.host.Username, w, h, scale)
	s.tick()
}

// tick redraws the status page each second while connecting, and probes
// the network when connecting takes long.
func (s *session) tick() {
	time.AfterFunc(time.Second, func() {
		s.win.Update(func() {
			if s.closed || s.state != stConnecting {
				return
			}
			if time.Since(s.started) > 8*time.Second && !s.probing && s.probe == nil {
				s.runProbe()
			}
			s.tick()
		})
	})
}

func (s *session) fail(status, detail string) {
	s.logf("%s: %s", status, detail)
	if s.client != nil {
		s.client.Show(false)
	}
	s.state, s.status, s.detail = stDisconnected, status, detail
	if s.fullScreen {
		// Show why in a window, not over the whole screen.
		s.fullScreen = false
		s.win.SetFullScreen(false)
	}
	s.app.redraw()
	s.win.Invalidate()
}

func (s *session) onEvent(e rdp.Event) {
	if s.closed {
		return
	}
	switch e.ID {
	case rdp.EvConnecting:
		s.status = "正在连接…"
	case rdp.EvConnected:
		s.state, s.status = stConnected, "已连接"
		s.client.Show(true)
		s.client.Focus()
		s.logf("connected")
	case rdp.EvLoginComplete:
		s.state, s.status = stLoggedIn, "已登录"
		s.probe = nil
		s.logf("logged in")
		s.app.connected(s.host.Name)
		s.fitRemote()
	case rdp.EvAutoReconnecting:
		s.state = stReconnecting
		attempt := int64(0)
		if len(e.Args) > 1 {
			attempt = e.Args[1]
		}
		s.status = fmt.Sprintf("网络中断，正在重连（第 %d 次）", attempt)
		s.logf("auto reconnecting %v", e.Args)
	case rdp.EvAutoReconnected:
		s.state, s.status = stLoggedIn, "已重新连接"
		s.logf("auto reconnected")
	case rdp.EvRemoteDesktopSizeChange:
		if len(e.Args) == 2 {
			s.remoteW, s.remoteH = int(e.Args[0]), int(e.Args[1])
		}
	case rdp.EvRequestGoFullScreen: // Ctrl+Alt+Break
		s.setFullScreen(true)
	case rdp.EvRequestLeaveFullScreen: // the connection bar's restore button, or Ctrl+Alt+Break
		s.setFullScreen(false)
	case rdp.EvRequestMinimize: // the connection bar's minimize button
		s.win.Minimize()
	case rdp.EvLogonError, rdp.EvWarning, rdp.EvFatalError:
		s.logf("event %d %v", e.ID, e.Args)
		if e.ID == rdp.EvFatalError {
			s.fail("远程桌面组件出错", fmt.Sprintf("错误代码 %v", e.Args))
		}
	case rdp.EvDisconnected:
		var reason int64
		if len(e.Args) > 0 {
			reason = e.Args[0]
		}
		s.disconnected(reason)
	}
	s.app.redraw()
	s.win.Invalidate()
}

// disconnected handles the end of the connection: the window closes when
// the user ended it, here or by signing out, and otherwise says why.
func (s *session) disconnected(reason int64) {
	desc := s.client.ErrorDescription(reason)
	s.logf("disconnected, reason %d: %s", reason, desc)
	if s.closing || reason == 1 || reason == 2 || reason == 0xB08 /* logged off */ {
		s.win.Close()
		return
	}
	status := "连接已断开"
	neverUp := s.state == stConnecting
	if neverUp {
		status = "无法连接"
	}
	if desc == "" {
		desc = fmt.Sprintf("断开代码 %d", reason)
	}
	s.fail(status, desc)
	if neverUp && !s.probing {
		s.runProbe()
	}
}

// runProbe checks the network to explain a failure.
func (s *session) runProbe() {
	s.probing = true
	host, port := netprobe.SplitHostPort(s.host.Address, netprobe.DefaultPort)
	go func() {
		r := netprobe.Probe(context.Background(), host, port, 5*time.Second)
		s.win.Update(func() {
			s.probing = false
			s.probe = r.Lines()
		})
	}()
}

// layout fits the control to the window, and the remote desktop to the
// control once the window settles.
func (s *session) layout() {
	if s.client == nil || s.closed {
		return
	}
	w, h := winutil.ClientSize(s.win.NativeHandle())
	s.client.SetBounds(0, 0, w, h)
	s.resizeGen++
	gen := s.resizeGen
	time.AfterFunc(400*time.Millisecond, func() {
		s.win.Update(func() {
			if gen == s.resizeGen {
				s.fitRemote()
			}
		})
	})
}

// fitRemote changes the remote resolution to the window's size, where the
// server can; elsewhere the control scales the desktop to the window.
func (s *session) fitRemote() {
	if s.client == nil || s.closed || s.state != stLoggedIn {
		return
	}
	hwnd := s.win.NativeHandle()
	w, h := winutil.ClientSize(hwnd)
	if w < 200 || h < 200 || (abs(w-s.remoteW) <= 2 && abs(h-s.remoteH) <= 2) {
		return
	}
	if err := s.client.UpdateDisplay(w, h, winutil.Scale(hwnd)); err != nil {
		s.logf("resize remote desktop to %dx%d: %v (scaling instead)", w, h, err)
		return
	}
	s.logf("remote desktop resized to %dx%d", w, h)
}

func abs(v int) int {
	if v < 0 {
		return -v
	}
	return v
}

func (s *session) setFullScreen(on bool) {
	if s.fullScreen == on && s.win.IsFullScreen() == on {
		return
	}
	s.fullScreen = on
	s.win.SetFullScreen(on)
	if s.client != nil {
		s.client.SetFullScreen(on)
	}
	s.win.Focus()
}

// close disconnects and closes the window; it closes anyway after a few
// seconds if the server does not answer.
func (s *session) close() {
	if s.closed {
		return
	}
	s.closing = true
	if s.client == nil || s.client.State() == 0 {
		s.win.Close()
		return
	}
	s.status = "正在断开…"
	if err := s.client.Disconnect(); err != nil {
		s.logf("disconnect: %v", err)
		s.win.Destroy()
		return
	}
	time.AfterFunc(4*time.Second, func() {
		s.win.Update(func() {
			if !s.closed {
				s.logf("no answer to disconnect, closing")
				s.win.Destroy()
			}
		})
	})
}

func (s *session) cleanup() {
	if s.closed {
		return
	}
	s.closed = true
	if s.client != nil {
		s.client.Close()
		s.client = nil
	}
	if s.app.sessions[s.host.Name] == s {
		delete(s.app.sessions, s.host.Name)
	}
	s.logf("closed")
	s.app.redraw()
}

// front brings the session's window to the front.
func (s *session) front() {
	if s.win == nil {
		return
	}
	if s.win.IsMinimized() {
		s.win.Restore()
	}
	s.win.Show()
	s.win.Focus()
}

func (s *session) color(t *ui.Theme) ui.Color {
	switch s.state {
	case stLoggedIn, stConnected:
		return t.Success
	case stConnecting, stReconnecting:
		return t.Warning
	}
	return t.Danger
}

func (s *session) statusText() string {
	switch s.state {
	case stConnecting:
		return "正在连接"
	case stConnected:
		return "已连接"
	case stLoggedIn:
		return "会话中"
	case stReconnecting:
		return "正在重连"
	}
	return "已断开"
}

// view is what the window shows where the control does not: the progress
// of connecting, and why a connection ended.
func (s *session) view(c *ui.Context) {
	c.SetTheme(ui.DarkTheme())
	t := c.Theme()
	c.Root().Background(ui.Hex("#0b0d10"))
	if s.client != nil && s.client.Visible() && s.state != stDisconnected {
		return // the control covers the window
	}
	if s.state == stDisconnected {
		if c.Shortcut(0, ui.KeyEnter) {
			s.connect()
		}
		if c.Shortcut(0, ui.KeyEscape) {
			s.win.Close()
		}
	}
	ui.Column(c).Fill().Center().Children(func() {
		ui.Column(c).Width(520).MaxWidthPercent(92).Gap(14).Padding(28).Radius(12).Background(t.Surface).Border(1, t.Border).Children(func() {
			ui.Column(c).Gap(4).Children(func() {
				ui.Text(c, s.host.Name).FontSize(20).Bold()
				ui.Text(c, s.host.Address).FontSize(12).TextColor(t.TextMuted)
			})
			switch s.state {
			case stDisconnected:
				ui.Text(c, s.status).FontSize(16).Bold().TextColor(t.Danger)
				if s.detail != "" {
					ui.Text(c, s.detail).TextColor(t.TextMuted).Wrap().Selectable()
				}
			default:
				ui.Row(c).Gap(10).AlignItems(ui.Center).Children(func() {
					ui.Spinner(c).Size(16, 16)
					ui.Text(c, s.status)
					ui.Textf(c, "%d 秒", int(time.Since(s.started).Seconds())).TextColor(t.TextMuted)
				})
			}
			if s.probing {
				ui.Text(c, "正在检测网络…").FontSize(12).TextColor(t.TextMuted)
			}
			for _, l := range s.probe {
				ui.Text(c, l).FontSize(12).TextColor(t.TextMuted).Wrap().Selectable()
			}
			ui.Row(c).Gap(8).Children(func() {
				if s.state == stDisconnected {
					ui.PrimaryButton(c.Key("reconnect"), "重新连接").Tooltip("Enter").OnClick(s.connect)
				}
				label := "取消"
				if s.state == stDisconnected {
					label = "关闭"
				}
				ui.Button(c.Key("close"), label).Tooltip("Esc").OnClick(s.close)
			})
			if s.state == stDisconnected && strings.Contains(s.detail, "凭据") {
				ui.Text(c, "检查主窗口里的用户名和密码后重新连接。").FontSize(12).TextColor(t.TextMuted)
			}
		})
	})
}

// redraw redraws the main window, whose list shows the sessions.
func (a *app) redraw() {
	if a.win != nil {
		a.win.Invalidate()
	}
}
