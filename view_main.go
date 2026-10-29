package main

import (
	"context"
	"fmt"
	"strconv"
	"strings"
	"time"

	"github.com/egoist/mygo/ui"

	"github.com/Gongzhui/Advanced-RDP/internal/netprobe"
)

// resolutions a windowed session can start at; it then follows the window.
var resolutions = []string{"1280 × 720", "1366 × 768", "1600 × 900", "1920 × 1080", "2560 × 1440"}

func resolutionLabel(w, h int) string { return fmt.Sprintf("%d × %d", w, h) }

func parseResolution(s string) (int, int, bool) {
	ws, hs, ok := strings.Cut(s, " × ")
	w, err1 := strconv.Atoi(ws)
	h, err2 := strconv.Atoi(hs)
	return w, h, ok && err1 == nil && err2 == nil
}

func (a *app) view(c *ui.Context) {
	t := c.Theme()
	if c.Shortcut(ui.Cmd, ui.KeyN) {
		a.startNew(c)
	}
	if c.Shortcut(ui.Cmd, ui.KeyS) {
		a.saveWithToast(c)
	}
	if c.Shortcut(ui.Cmd, ui.KeyEnter) {
		a.connect(c)
	}
	if c.Shortcut(ui.Cmd, ui.KeyF) {
		a.searchField.Focus(c)
	}
	if c.Shortcut(ui.Cmd, ui.KeyT) {
		a.runProbe()
	}

	ui.Row(c).Fill().AlignItems(ui.Stretch).Children(func() {
		ui.Column(c).Width(268).Background(t.Surface).Padding(14, 12).Gap(10).Children(func() {
			a.sidebar(c)
		})
		ui.Box(c).Width(1).Background(t.Border)
		ui.Scroll(c).Grow(1).Children(func() {
			ui.Column(c).Padding(24, 28).Gap(18).MaxWidth(760).Children(func() {
				a.details(c)
			})
		})
	})

	switch ui.AlertDialog(c, &a.confirmDelete, fmt.Sprintf("删除 “%s”？", a.draft.Name),
		"主机和保存的密码都会删除，不能撤销。", "取消", "删除") {
	case 1:
		if err := a.remove(); err != nil {
			c.Toast(err.Error())
		}
	}
}

func (a *app) sidebar(c *ui.Context) {
	t := c.Theme()
	ui.Row(c).AlignItems(ui.Center).Gap(8).Children(func() {
		ui.Text(c, "Advanced RDP").FontSize(16).Bold().Grow(1)
		ui.Button(c.Key("new"), "+ 新建").Tooltip("新建主机 (Ctrl+N)").OnClick(func() { a.startNew(c) })
	})
	if ui.SearchField(c, &a.query).Placeholder("搜索 (Ctrl+F)").Bind(&a.searchField).Changed() {
		a.refilter()
	}

	list := ui.List(c, &a.list, len(a.visible), func(i int) {
		h := a.hosts[a.visible[i]]
		chosen := i == a.listSel
		text, muted := t.Text, t.TextMuted
		if chosen {
			text, muted = t.AccentText, t.AccentText
		}
		ui.Row(c).Padding(7, 10).Gap(10).AlignItems(ui.Center).Children(func() {
			dot := ui.Box(c).Size(8, 8).Radius(4)
			if s := a.sessions[h.Name]; s != nil && !s.closed {
				dot.Background(s.color(t)).Tooltip(s.statusText())
			} else {
				dot.Background(ui.Transparent)
			}
			ui.Column(c).Grow(1).Gap(2).Children(func() {
				ui.Text(c, h.Name).Bold().TextColor(text).NoWrap().Ellipsis("…")
				sub := h.Address
				if h.Username != "" {
					sub = h.Username + " @ " + h.Address
				}
				ui.Text(c, sub).FontSize(12).TextColor(muted).NoWrap().Ellipsis("…")
			})
		})
	}).Grow(1).Radius(t.Radius)
	list.Children(func() {
		if len(a.visible) == 0 {
			msg := "还没有主机。点 “+ 新建” 添加。"
			if a.query != "" {
				msg = "没有匹配的主机。"
			}
			ui.Text(c, msg).TextColor(t.TextMuted).Padding(12)
		}
	})
	if list.Changed() && a.listSel >= 0 && a.listSel < len(a.visible) {
		a.choose(c, a.visible[a.listSel])
	}
	if list.Submitted() {
		a.connect(c)
	}

	open := 0
	for _, s := range a.sessions {
		if !s.closed {
			open++
		}
	}
	footer := fmt.Sprintf("%d 台主机", len(a.hosts))
	if open > 0 {
		footer += fmt.Sprintf(" · %d 个会话", open)
	}
	ui.Text(c, footer).FontSize(12).TextColor(t.TextMuted)
}

func (a *app) startNew(c *ui.Context) {
	if a.dirty() && a.sel >= 0 {
		if err := a.save(); err != nil {
			c.Toast(err.Error())
			return
		}
	}
	a.newHost()
	a.nameField.Focus(c)
}

func (a *app) saveWithToast(c *ui.Context) {
	if !a.dirty() && a.sel >= 0 {
		return
	}
	if err := a.save(); err != nil {
		c.Toast(err.Error())
		return
	}
	c.Toast("已保存")
}

func (a *app) details(c *ui.Context) {
	t := c.Theme()
	if a.loadErr != "" {
		ui.Text(c, "读取主机列表出错："+a.loadErr).TextColor(t.Danger).Wrap()
	}

	ui.Row(c).AlignItems(ui.Center).Gap(12).Children(func() {
		ui.Column(c).Grow(1).Gap(4).Children(func() {
			title := strings.TrimSpace(a.draft.Name)
			if title == "" {
				title = "新主机"
			}
			ui.Text(c, title).FontSize(22).Bold().NoWrap().Ellipsis("…")
			sub := "填写地址和账号后按 Enter 连接"
			if a.sel >= 0 {
				sub = lastUsed(a.hosts[a.sel].LastConnected)
			}
			if a.dirty() && a.sel >= 0 {
				sub += " · 有未保存的修改"
			}
			ui.Text(c, sub).FontSize(12).TextColor(t.TextMuted)
		})
		if a.sel >= 0 {
			if s := a.sessions[a.hosts[a.sel].Name]; s != nil && !s.closed {
				ui.Row(c).Gap(6).AlignItems(ui.Center).Padding(4, 10).Radius(12).Background(t.SurfaceHover).Children(func() {
					ui.Box(c).Size(8, 8).Radius(4).Background(s.color(t))
					ui.Text(c, s.statusText()).FontSize(12)
				})
			}
		}
	})

	submit := func(e ui.Element) {
		if e.Submitted() {
			a.connect(c)
		}
	}
	ui.Grid(c).ColumnTracks(ui.Fr(1), ui.Fr(1)).GapX(16).GapY(14).Children(func() {
		ui.Field(c, "名称", func() {
			submit(ui.TextInput(c.Key("name"), &a.draft.Name).Placeholder("例如：实验室 3400").Bind(&a.nameField))
		})
		ui.Field(c, "地址", func() {
			submit(ui.TextInput(c.Key("address"), &a.draft.Address).Placeholder("主机名或 IP，可带 :端口").InputMethod(false))
		})
		ui.Field(c, "用户名", func() {
			submit(ui.TextInput(c.Key("user"), &a.draft.Username).Placeholder("用户名"))
		})
		ui.Field(c, "域（可选）", func() {
			submit(ui.TextInput(c.Key("domain"), &a.draft.Domain).Placeholder("工作组时留空"))
		})
		ui.Field(c, "密码", func() {
			ui.Row(c).Gap(8).Children(func() {
				// No input method: Pinyin would turn the letters of a
				// password into characters, as Windows' own password boxes avoid.
				in := ui.TextInput(c.Key("password"), &a.password).Placeholder("密码").Grow(1).InputMethod(false)
				if !a.showPassword {
					in.Password()
				}
				submit(in)
				label := "显示"
				if a.showPassword {
					label = "隐藏"
				}
				ui.Button(c.Key("reveal"), label).OnClick(func() { a.showPassword = !a.showPassword })
			})
		}).ColumnSpan(2)
	})

	ui.Row(c).Gap(24).Wrap().Children(func() {
		ui.Checkbox(c.Key("fullscreen"), &a.draft.FullScreen, "全屏")
		ui.Checkbox(c.Key("clipboard"), &a.draft.RedirectClipboard, "共享剪贴板")
		ui.Checkbox(c.Key("remember"), &a.savePassword, "记住密码")
	})
	if !a.draft.FullScreen {
		res := resolutionLabel(a.draft.DesktopWidth, a.draft.DesktopHeight)
		options := resolutions
		if !contains(options, res) {
			options = append([]string{res}, options...)
		}
		ui.Field(c, "窗口初始分辨率（之后随窗口大小自动调整）", func() {
			if ui.Select(c.Key("resolution"), &res, options).Width(220).Changed() {
				if w, h, ok := parseResolution(res); ok {
					a.draft.DesktopWidth, a.draft.DesktopHeight = w, h
				}
			}
		})
	}

	ui.Row(c).Gap(8).AlignItems(ui.Center).Children(func() {
		label := "连接"
		if a.sel >= 0 {
			if s := a.sessions[a.hosts[a.sel].Name]; s != nil && !s.closed {
				label = "切换到会话"
			}
		}
		ui.PrimaryButton(c.Key("connect"), label).Tooltip("Enter 或 Ctrl+Enter").OnClick(func() { a.connect(c) })
		ui.Button(c.Key("save"), "保存").Tooltip("Ctrl+S").Disabled(!a.dirty()).OnClick(func() { a.saveWithToast(c) })
		ui.Button(c.Key("probe"), "测试连接").Tooltip("检查域名、端口和 RDP 服务 (Ctrl+T)").
			Disabled(a.probe.running || strings.TrimSpace(a.draft.Address) == "").OnClick(a.runProbe)
		ui.Spacer(c)
		if a.sel >= 0 {
			ui.Button(c.Key("delete"), "删除").TextColor(t.Danger).OnClick(func() { a.confirmDelete = true })
		}
	})

	if a.probe.running || len(a.probe.lines) > 0 {
		ui.Column(c).Gap(6).Padding(12, 14).Radius(t.Radius).Background(t.Surface).Border(1, t.Border).Children(func() {
			if a.probe.running {
				ui.Row(c).Gap(8).AlignItems(ui.Center).Children(func() {
					ui.Spinner(c).Size(14, 14)
					ui.Text(c, "正在测试 "+a.probe.host+" …")
				})
				return
			}
			head, color := "可以连接", t.Success
			if !a.probe.ok {
				head, color = "无法连接", t.Danger
			}
			ui.Text(c, head).Bold().TextColor(color)
			for _, l := range a.probe.lines {
				ui.Text(c, l).FontSize(12).TextColor(t.TextMuted).Wrap().Selectable()
			}
		})
	}

	ui.Text(c, "全屏时，把鼠标移到屏幕顶端中间会出现连接栏：可最小化、还原为窗口或断开。Ctrl+Alt+Break 在全屏和窗口之间切换。").
		FontSize(12).TextColor(t.TextMuted).Wrap()
}

func contains(list []string, s string) bool {
	for _, x := range list {
		if x == s {
			return true
		}
	}
	return false
}

func lastUsed(t time.Time) string {
	if t.IsZero() {
		return "尚未用新版连接过"
	}
	d := time.Since(t)
	switch {
	case d < time.Minute:
		return "刚刚连接过"
	case d < time.Hour:
		return fmt.Sprintf("%d 分钟前连接过", int(d.Minutes()))
	case d < 24*time.Hour:
		return fmt.Sprintf("%d 小时前连接过", int(d.Hours()))
	case d < 30*24*time.Hour:
		return fmt.Sprintf("%d 天前连接过", int(d.Hours()/24))
	}
	return "上次连接：" + t.Format("2006-01-02")
}

// runProbe tests the address in the form, off the main thread.
func (a *app) runProbe() {
	addr := strings.TrimSpace(a.draft.Address)
	if addr == "" || a.probe.running {
		return
	}
	a.probe = probeState{running: true, host: addr}
	host, port := netprobe.SplitHostPort(addr, netprobe.DefaultPort)
	go func() {
		r := netprobe.Probe(context.Background(), host, port, 5*time.Second)
		a.post(func() {
			if a.probe.host != addr {
				return // the form moved on
			}
			a.probe = probeState{host: addr, lines: r.Lines(), ok: r.OK()}
		})
	}()
}

// post runs fn on the main thread and redraws the main window.
func (a *app) post(fn func()) {
	if a.win != nil {
		a.win.Update(fn)
		return
	}
	if a.postFn != nil {
		a.postFn(fn)
	}
}
