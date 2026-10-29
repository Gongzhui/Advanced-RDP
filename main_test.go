package main

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/egoist/mygo/ui"

	"github.com/Gongzhui/Advanced-RDP/internal/store"
)

// memVault keeps passwords in memory.
type memVault map[string]string

func (m memVault) Get(k string) (string, bool, error) { v, ok := m[k]; return v, ok, nil }
func (m memVault) Set(k, v string) error              { m[k] = v; return nil }
func (m memVault) Delete(k string) error              { delete(m, k); return nil }
func (m memVault) Rename(a, b string) error {
	if v, ok := m[a]; ok {
		m[b] = v
		delete(m, a)
	}
	return nil
}

const v1Hosts = `[
  {"name": "3400", "address": "lab.example.com:3400", "username": "alice", "domain": "", "desktopWidth": 1280, "desktopHeight": 720, "redirectClipboard": true, "fullScreen": true},
  {"name": "3390", "address": "lab.example.com:3390", "username": "alice", "domain": "", "desktopWidth": 1280, "desktopHeight": 720, "redirectClipboard": true, "fullScreen": true},
  {"name": "home", "address": "10.0.0.8", "username": "bob", "domain": "", "desktopWidth": 1600, "desktopHeight": 900, "redirectClipboard": false, "fullScreen": false}
]`

type fixture struct {
	a      *app
	tt     *ui.Tester
	vault  memVault
	opened []string
	dir    string
}

func newFixture(t *testing.T, hosts string) *fixture {
	t.Helper()
	dir := t.TempDir()
	if hosts != "" {
		os.WriteFile(filepath.Join(dir, "hosts.json"), []byte(hosts), 0o600)
	}
	f := &fixture{dir: dir, vault: memVault{"advancedrdp:3400": "pw3400", "advancedrdp:3390": "pw3390"}}
	f.a = newApp(store.Open(dir), f.vault, nil)
	f.a.postFn = func(fn func()) { fn() }
	f.a.openSession = func(h store.Host, pw string) *session {
		f.opened = append(f.opened, h.Name+"/"+pw)
		return &session{app: f.a, host: h, state: stConnecting}
	}
	f.tt = ui.NewTester(f.a.view, 920, 600)
	return f
}

func (f *fixture) click(t *testing.T, label string) {
	t.Helper()
	if err := f.tt.Click(label); err != nil {
		t.Fatalf("click %q: %v (texts %q)", label, err, f.tt.Texts())
	}
}

func (f *fixture) saved(t *testing.T) []store.Host {
	t.Helper()
	hosts, err := store.Open(f.dir).Load()
	if err != nil {
		t.Fatal(err)
	}
	return hosts
}

func TestShowsInheritedHosts(t *testing.T) {
	f := newFixture(t, v1Hosts)
	for _, s := range []string{"3400", "3390", "home", "alice @ lab.example.com:3400", "3 台主机"} {
		if !f.tt.HasText(s) {
			t.Errorf("missing %q in %q", s, f.tt.Texts())
		}
	}
	// The first host is open in the form, with its saved password.
	if f.a.draft.Name != "3400" || f.a.password != "pw3400" {
		t.Errorf("form %+v, password %q", f.a.draft, f.a.password)
	}
}

func TestConnectUsesSavedPassword(t *testing.T) {
	f := newFixture(t, v1Hosts)
	f.click(t, "连接")
	if len(f.opened) != 1 || f.opened[0] != "3400/pw3400" {
		t.Fatalf("opened %v", f.opened)
	}
	// Connecting again brings the open session forward instead.
	f.tt.Frame()
	f.click(t, "切换到会话")
	if len(f.opened) != 1 {
		t.Errorf("opened twice: %v", f.opened)
	}
}

func TestChooseAnotherHost(t *testing.T) {
	f := newFixture(t, v1Hosts)
	f.click(t, "3390")
	if f.a.draft.Name != "3390" || f.a.password != "pw3390" {
		t.Fatalf("form %+v, password %q", f.a.draft, f.a.password)
	}
	// Ctrl+Enter connects.
	f.tt.Key(ui.Cmd, ui.KeyEnter)
	if len(f.opened) != 1 || f.opened[0] != "3390/pw3390" {
		t.Fatalf("opened %v", f.opened)
	}
}

func TestSearchFilters(t *testing.T) {
	f := newFixture(t, v1Hosts)
	f.tt.Key(ui.Cmd, ui.KeyF)
	f.tt.Type("bob")
	f.tt.Frame()
	if f.tt.HasText("3390") || !f.tt.HasText("home") {
		t.Errorf("texts %q", f.tt.Texts())
	}
	f.tt.Type("zzz")
	f.tt.Frame()
	if !f.tt.HasText("没有匹配的主机。") {
		t.Errorf("texts %q", f.tt.Texts())
	}
}

func TestNewHostSavesToFileAndVault(t *testing.T) {
	f := newFixture(t, v1Hosts)
	f.tt.Key(ui.Cmd, ui.KeyN)
	f.tt.Frame()
	f.tt.Type("NAS")
	f.a.draft.Address = "nas.local:3391"
	f.a.draft.Username = "carol"
	f.a.password = "secret"
	f.tt.Frame()
	f.tt.Key(ui.Cmd, ui.KeyS)
	hosts := f.saved(t)
	if len(hosts) != 4 || hosts[3].Name != "NAS" || hosts[3].Address != "nas.local:3391" {
		t.Fatalf("saved %+v", hosts)
	}
	if f.vault["advancedrdp:NAS"] != "secret" {
		t.Errorf("vault %v", f.vault)
	}
	// A backup of the first version's file is kept.
	if b, err := os.ReadFile(filepath.Join(f.dir, "hosts.v1-backup.json")); err != nil || string(b) != v1Hosts {
		t.Errorf("backup %q, %v", b, err)
	}
}

func TestRenameMovesPassword(t *testing.T) {
	f := newFixture(t, v1Hosts)
	f.a.draft.Name = "Lab 3400"
	f.tt.Frame()
	f.tt.Key(ui.Cmd, ui.KeyS)
	if _, ok := f.vault["advancedrdp:3400"]; ok || f.vault["advancedrdp:Lab 3400"] != "pw3400" {
		t.Errorf("vault %v", f.vault)
	}
	if h := f.saved(t); h[0].Name != "Lab 3400" {
		t.Errorf("saved %+v", h[0])
	}
}

func TestDuplicateNameRefused(t *testing.T) {
	f := newFixture(t, v1Hosts)
	f.a.draft.Name = "home"
	f.tt.Frame()
	f.tt.Key(ui.Cmd, ui.KeyS)
	if h := f.saved(t); h[0].Name != "3400" {
		t.Errorf("saved a duplicate: %+v", h)
	}
	f.tt.Frame()
	if !f.tt.HasText("已有名为 “home” 的主机") {
		t.Errorf("no message in %q", f.tt.Texts())
	}
}

func TestSwitchingSavesEdits(t *testing.T) {
	f := newFixture(t, v1Hosts)
	f.a.draft.Username = "alice2"
	f.tt.Frame()
	f.click(t, "home")
	if h := f.saved(t); h[0].Username != "alice2" {
		t.Errorf("edit not saved: %+v", h[0])
	}
	if f.a.draft.Name != "home" {
		t.Errorf("form shows %q", f.a.draft.Name)
	}
}

func TestForgetPassword(t *testing.T) {
	f := newFixture(t, v1Hosts)
	f.click(t, "记住密码")
	f.tt.Key(ui.Cmd, ui.KeyS)
	if _, ok := f.vault["advancedrdp:3400"]; ok {
		t.Errorf("password kept: %v", f.vault)
	}
}

func TestDeleteAsksFirst(t *testing.T) {
	f := newFixture(t, v1Hosts)
	f.click(t, "删除")
	f.tt.Frame()
	if len(f.saved(t)) != 3 {
		t.Fatal("deleted without asking")
	}
	if !f.tt.HasText("删除 “3400”？") {
		t.Fatalf("no alert in %q", f.tt.Texts())
	}
	f.tt.Key(0, ui.KeyEnter) // the alert's default button: 删除
	f.tt.Frame()
	hosts := f.saved(t)
	if len(hosts) != 2 || hosts[0].Name != "3390" {
		t.Fatalf("saved %+v", hosts)
	}
	if _, ok := f.vault["advancedrdp:3400"]; ok {
		t.Error("password of the deleted host kept")
	}
	if f.a.draft.Name != "3390" {
		t.Errorf("form shows %q", f.a.draft.Name)
	}
}

func TestWindowedShowsResolution(t *testing.T) {
	f := newFixture(t, v1Hosts)
	f.click(t, "home")
	f.tt.Frame()
	if !f.tt.HasText("1600 × 900") {
		t.Errorf("texts %q", f.tt.Texts())
	}
}

func TestEmptyStart(t *testing.T) {
	f := newFixture(t, "")
	if !f.tt.HasText("新主机") || !f.tt.HasText("还没有主机。点 “+ 新建” 添加。") {
		t.Errorf("texts %q", f.tt.Texts())
	}
	f.click(t, "连接")
	if len(f.opened) != 0 || len(f.saved(t)) != 0 {
		t.Error("connected an empty host")
	}
}

func TestParseResolution(t *testing.T) {
	if w, h, ok := parseResolution(resolutionLabel(1920, 1080)); !ok || w != 1920 || h != 1080 {
		t.Errorf("%d %d %v", w, h, ok)
	}
	if _, _, ok := parseResolution("x"); ok {
		t.Error("parsed junk")
	}
}

func TestEnterInPasswordConnects(t *testing.T) {
	f := newFixture(t, v1Hosts)
	f.click(t, "home")
	f.tt.Frame()
	f.click(t, "密码") // the field's label focuses it
	f.tt.Type("pw-home")
	f.tt.Key(0, ui.KeyEnter)
	f.tt.Frame()
	if len(f.opened) != 1 || f.opened[0] != "home/pw-home" {
		t.Fatalf("opened %v, focused password: %v", f.opened, f.tt.Focused("密码"))
	}
}
