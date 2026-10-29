// Advanced RDP is a remote desktop client for Windows: saved hosts with
// their passwords in the Credential Manager, and sessions in windows of
// their own, drawn by the Remote Desktop control of Windows.
package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"slices"
	"strings"
	"time"

	"github.com/egoist/mygo"
	"github.com/egoist/mygo/ui"

	"github.com/Gongzhui/Advanced-RDP/internal/store"
)

// vault keeps passwords; cred.Vault on Windows, a map in tests.
type vault interface {
	Get(key string) (string, bool, error)
	Set(key, password string) error
	Delete(key string) error
	Rename(from, to string) error
}

// app is the state of the main window.
type app struct {
	store *store.Store
	vault vault
	log   *logger
	win   *mygo.Window

	hosts []store.Host
	// sel is the host being edited, an index in hosts, or -1 for a new one.
	sel     int
	listSel int   // the row chosen in the list, an index in visible
	visible []int // the hosts the search shows, indices in hosts
	list    ui.ListState
	query   string

	// The form: a copy of the host being edited, and its password.
	draft        store.Host
	password     string
	savedPw      string // the password in the vault, to tell edits
	savePassword bool
	showPassword bool
	nameField    ui.Handle
	searchField  ui.Handle

	confirmDelete bool
	loadErr       string

	probe probeState

	sessions map[string]*session // open sessions, by host name

	// postFn runs a function on the main thread when there is no window,
	// as in tests.
	postFn func(func())
	// openSession opens a session window; tests replace it.
	openSession func(h store.Host, password string) *session
}

type probeState struct {
	running bool
	host    string
	lines   []string
	ok      bool
}

func newApp(s *store.Store, v vault, l *logger) *app {
	a := &app{store: s, vault: v, log: l, sel: -1, sessions: map[string]*session{}}
	hosts, err := s.Load()
	if err != nil {
		a.loadErr = err.Error()
		l.Printf("load hosts: %v", err)
	}
	a.hosts = hosts
	a.list.Selected = &a.listSel
	a.list.Key = func(i int) any { return a.hosts[a.visible[i]].Name }
	a.list.Label = func(i int) string { return a.hosts[a.visible[i]].Name }
	a.refilter()
	if len(a.hosts) > 0 {
		a.edit(0)
	} else {
		a.newHost()
	}
	return a
}

// refilter computes the hosts the search shows, keeping the chosen row on
// the host being edited.
func (a *app) refilter() {
	q := strings.ToLower(strings.TrimSpace(a.query))
	a.visible = a.visible[:0]
	for i, h := range a.hosts {
		if q == "" || strings.Contains(strings.ToLower(h.Name), q) ||
			strings.Contains(strings.ToLower(h.Address), q) ||
			strings.Contains(strings.ToLower(h.Username), q) {
			a.visible = append(a.visible, i)
		}
	}
	a.listSel = slices.Index(a.visible, a.sel)
}

// edit loads host i into the form.
func (a *app) edit(i int) {
	a.sel = i
	a.draft = a.hosts[i]
	a.savedPw = ""
	if pw, ok, err := a.vault.Get(store.CredentialKey(a.draft.Name)); err == nil && ok {
		a.savedPw = pw
	}
	a.password = a.savedPw
	a.savePassword = true
	a.showPassword = false
	a.confirmDelete = false
	a.listSel = slices.Index(a.visible, i)
	if a.probe.host != a.draft.Address {
		a.probe = probeState{}
	}
}

// newHost starts a new host in the form.
func (a *app) newHost() {
	a.sel = -1
	a.draft = store.NewHost()
	a.password, a.savedPw = "", ""
	a.savePassword = true
	a.showPassword = false
	a.confirmDelete = false
	a.listSel = -1
	a.probe = probeState{}
}

// dirty reports whether the form differs from what is saved.
func (a *app) dirty() bool {
	if a.sel < 0 {
		return a.draft.Name != "" || a.draft.Address != "" || a.draft.Username != "" || a.password != ""
	}
	saved := a.hosts[a.sel]
	d := a.draft
	d.LastConnected = saved.LastConnected
	return d != saved || a.password != a.savedPw || (!a.savePassword && a.savedPw != "")
}

// save stores the form. It returns an error for the user when the form is
// not valid or cannot be written.
func (a *app) save() error {
	h := a.draft
	h.Name = strings.TrimSpace(h.Name)
	h.Address = strings.TrimSpace(h.Address)
	h.Username = strings.TrimSpace(h.Username)
	h.Domain = strings.TrimSpace(h.Domain)
	if err := store.Validate(h, a.hosts, a.sel); err != nil {
		return err
	}
	hosts := slices.Clone(a.hosts)
	oldName := ""
	if a.sel < 0 {
		hosts = append(hosts, h)
	} else {
		oldName = hosts[a.sel].Name
		hosts[a.sel] = h
	}
	if err := a.store.Save(hosts); err != nil {
		a.log.Printf("save hosts: %v", err)
		return fmt.Errorf("保存失败: %v", err)
	}
	if oldName != "" && oldName != h.Name {
		if err := a.vault.Rename(store.CredentialKey(oldName), store.CredentialKey(h.Name)); err != nil {
			a.log.Printf("rename credential: %v", err)
		}
		if s := a.sessions[oldName]; s != nil {
			delete(a.sessions, oldName)
			a.sessions[h.Name] = s
		}
	}
	key := store.CredentialKey(h.Name)
	switch {
	case a.savePassword && a.password != "":
		if err := a.vault.Set(key, a.password); err != nil {
			a.log.Printf("save password: %v", err)
			return fmt.Errorf("主机已保存，但密码保存失败: %v", err)
		}
		a.savedPw = a.password
	case !a.savePassword || a.password == "":
		if a.savedPw != "" {
			if err := a.vault.Delete(key); err != nil {
				a.log.Printf("delete password: %v", err)
			}
		}
		a.savedPw = ""
	}
	a.hosts = hosts
	if a.sel < 0 {
		a.sel = len(hosts) - 1
	}
	a.draft = h
	a.refilter()
	return nil
}

// choose switches the form to host i, saving the edits of the host before.
func (a *app) choose(c *ui.Context, i int) {
	if i == a.sel {
		return
	}
	if a.dirty() {
		if err := a.save(); err != nil {
			c.Toast(err.Error())
			a.listSel = slices.Index(a.visible, a.sel)
			return
		}
	}
	a.edit(i)
}

// remove deletes the host being edited and its password.
func (a *app) remove() error {
	if a.sel < 0 {
		a.newHost()
		return nil
	}
	name := a.hosts[a.sel].Name
	hosts := slices.Delete(slices.Clone(a.hosts), a.sel, a.sel+1)
	if err := a.store.Save(hosts); err != nil {
		return fmt.Errorf("删除失败: %v", err)
	}
	if err := a.vault.Delete(store.CredentialKey(name)); err != nil {
		a.log.Printf("delete password: %v", err)
	}
	a.hosts = hosts
	next := min(a.sel, len(hosts)-1)
	a.refilter()
	if next >= 0 {
		a.edit(next)
	} else {
		a.newHost()
	}
	return nil
}

// connect saves the form and opens its host's session, or brings the open
// one to the front.
func (a *app) connect(c *ui.Context) {
	if a.dirty() || a.sel < 0 {
		if err := a.save(); err != nil {
			c.Toast(err.Error())
			return
		}
	}
	h := a.hosts[a.sel]
	if s := a.sessions[h.Name]; s != nil && !s.closed {
		s.front()
		return
	}
	if a.password == "" {
		c.Toast("请输入密码")
		return
	}
	if a.openSession != nil {
		a.sessions[h.Name] = a.openSession(h, a.password)
	}
}

// connected records that host name signed in, for the list's order of use.
func (a *app) connected(name string) {
	i := slices.IndexFunc(a.hosts, func(h store.Host) bool { return h.Name == name })
	if i < 0 {
		return
	}
	hosts := slices.Clone(a.hosts)
	hosts[i].LastConnected = time.Now()
	if err := a.store.Save(hosts); err != nil {
		a.log.Printf("save hosts: %v", err)
		return
	}
	a.hosts = hosts
	if a.sel == i {
		a.draft.LastConnected = hosts[i].LastConnected
	}
}

func main() {
	mygo.App.SetName("Advanced RDP")
	if !mygo.App.RequestSingleInstanceLock() {
		return
	}
	dir, err := store.DefaultDir()
	if err != nil {
		log.Fatal(err)
	}
	if d := os.Getenv("ARDP_DATA_DIR"); d != "" { // for trying builds without the real data
		dir = d
	}
	l := newLogger(filepath.Join(dir, "logs", "advancedrdp.log"))
	defer l.Close()
	a := newApp(store.Open(dir), newVault(dir, os.Getenv("ARDP_VAULT_APP")), l)
	a.openSession = func(h store.Host, password string) *session { return openSession(a, h, password) }

	mygo.App.OnSecondInstance(func(args []string, workingDir string) {
		if a.win != nil {
			a.win.Restore()
			a.win.Show()
			a.win.Focus()
		}
	})
	mygo.App.WhenReady(func() {
		a.win = mygo.NewWindow(mygo.WindowOptions{
			Title:     "Advanced RDP",
			Width:     920,
			Height:    600,
			MinWidth:  680,
			MinHeight: 460,
			StateKey:  "main",
			Content:   ui.View(a.view),
		})
	})
	if err := mygo.App.Run(); err != nil {
		l.Printf("run: %v", err)
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}
