package store

import (
	"encoding/json"
	"os"
	"path/filepath"
	"testing"
	"time"
)

// A file as the .NET version wrote it: camelCase, indented, with a BOM.
const v1File = "\xef\xbb\xbf" + `[
  {
    "name": "3400",
    "address": "3dp.example.com:3400",
    "username": "alice",
    "domain": "",
    "desktopWidth": 1280,
    "desktopHeight": 720,
    "redirectClipboard": true,
    "fullScreen": true
  },
  {
    "name": " lab ",
    "address": "10.0.0.5",
    "username": "bob",
    "domain": "LAB",
    "desktopWidth": 0,
    "desktopHeight": 0,
    "redirectClipboard": false,
    "fullScreen": false
  }
]`

func TestLoadV1File(t *testing.T) {
	dir := t.TempDir()
	if err := os.WriteFile(filepath.Join(dir, "hosts.json"), []byte(v1File), 0o600); err != nil {
		t.Fatal(err)
	}
	hosts, err := Open(dir).Load()
	if err != nil {
		t.Fatal(err)
	}
	if len(hosts) != 2 {
		t.Fatalf("got %d hosts", len(hosts))
	}
	if h := hosts[0]; h.Name != "3400" || h.Address != "3dp.example.com:3400" || h.Username != "alice" || !h.FullScreen || !h.RedirectClipboard {
		t.Errorf("first host %+v", h)
	}
	if h := hosts[1]; h.Name != "lab" || h.Domain != "LAB" || h.DesktopWidth != DefaultWidth || h.DesktopHeight != DefaultHeight || h.FullScreen {
		t.Errorf("second host %+v", h)
	}
}

func TestLoadMissing(t *testing.T) {
	hosts, err := Open(t.TempDir()).Load()
	if err != nil || hosts != nil {
		t.Fatalf("hosts %v, err %v", hosts, err)
	}
}

func TestSaveBacksUpOnceAndStaysCompatible(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "hosts.json")
	os.WriteFile(path, []byte(v1File), 0o600)
	s := Open(dir)
	hosts, _ := s.Load()
	hosts[0].LastConnected = time.Date(2026, 10, 29, 12, 0, 0, 0, time.UTC)
	if err := s.Save(hosts); err != nil {
		t.Fatal(err)
	}
	backup, err := os.ReadFile(s.BackupPath())
	if err != nil || string(backup) != v1File {
		t.Fatalf("backup %q, err %v", backup, err)
	}
	// A second save keeps the first backup.
	if err := s.Save(hosts[:1]); err != nil {
		t.Fatal(err)
	}
	if b, _ := os.ReadFile(s.BackupPath()); string(b) != v1File {
		t.Error("backup replaced")
	}
	// The keys stay those of the first version, which reads them back.
	var raw []map[string]any
	data, _ := os.ReadFile(path)
	if err := json.Unmarshal(data, &raw); err != nil {
		t.Fatal(err)
	}
	for _, k := range []string{"name", "address", "username", "domain", "desktopWidth", "desktopHeight", "redirectClipboard", "fullScreen", "lastConnected"} {
		if _, ok := raw[0][k]; !ok {
			t.Errorf("key %q missing in %s", k, data)
		}
	}
	again, _ := s.Load()
	if len(again) != 1 || !again[0].LastConnected.Equal(hosts[0].LastConnected) {
		t.Errorf("reloaded %+v", again)
	}
	// No temporary files are left behind.
	entries, _ := os.ReadDir(dir)
	if len(entries) != 2 {
		t.Errorf("files %v", entries)
	}
}

func TestSaveEmptyWritesArray(t *testing.T) {
	dir := t.TempDir()
	if err := Open(dir).Save(nil); err != nil {
		t.Fatal(err)
	}
	if b, _ := os.ReadFile(filepath.Join(dir, "hosts.json")); string(b) != "[]" {
		t.Errorf("wrote %q", b)
	}
}

func TestValidate(t *testing.T) {
	hosts := []Host{{Name: "Lab", Address: "a"}, {Name: "Home", Address: "b"}}
	cases := []struct {
		h     Host
		index int
		ok    bool
	}{
		{Host{Name: "New", Address: "c"}, -1, true},
		{Host{Name: "lab", Address: "c"}, -1, false},
		{Host{Name: "Lab", Address: "c"}, 0, true},
		{Host{Name: "Home", Address: "c"}, 0, false},
		{Host{Name: " ", Address: "c"}, -1, false},
		{Host{Name: "x", Address: ""}, -1, false},
	}
	for _, tc := range cases {
		if err := Validate(tc.h, hosts, tc.index); (err == nil) != tc.ok {
			t.Errorf("Validate(%+v, %d) = %v", tc.h, tc.index, err)
		}
	}
}
