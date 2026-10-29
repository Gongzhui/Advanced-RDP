package cred

import (
	"bytes"
	"encoding/base64"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"testing"
	"time"
)

// testVault uses an application name of its own, so the tests never touch
// the passwords of Advanced RDP.
func testVault(t *testing.T) *Vault {
	t.Helper()
	app := fmt.Sprintf("AdvancedRdpTest%d", time.Now().UnixNano())
	return New(app, t.TempDir())
}

func TestSetGetRenameDelete(t *testing.T) {
	v := testVault(t)
	const pw = "p@ss 密码 🔑"
	t.Cleanup(func() { v.Delete("advancedrdp:a"); v.Delete("advancedrdp:b") })

	if _, ok, err := v.Get("advancedrdp:a"); ok || err != nil {
		t.Fatalf("missing key: ok %v, err %v", ok, err)
	}
	if err := v.Set("advancedrdp:a", pw); err != nil {
		t.Fatal(err)
	}
	if got, ok, err := v.Get("advancedrdp:a"); !ok || err != nil || got != pw {
		t.Fatalf("Get = %q, %v, %v", got, ok, err)
	}
	if err := v.Rename("advancedrdp:a", "advancedrdp:b"); err != nil {
		t.Fatal(err)
	}
	if _, ok, _ := v.Get("advancedrdp:a"); ok {
		t.Error("old key kept after rename")
	}
	if got, ok, _ := v.Get("advancedrdp:b"); !ok || got != pw {
		t.Errorf("renamed key = %q, %v", got, ok)
	}
	if err := v.Delete("advancedrdp:b"); err != nil {
		t.Fatal(err)
	}
	if _, ok, _ := v.Get("advancedrdp:b"); ok {
		t.Error("key kept after delete")
	}
	if err := v.Delete("advancedrdp:b"); err != nil {
		t.Errorf("deleting a missing key: %v", err)
	}
}

// A password only in the first version's secrets.json is read and moved
// into the Credential Manager.
func TestLegacySecretsFile(t *testing.T) {
	v := testVault(t)
	t.Cleanup(func() { v.Delete("advancedrdp:old") })
	cipher, err := protect([]byte("legacy-pw"))
	if err != nil {
		t.Fatal(err)
	}
	data, _ := json.Marshal(map[string]string{"advancedrdp:old": base64.StdEncoding.EncodeToString(cipher)})
	if err := os.WriteFile(v.secrets, data, 0o600); err != nil {
		t.Fatal(err)
	}
	if got, ok, err := v.Get("advancedrdp:old"); !ok || err != nil || got != "legacy-pw" {
		t.Fatalf("Get = %q, %v, %v", got, ok, err)
	}
	if got, ok, _ := credRead(v.target("advancedrdp:old")); !ok || got != "legacy-pw" {
		t.Errorf("not migrated: %q, %v", got, ok)
	}
	v.Delete("advancedrdp:old")
	if m := v.loadLegacy(); len(m) != 0 {
		t.Errorf("secrets.json still has %v", m)
	}
}

func TestBlobRoundTrip(t *testing.T) {
	for _, s := range []string{"", "a", "密码", "🔑x"} {
		if got := decodeBlob(encodeBlob(s)); got != s {
			t.Errorf("%q -> %q", s, got)
		}
	}
	// A blob with a terminating NUL, as some writers store.
	if got := decodeBlob(append(encodeBlob("ab"), 0, 0)); got != "ab" {
		t.Errorf("got %q", got)
	}
}

// TestRealVault checks, read only, that the passwords of the installed
// application are readable. Run with ARDP_REAL_VAULT=1.
func TestRealVault(t *testing.T) {
	if os.Getenv("ARDP_REAL_VAULT") == "" {
		t.Skip("set ARDP_REAL_VAULT=1")
	}
	dir := filepath.Join(os.Getenv("LOCALAPPDATA"), "AdvancedRdp")
	data, err := os.ReadFile(filepath.Join(dir, "hosts.json"))
	if err != nil {
		t.Skip(err)
	}
	var hosts []struct{ Name string }
	json.Unmarshal(bytes.TrimPrefix(data, []byte("\xef\xbb\xbf")), &hosts)
	v := New("AdvancedRdp", dir)
	for _, h := range hosts {
		_, ok, err := credRead(v.target("advancedrdp:" + h.Name))
		_, legacy := v.readLegacy("advancedrdp:" + h.Name)
		t.Logf("%s: credential manager %v (%v), secrets.json %v", h.Name, ok, err, legacy)
		if !ok && !legacy {
			t.Errorf("%s: no password", h.Name)
		}
	}
}
