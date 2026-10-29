// Package cred keeps passwords in the Windows Credential Manager, under the
// targets the first version of Advanced RDP used ("AdvancedRdp/<key>"), and
// reads the DPAPI-encrypted secrets.json that version also wrote.
package cred

import (
	"encoding/base64"
	"encoding/json"
	"errors"
	"os"
	"os/user"
	"path/filepath"
	"syscall"
	"unicode/utf16"
	"unsafe"

	"golang.org/x/sys/windows"
)

var (
	advapi32       = windows.NewLazySystemDLL("advapi32.dll")
	procCredReadW  = advapi32.NewProc("CredReadW")
	procCredWriteW = advapi32.NewProc("CredWriteW")
	procCredDelete = advapi32.NewProc("CredDeleteW")
	procCredFree   = advapi32.NewProc("CredFree")
)

const (
	credTypeGeneric         = 1
	credPersistLocalMachine = 2
	errNotFound             = syscall.Errno(1168) // ERROR_NOT_FOUND
)

// credential is CREDENTIALW.
type credential struct {
	Flags              uint32
	Type               uint32
	TargetName         *uint16
	Comment            *uint16
	LastWritten        windows.Filetime
	CredentialBlobSize uint32
	CredentialBlob     *byte
	Persist            uint32
	AttributeCount     uint32
	Attributes         uintptr
	TargetAlias        *uint16
	UserName           *uint16
}

// Vault stores passwords for one application.
type Vault struct {
	app     string
	secrets string // secrets.json of the first version
}

// New returns the vault of app, whose legacy secrets.json is in dir.
func New(app, dir string) *Vault {
	return &Vault{app: app, secrets: filepath.Join(dir, "secrets.json")}
}

func (v *Vault) target(key string) string { return v.app + "/" + key }

// Get returns the password of key, and whether there is one. A password the
// first version kept only in secrets.json moves to the Credential Manager.
func (v *Vault) Get(key string) (string, bool, error) {
	pw, ok, err := credRead(v.target(key))
	if ok || err != nil {
		return pw, ok, err
	}
	pw, ok = v.readLegacy(key)
	if ok {
		_ = credWrite(v.target(key), pw)
	}
	return pw, ok, nil
}

// Set stores the password of key.
func (v *Vault) Set(key, password string) error {
	return credWrite(v.target(key), password)
}

// Delete removes the password of key, from secrets.json too.
func (v *Vault) Delete(key string) error {
	err := credDelete(v.target(key))
	v.deleteLegacy(key)
	return err
}

// Rename moves the password of from to to, for a renamed host.
func (v *Vault) Rename(from, to string) error {
	if from == to {
		return nil
	}
	pw, ok, err := v.Get(from)
	if err != nil || !ok {
		return err
	}
	if err := v.Set(to, pw); err != nil {
		return err
	}
	return v.Delete(from)
}

func credRead(target string) (string, bool, error) {
	t, err := windows.UTF16PtrFromString(target)
	if err != nil {
		return "", false, err
	}
	var p *credential
	r, _, e := procCredReadW.Call(uintptr(unsafe.Pointer(t)), credTypeGeneric, 0, uintptr(unsafe.Pointer(&p)))
	if r == 0 {
		if errors.Is(e, errNotFound) {
			return "", false, nil
		}
		return "", false, e
	}
	defer procCredFree.Call(uintptr(unsafe.Pointer(p)))
	n := int(p.CredentialBlobSize)
	if n == 0 || p.CredentialBlob == nil {
		return "", false, nil
	}
	return decodeBlob(unsafe.Slice(p.CredentialBlob, n)), true, nil
}

// decodeBlob reads the UTF-16LE password the first version stored, without
// a terminating NUL.
func decodeBlob(b []byte) string {
	u := make([]uint16, len(b)/2)
	for i := range u {
		u[i] = uint16(b[2*i]) | uint16(b[2*i+1])<<8
	}
	for len(u) > 0 && u[len(u)-1] == 0 {
		u = u[:len(u)-1]
	}
	return string(utf16.Decode(u))
}

func encodeBlob(s string) []byte {
	u := utf16.Encode([]rune(s))
	b := make([]byte, 2*len(u))
	for i, c := range u {
		b[2*i], b[2*i+1] = byte(c), byte(c>>8)
	}
	return b
}

func credWrite(target, password string) error {
	t, err := windows.UTF16PtrFromString(target)
	if err != nil {
		return err
	}
	name := "user"
	if u, err := user.Current(); err == nil {
		name = filepath.Base(u.Username)
	}
	un, _ := windows.UTF16PtrFromString(name)
	blob := encodeBlob(password)
	c := credential{
		Type:               credTypeGeneric,
		TargetName:         t,
		Persist:            credPersistLocalMachine,
		UserName:           un,
		CredentialBlobSize: uint32(len(blob)),
	}
	if len(blob) > 0 {
		c.CredentialBlob = &blob[0]
	}
	r, _, e := procCredWriteW.Call(uintptr(unsafe.Pointer(&c)), 0)
	clear(blob)
	if r == 0 {
		return e
	}
	return nil
}

func credDelete(target string) error {
	t, err := windows.UTF16PtrFromString(target)
	if err != nil {
		return err
	}
	r, _, e := procCredDelete.Call(uintptr(unsafe.Pointer(t)), credTypeGeneric, 0)
	if r == 0 && !errors.Is(e, errNotFound) {
		return e
	}
	return nil
}

func (v *Vault) loadLegacy() map[string]string {
	data, err := os.ReadFile(v.secrets)
	if err != nil {
		return nil
	}
	m := map[string]string{}
	if json.Unmarshal(data, &m) != nil {
		return nil
	}
	return m
}

func (v *Vault) readLegacy(key string) (string, bool) {
	b64, ok := v.loadLegacy()[key]
	if !ok {
		return "", false
	}
	cipher, err := base64.StdEncoding.DecodeString(b64)
	if err != nil || len(cipher) == 0 {
		return "", false
	}
	plain, err := unprotect(cipher)
	if err != nil {
		return "", false
	}
	return string(plain), true
}

func (v *Vault) deleteLegacy(key string) {
	m := v.loadLegacy()
	if _, ok := m[key]; !ok {
		return
	}
	delete(m, key)
	if data, err := json.Marshal(m); err == nil {
		os.WriteFile(v.secrets, data, 0o600)
	}
}

// unprotect decrypts DPAPI data of the current user.
func unprotect(cipher []byte) ([]byte, error) {
	in := windows.DataBlob{Size: uint32(len(cipher)), Data: &cipher[0]}
	var out windows.DataBlob
	if err := windows.CryptUnprotectData(&in, nil, nil, 0, nil, windows.CRYPTPROTECT_UI_FORBIDDEN, &out); err != nil {
		return nil, err
	}
	defer windows.LocalFree(windows.Handle(unsafe.Pointer(out.Data)))
	return append([]byte(nil), unsafe.Slice(out.Data, out.Size)...), nil
}

// protect encrypts data for the current user, as the first version did.
func protect(plain []byte) ([]byte, error) {
	in := windows.DataBlob{Size: uint32(len(plain))}
	if len(plain) > 0 {
		in.Data = &plain[0]
	}
	var out windows.DataBlob
	if err := windows.CryptProtectData(&in, nil, nil, 0, nil, windows.CRYPTPROTECT_UI_FORBIDDEN, &out); err != nil {
		return nil, err
	}
	defer windows.LocalFree(windows.Handle(unsafe.Pointer(out.Data)))
	return append([]byte(nil), unsafe.Slice(out.Data, out.Size)...), nil
}
