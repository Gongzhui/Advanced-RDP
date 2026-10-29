// Package store keeps the saved hosts in hosts.json, in the format the first
// (.NET) version of Advanced RDP wrote, so both versions read the same file.
package store

import (
	"bytes"
	"encoding/json"
	"errors"
	"fmt"
	"io/fs"
	"os"
	"path/filepath"
	"strings"
	"time"
)

// Default size of a windowed session, as the first version used.
const (
	DefaultWidth  = 1280
	DefaultHeight = 720
)

// Host is one saved connection. The JSON names are those of the first
// version; fields it did not have are omitted when empty, and it ignores them.
type Host struct {
	Name              string    `json:"name"`
	Address           string    `json:"address"`
	Username          string    `json:"username"`
	Domain            string    `json:"domain"`
	DesktopWidth      int       `json:"desktopWidth"`
	DesktopHeight     int       `json:"desktopHeight"`
	RedirectClipboard bool      `json:"redirectClipboard"`
	FullScreen        bool      `json:"fullScreen"`
	LastConnected     time.Time `json:"lastConnected,omitzero"`
}

// NewHost returns a host with the defaults of a new entry.
func NewHost() Host {
	return Host{DesktopWidth: DefaultWidth, DesktopHeight: DefaultHeight, RedirectClipboard: true, FullScreen: true}
}

// CredentialKey is the key of the host's password, as the first version
// named it; the vault prefixes it with the application name.
func CredentialKey(name string) string { return "advancedrdp:" + name }

// Store reads and writes hosts.json in a directory.
type Store struct {
	dir string
}

// DefaultDir is %LOCALAPPDATA%\AdvancedRdp, where the first version kept its data.
func DefaultDir() (string, error) {
	base, err := os.UserCacheDir() // %LOCALAPPDATA% on Windows
	if err != nil {
		return "", err
	}
	return filepath.Join(base, "AdvancedRdp"), nil
}

// Open returns the store of dir, which it creates when it writes.
func Open(dir string) *Store { return &Store{dir: dir} }

// Dir is the store's directory.
func (s *Store) Dir() string { return s.dir }

func (s *Store) path() string { return filepath.Join(s.dir, "hosts.json") }

// BackupPath is where Save keeps the file it first replaces.
func (s *Store) BackupPath() string { return filepath.Join(s.dir, "hosts.v1-backup.json") }

// Load reads the hosts. A missing file is no hosts.
func (s *Store) Load() ([]Host, error) {
	data, err := os.ReadFile(s.path())
	if errors.Is(err, fs.ErrNotExist) {
		return nil, nil
	}
	if err != nil {
		return nil, err
	}
	data = bytes.TrimPrefix(data, []byte("\xef\xbb\xbf")) // .NET may write a BOM
	if len(bytes.TrimSpace(data)) == 0 {
		return nil, nil
	}
	var hosts []Host
	if err := json.Unmarshal(data, &hosts); err != nil {
		return nil, fmt.Errorf("读取 %s 失败: %w", s.path(), err)
	}
	for i := range hosts {
		normalize(&hosts[i])
	}
	return hosts, nil
}

// Save writes the hosts atomically. The first time it replaces a file it
// did not write, it keeps a copy of it at BackupPath.
func (s *Store) Save(hosts []Host) error {
	if err := os.MkdirAll(s.dir, 0o755); err != nil {
		return err
	}
	if _, err := os.Stat(s.BackupPath()); errors.Is(err, fs.ErrNotExist) {
		if old, err := os.ReadFile(s.path()); err == nil {
			if err := os.WriteFile(s.BackupPath(), old, 0o600); err != nil {
				return fmt.Errorf("备份 hosts.json 失败: %w", err)
			}
		}
	}
	if hosts == nil {
		hosts = []Host{}
	}
	data, err := json.MarshalIndent(hosts, "", "  ")
	if err != nil {
		return err
	}
	tmp, err := os.CreateTemp(s.dir, "hosts-*.tmp")
	if err != nil {
		return err
	}
	_, werr := tmp.Write(data)
	cerr := tmp.Close()
	if werr != nil || cerr != nil {
		os.Remove(tmp.Name())
		return errors.Join(werr, cerr)
	}
	if err := os.Rename(tmp.Name(), s.path()); err != nil {
		os.Remove(tmp.Name())
		return err
	}
	return nil
}

func normalize(h *Host) {
	h.Name = strings.TrimSpace(h.Name)
	h.Address = strings.TrimSpace(h.Address)
	if h.DesktopWidth <= 0 || h.DesktopHeight <= 0 {
		h.DesktopWidth, h.DesktopHeight = DefaultWidth, DefaultHeight
	}
}

// Validate checks a host before it is saved among others; index is its
// position in hosts, or -1 for a new one.
func Validate(h Host, hosts []Host, index int) error {
	if strings.TrimSpace(h.Name) == "" {
		return errors.New("名称不能为空")
	}
	if strings.TrimSpace(h.Address) == "" {
		return errors.New("地址不能为空")
	}
	for i, o := range hosts {
		if i != index && strings.EqualFold(strings.TrimSpace(o.Name), strings.TrimSpace(h.Name)) {
			return fmt.Errorf("已有名为 “%s” 的主机", o.Name)
		}
	}
	return nil
}
