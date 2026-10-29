package main

import (
	"fmt"
	"os"
	"path/filepath"
	"sync"
	"time"

	"github.com/Gongzhui/Advanced-RDP/internal/cred"
)

// newVault returns the vault of the first version, "AdvancedRdp", or of
// app when set, for trying builds.
func newVault(dir, app string) vault {
	if app == "" {
		app = "AdvancedRdp"
	}
	return cred.New(app, dir)
}

// maxLog is the size at which the log starts over, keeping one old file.
const maxLog = 1 << 20

// logger appends lines to one file, which it keeps small.
type logger struct {
	mu   sync.Mutex
	path string
	f    *os.File
	size int64
}

func newLogger(path string) *logger { return &logger{path: path} }

func (l *logger) Printf(format string, args ...any) {
	if l == nil {
		return
	}
	line := time.Now().Format("2006-01-02 15:04:05.000 ") + fmt.Sprintf(format, args...) + "\n"
	l.mu.Lock()
	defer l.mu.Unlock()
	if l.f == nil || l.size > maxLog {
		l.reopen()
	}
	if l.f != nil {
		n, _ := l.f.WriteString(line)
		l.size += int64(n)
	}
}

func (l *logger) reopen() {
	if l.f != nil {
		l.f.Close()
		l.f = nil
		os.Rename(l.path, l.path+".1")
	}
	os.MkdirAll(filepath.Dir(l.path), 0o755)
	f, err := os.OpenFile(l.path, os.O_CREATE|os.O_APPEND|os.O_WRONLY, 0o644)
	if err != nil {
		return
	}
	if st, err := f.Stat(); err == nil {
		l.size = st.Size()
	}
	l.f = f
}

func (l *logger) Close() {
	if l == nil {
		return
	}
	l.mu.Lock()
	defer l.mu.Unlock()
	if l.f != nil {
		l.f.Close()
		l.f = nil
	}
}
