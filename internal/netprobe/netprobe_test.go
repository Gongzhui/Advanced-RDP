package netprobe

import (
	"context"
	"io"
	"net"
	"strings"
	"testing"
	"time"
)

func TestSplitHostPort(t *testing.T) {
	cases := []struct {
		in   string
		host string
		port int
	}{
		{"3dp.example.com:3400", "3dp.example.com", 3400},
		{"  10.0.0.5  ", "10.0.0.5", DefaultPort},
		{"host:", "host:", DefaultPort},
		{"host:99999", "host:99999", DefaultPort},
		{"[fe80::1]:3390", "fe80::1", 3390},
		{"[fe80::1]", "fe80::1", DefaultPort},
		{"fe80::1", "fe80::1", DefaultPort},
		{"::1", "::1", DefaultPort},
		{"", "", DefaultPort},
	}
	for _, c := range cases {
		h, p := SplitHostPort(c.in, DefaultPort)
		if h != c.host || p != c.port {
			t.Errorf("SplitHostPort(%q) = %q, %d; want %q, %d", c.in, h, p, c.host, c.port)
		}
	}
}

// fakeServer answers the connection request with reply.
func fakeServer(t *testing.T, reply []byte) (string, int) {
	t.Helper()
	ln, err := net.Listen("tcp", "127.0.0.1:0")
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { ln.Close() })
	go func() {
		c, err := ln.Accept()
		if err != nil {
			return
		}
		defer c.Close()
		req := make([]byte, len(x224Request))
		io.ReadFull(c, req)
		c.Write(reply)
	}()
	a := ln.Addr().(*net.TCPAddr)
	return "127.0.0.1", a.Port
}

func TestProbeNLA(t *testing.T) {
	host, port := fakeServer(t, []byte{
		0x03, 0x00, 0x00, 0x13,
		0x0e, 0xd0, 0x00, 0x00, 0x12, 0x34, 0x00,
		0x02, 0x1f, 0x08, 0x00, 0x02, 0x00, 0x00, 0x00,
	})
	r := Probe(context.Background(), host, port, 2*time.Second)
	if !r.OK() || r.Protocol != "NLA" {
		t.Fatalf("result %+v", r)
	}
	if lines := strings.Join(r.Lines(), "\n"); !strings.Contains(lines, "NLA") || !strings.Contains(lines, "TCP") {
		t.Errorf("lines %q", lines)
	}
}

func TestProbeNotRDP(t *testing.T) {
	host, port := fakeServer(t, []byte("HTTP/1.1 400 Bad Request\r\n\r\n"))
	r := Probe(context.Background(), host, port, 2*time.Second)
	if r.OK() || r.Stage != "握手" {
		t.Fatalf("result %+v", r)
	}
}

func TestProbeClosedPort(t *testing.T) {
	ln, _ := net.Listen("tcp", "127.0.0.1:0")
	port := ln.Addr().(*net.TCPAddr).Port
	ln.Close()
	r := Probe(context.Background(), "127.0.0.1", port, 2*time.Second)
	if r.OK() || r.Stage != "连接" {
		t.Fatalf("result %+v", r)
	}
	if l := strings.Join(r.Lines(), " "); !strings.Contains(l, "没有服务在运行") {
		t.Errorf("lines %q", l)
	}
}

// A name that does not exist fails, at resolution or, behind a proxy that
// resolves every name to a fake address, at a later stage.
func TestProbeUnresolvable(t *testing.T) {
	r := Probe(context.Background(), "no-such-host.invalid", 3389, 2*time.Second)
	if r.OK() || r.Stage == "" {
		t.Fatalf("result %+v", r)
	}
}

func TestProbeLegacySecurityServer(t *testing.T) {
	host, port := fakeServer(t, []byte{
		0x03, 0x00, 0x00, 0x13,
		0x0e, 0xd0, 0x00, 0x00, 0x12, 0x34, 0x00,
		0x03, 0x00, 0x08, 0x00, 0x02, 0x00, 0x00, 0x00, // SSL_NOT_ALLOWED_BY_SERVER
	})
	if r := Probe(context.Background(), host, port, 2*time.Second); !r.OK() || r.Protocol != "RDP" {
		t.Fatalf("result %+v", r)
	}
}

// TestProbeLocalRDP probes this computer's own RDP service, when it runs.
func TestProbeLocalRDP(t *testing.T) {
	c, err := net.DialTimeout("tcp", "127.0.0.1:3389", time.Second)
	if err != nil {
		t.Skip("no local RDP service")
	}
	c.Close()
	r := Probe(context.Background(), "127.0.0.1", 3389, 3*time.Second)
	if !r.OK() {
		t.Fatalf("result %+v", r)
	}
	t.Log(r.Lines())
}
