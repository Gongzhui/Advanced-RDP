// Package netprobe parses host addresses and checks that a host answers
// as an RDP server: name resolution, a TCP connection and the first
// message of the RDP handshake.
package netprobe

import (
	"context"
	"encoding/binary"
	"errors"
	"fmt"
	"io"
	"net"
	"os"
	"strconv"
	"strings"
	"syscall"
	"time"
)

// DefaultPort is the RDP port.
const DefaultPort = 3389

// SplitHostPort splits "host", "host:port", "[v6]" or "[v6]:port". A bare
// IPv6 address keeps the default port.
func SplitHostPort(addr string, def int) (string, int) {
	a := strings.TrimSpace(addr)
	if a == "" {
		return "", def
	}
	if strings.HasPrefix(a, "[") {
		end := strings.IndexByte(a, ']')
		if end < 0 {
			return a, def
		}
		host, rest := a[1:end], a[end+1:]
		if p, ok := parsePort(strings.TrimPrefix(rest, ":")); ok && strings.HasPrefix(rest, ":") {
			return host, p
		}
		return host, def
	}
	if strings.Count(a, ":") > 1 {
		return a, def // a bare IPv6 address
	}
	if i := strings.LastIndexByte(a, ':'); i > 0 {
		if p, ok := parsePort(a[i+1:]); ok {
			return a[:i], p
		}
	}
	return a, def
}

func parsePort(s string) (int, bool) {
	p, err := strconv.Atoi(s)
	return p, err == nil && p > 0 && p < 65536
}

// Result is what Probe found.
type Result struct {
	Host      string
	Port      int
	Addresses []string
	Resolve   time.Duration
	Connect   time.Duration
	Handshake time.Duration
	// Protocol is the security protocol the server chose: "NLA", "TLS",
	// "RDP", or "" when it did not answer the handshake.
	Protocol string
	Err      error
	Stage    string // the stage that failed: "解析", "连接" or "握手"
}

// OK reports whether the host answered as an RDP server.
func (r Result) OK() bool { return r.Err == nil }

// Lines describes the result for the user.
func (r Result) Lines() []string {
	var out []string
	if len(r.Addresses) > 0 {
		out = append(out, fmt.Sprintf("解析 %s → %s（%s）", r.Host, strings.Join(r.Addresses, ", "), ms(r.Resolve)))
	}
	if r.Connect > 0 {
		out = append(out, fmt.Sprintf("TCP 端口 %d 已连通（%s）", r.Port, ms(r.Connect)))
	}
	if r.Protocol != "" {
		out = append(out, fmt.Sprintf("RDP 服务已应答，安全协议 %s（%s）", r.Protocol, ms(r.Handshake)))
	}
	if r.Err != nil {
		out = append(out, fmt.Sprintf("%s失败：%s", r.Stage, explain(r.Err)))
	}
	return out
}

// explain says what a network error means, in the words of the user.
func explain(err error) string {
	var dns *net.DNSError
	switch {
	case errors.As(err, &dns) && dns.IsNotFound:
		return "域名不存在"
	case errors.As(err, &dns):
		return "无法解析域名（DNS 无应答）"
	case errors.Is(err, syscall.ECONNREFUSED) || errors.Is(err, syscall.Errno(10061)):
		return "对方拒绝连接：这个端口上没有服务在运行"
	case errors.Is(err, syscall.Errno(10065)) || errors.Is(err, syscall.Errno(10051)):
		return "网络不可达"
	case errors.Is(err, os.ErrDeadlineExceeded) || isTimeout(err):
		return "超时没有应答（主机关机、端口未开放或被防火墙拦截）"
	case errors.Is(err, io.EOF) || errors.Is(err, syscall.Errno(10054)):
		return "连接被对方关闭（端口可能被代理或转发服务占用，但后面的电脑没有应答）"
	}
	return err.Error()
}

func isTimeout(err error) bool {
	var t interface{ Timeout() bool }
	return errors.As(err, &t) && t.Timeout()
}

func ms(d time.Duration) string { return fmt.Sprintf("%d ms", d.Milliseconds()) }

// x224Request is a TPKT/X.224 connection request with an RDP negotiation
// request for TLS and CredSSP (MS-RDPBCGR 2.2.1.1).
var x224Request = []byte{
	0x03, 0x00, 0x00, 0x13, // TPKT, length 19
	0x0e, 0xe0, 0x00, 0x00, 0x00, 0x00, 0x00, // X.224 CR
	0x01, 0x00, 0x08, 0x00, 0x03, 0x00, 0x00, 0x00, // RDP_NEG_REQ: SSL | HYBRID
}

// Probe resolves host, connects to port and starts the RDP handshake,
// within timeout for each stage.
func Probe(ctx context.Context, host string, port int, timeout time.Duration) Result {
	r := Result{Host: host, Port: port}
	start := time.Now()
	rctx, cancel := context.WithTimeout(ctx, timeout)
	ips, err := net.DefaultResolver.LookupIPAddr(rctx, host)
	cancel()
	r.Resolve = time.Since(start)
	if err != nil {
		r.Err, r.Stage = err, "解析"
		return r
	}
	for _, ip := range ips {
		r.Addresses = append(r.Addresses, ip.String())
	}

	start = time.Now()
	d := net.Dialer{Timeout: timeout}
	conn, err := d.DialContext(ctx, "tcp", net.JoinHostPort(host, strconv.Itoa(port)))
	if err != nil {
		r.Err, r.Stage = err, "连接"
		return r
	}
	r.Connect = max(time.Since(start), time.Microsecond)
	defer conn.Close()

	start = time.Now()
	conn.SetDeadline(time.Now().Add(timeout))
	proto, err := handshake(conn)
	r.Handshake = time.Since(start)
	if err != nil {
		r.Err, r.Stage = err, "握手"
		return r
	}
	r.Protocol = proto
	return r
}

func handshake(rw io.ReadWriter) (string, error) {
	if _, err := rw.Write(x224Request); err != nil {
		return "", err
	}
	var hdr [4]byte
	if _, err := io.ReadFull(rw, hdr[:]); err != nil {
		return "", fmt.Errorf("没有应答 RDP 握手: %w", err)
	}
	if hdr[0] != 0x03 {
		return "", fmt.Errorf("端口上不是 RDP 服务")
	}
	n := int(binary.BigEndian.Uint16(hdr[2:]))
	if n < 11 || n > 512 {
		return "", fmt.Errorf("RDP 应答长度异常 (%d)", n)
	}
	body := make([]byte, n-4)
	if _, err := io.ReadFull(rw, body); err != nil {
		return "", err
	}
	if body[1]&0xf0 != 0xd0 {
		return "", fmt.Errorf("RDP 服务拒绝连接")
	}
	neg := body[7:]
	if len(neg) < 8 {
		return "RDP", nil // an old server, without negotiation
	}
	switch neg[0] {
	case 0x02: // RDP_NEG_RSP
		switch binary.LittleEndian.Uint32(neg[4:]) {
		case 0:
			return "RDP", nil
		case 1:
			return "TLS", nil
		case 2, 8, 10:
			return "NLA", nil
		default:
			return "其他", nil
		}
	case 0x03: // RDP_NEG_FAILURE: an RDP server all the same, which wants another protocol
		switch binary.LittleEndian.Uint32(neg[4:]) {
		case 2: // SSL_NOT_ALLOWED_BY_SERVER
			return "RDP", nil
		case 5, 6: // HYBRID_REQUIRED_BY_SERVER, SSL_WITH_USER_AUTH_REQUIRED_BY_SERVER
			return "NLA", nil
		default:
			return "TLS", nil
		}
	}
	return "RDP", nil
}
