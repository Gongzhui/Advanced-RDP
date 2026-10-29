// Package winutil has the few Win32 calls the windows of the app need
// beyond what MyGo offers.
package winutil

import (
	"unsafe"

	"golang.org/x/sys/windows"
)

var (
	user32              = windows.NewLazySystemDLL("user32.dll")
	procGetClientRect   = user32.NewProc("GetClientRect")
	procGetDpiForWindow = user32.NewProc("GetDpiForWindow")
	procMonitorFromWnd  = user32.NewProc("MonitorFromWindow")
	procGetMonitorInfo  = user32.NewProc("GetMonitorInfoW")
)

// ClientSize is the size of a window's client area, in pixels.
func ClientSize(hwnd uintptr) (int, int) {
	var r [4]int32
	procGetClientRect.Call(hwnd, uintptr(unsafe.Pointer(&r)))
	return int(r[2] - r[0]), int(r[3] - r[1])
}

// Scale is the window's DPI scale in percent: 100 at 96 DPI.
func Scale(hwnd uintptr) int {
	if procGetDpiForWindow.Find() != nil {
		return 100
	}
	dpi, _, _ := procGetDpiForWindow.Call(hwnd)
	if dpi == 0 {
		return 100
	}
	return int(dpi) * 100 / 96
}

// MonitorSize is the size, in pixels, of the display the window is on.
func MonitorSize(hwnd uintptr) (int, int) {
	mon, _, _ := procMonitorFromWnd.Call(hwnd, 2 /* MONITOR_DEFAULTTONEAREST */)
	info := struct {
		Size    uint32
		Monitor [4]int32
		Work    [4]int32
		Flags   uint32
	}{}
	info.Size = uint32(unsafe.Sizeof(info))
	procGetMonitorInfo.Call(mon, uintptr(unsafe.Pointer(&info)))
	return int(info.Monitor[2] - info.Monitor[0]), int(info.Monitor[3] - info.Monitor[1])
}
