package rdp

import (
	"sync"
	"syscall"
	"unsafe"

	"golang.org/x/sys/windows"
)

// diidEvents is IMsTscAxEvents, the control's outgoing interface.
var diidEvents = mustGUID("{336D5562-EFA8-482E-8CB3-C5C0FC7A7DB6}")

// DISPIDs of IMsTscAxEvents, read from the type library of mstscax.dll.
const (
	evConnecting            = 0x01
	evConnected             = 0x02
	evLoginComplete         = 0x03
	evDisconnected          = 0x04
	evEnterFullScreenMode   = 0x05
	evLeaveFullScreenMode   = 0x06
	evRequestGoFullScreen   = 0x08
	evRequestLeaveFullScr   = 0x09
	evFatalError            = 0x0a
	evWarning               = 0x0b
	evRemoteDesktopSize     = 0x0c
	evRequestContainerMin   = 0x0e
	evConfirmClose          = 0x0f
	evAutoReconnecting      = 0x11
	evLogonError            = 0x16
	evFocusReleased         = 0x17
	evAutoReconnected       = 0x21
	evAutoReconnecting2     = 0x22
	evAuthWarningDisplayed  = 0x12
	evAuthWarningDismissed  = 0x13
	evNetworkStatusChanged  = 0x20
	evServiceMessage        = 0x1c
	evUserNameAcquired      = 0x18
	evMouseInputModeChanged = 0x1a
)

// Event is something the control reported.
type Event struct {
	ID   int     // one of the Ev constants
	Args []int64 // its integer arguments, in order
}

// Events the client reports.
const (
	EvConnecting = iota + 1
	EvConnected
	EvLoginComplete
	EvDisconnected        // Args[0]: the disconnect reason
	EvFatalError          // Args[0]: the error code
	EvWarning             // Args[0]: the warning code
	EvLogonError          // Args[0]: the error code
	EvRequestGoFullScreen // the control asks the container for full screen
	EvRequestLeaveFullScreen
	EvRequestMinimize
	EvAutoReconnecting // Args[0]: the reason, Args[1]: the attempt
	EvAutoReconnected
	EvRemoteDesktopSizeChange // Args[0], Args[1]: width and height
	EvFocusReleased
)

// sink is a COM object implementing IMsTscAxEvents. All sinks share one
// vtable, whose callbacks are created once.
type sink struct {
	vtbl *[7]uintptr
	refs int32
	fn   func(Event)
}

var (
	sinkVtbl [7]uintptr
	sinksMu  sync.Mutex
	sinks    = map[uintptr]*sink{}
)

const (
	sOK          = 0
	eNoInterface = 0x80004002
	eNotImpl     = 0x80004001
)

func init() {
	sinkVtbl = [7]uintptr{
		syscall.NewCallback(func(this, riid, out uintptr) uintptr { // QueryInterface
			iid := *(*windows.GUID)(native(riid))
			if iid != iidIUnknown && iid != iidIDispatch && iid != diidEvents {
				*(*uintptr)(native(out)) = 0
				return eNoInterface
			}
			*(*uintptr)(native(out)) = this
			sinkRef(this, 1)
			return sOK
		}),
		syscall.NewCallback(func(this uintptr) uintptr { return uintptr(sinkRef(this, 1)) }),
		syscall.NewCallback(func(this uintptr) uintptr { return uintptr(sinkRef(this, -1)) }),
		syscall.NewCallback(func(this, count uintptr) uintptr { // GetTypeInfoCount
			*(*uint32)(native(count)) = 0
			return sOK
		}),
		syscall.NewCallback(func(this, i, lcid, out uintptr) uintptr { return eNotImpl }),
		syscall.NewCallback(func(this, riid, names, n, lcid, ids uintptr) uintptr { return eNotImpl }),
		syscall.NewCallback(sinkInvoke),
	}
}

func sinkRef(this uintptr, delta int32) int32 {
	sinksMu.Lock()
	defer sinksMu.Unlock()
	s := sinks[this]
	if s == nil {
		return 0
	}
	s.refs += delta
	if s.refs <= 0 {
		delete(sinks, this)
		return 0
	}
	return s.refs
}

// newSink returns a sink holding one reference, for the caller to release.
func newSink(fn func(Event)) uintptr {
	s := &sink{vtbl: &sinkVtbl, refs: 1, fn: fn}
	p := uintptr(unsafe.Pointer(s))
	sinksMu.Lock()
	sinks[p] = s
	sinksMu.Unlock()
	return p
}

// args reads the integer arguments of an event, in declaration order.
func args(params uintptr) []int64 {
	if params == 0 {
		return nil
	}
	dp := (*dispParams)(native(params))
	if dp.NArgs == 0 || dp.Args == nil {
		return nil
	}
	vars := unsafe.Slice(dp.Args, dp.NArgs)
	out := make([]int64, len(vars))
	for i := range vars {
		out[i] = vars[len(vars)-1-i].int()
	}
	return out
}

func sinkInvoke(this, dispid, riid, lcid, flags, params, result, excep, argErr uintptr) uintptr {
	sinksMu.Lock()
	s := sinks[this]
	sinksMu.Unlock()
	if s == nil || s.fn == nil {
		return sOK
	}
	a := args(params)
	var ev Event
	switch int32(dispid) {
	case evConnecting:
		ev.ID = EvConnecting
	case evConnected:
		ev.ID = EvConnected
	case evLoginComplete:
		ev.ID = EvLoginComplete
	case evDisconnected:
		ev.ID = EvDisconnected
	case evFatalError:
		ev.ID = EvFatalError
	case evWarning:
		ev.ID = EvWarning
	case evLogonError:
		ev.ID = EvLogonError
	case evRequestGoFullScreen:
		ev.ID = EvRequestGoFullScreen
	case evRequestLeaveFullScr:
		ev.ID = EvRequestLeaveFullScreen
	case evRequestContainerMin:
		ev.ID = EvRequestMinimize
	case evAutoReconnecting:
		ev.ID = EvAutoReconnecting
	case evAutoReconnecting2: // reason, network available, attempt, max attempts
		ev.ID = EvAutoReconnecting
		if len(a) >= 3 {
			a = []int64{a[0], a[2]}
		}
	case evAutoReconnected:
		ev.ID = EvAutoReconnected
	case evRemoteDesktopSize:
		ev.ID = EvRemoteDesktopSizeChange
	case evFocusReleased:
		ev.ID = EvFocusReleased
	case evConfirmClose:
		// Allow the close: the container asked for it.
		if params != 0 {
			dp := (*dispParams)(native(params))
			if dp.NArgs > 0 {
				v := unsafe.Slice(dp.Args, dp.NArgs)[0]
				if v.VT&vtByRef != 0 && v.Val != 0 {
					*(*int16)(native(v.Val)) = -1
				}
			}
		}
		return sOK
	default:
		return sOK
	}
	ev.Args = a
	s.fn(ev)
	return sOK
}
