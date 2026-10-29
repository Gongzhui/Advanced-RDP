package rdp

import (
	"errors"
	"fmt"
	"syscall"
	"unsafe"

	"golang.org/x/sys/windows"
)

var (
	ole32    = windows.NewLazySystemDLL("ole32.dll")
	oleaut32 = windows.NewLazySystemDLL("oleaut32.dll")

	procOleInitialize    = ole32.NewProc("OleInitialize")
	procCoCreateInstance = ole32.NewProc("CoCreateInstance")
	procCLSIDFromString  = ole32.NewProc("CLSIDFromString")
	procCLSIDFromProgID  = ole32.NewProc("CLSIDFromProgID")
	procSysAllocString   = oleaut32.NewProc("SysAllocString")
	procSysFreeString    = oleaut32.NewProc("SysFreeString")
	procSysStringLen     = oleaut32.NewProc("SysStringLen")
	procVariantClear     = oleaut32.NewProc("VariantClear")
)

// native turns an address COM gave us into a pointer, without the
// uintptr-to-pointer conversion vet warns about.
func native(addr uintptr) unsafe.Pointer { return *(*unsafe.Pointer)(unsafe.Pointer(&addr)) }

// comCall calls method i of the COM interface obj.
//
//go:uintptrescapes
func comCall(obj uintptr, i int, args ...uintptr) uintptr {
	vtbl := *(*uintptr)(native(obj))
	fn := *(*uintptr)(unsafe.Add(native(vtbl), i*int(unsafe.Sizeof(uintptr(0)))))
	var a [16]uintptr
	a[0] = obj
	n := copy(a[1:], args)
	r, _, _ := syscall.SyscallN(fn, a[:n+1]...)
	return r
}

func failed(hr uintptr) bool { return int32(uint32(hr)) < 0 }

// HRESULTError is a failed COM call.
type HRESULTError struct {
	What string
	HR   uint32
	Desc string
}

func (e *HRESULTError) Error() string {
	if e.Desc != "" {
		return fmt.Sprintf("%s: %s (0x%08X)", e.What, e.Desc, e.HR)
	}
	return fmt.Sprintf("%s 失败 (0x%08X)", e.What, e.HR)
}

func hrError(what string, hr uintptr) error { return &HRESULTError{What: what, HR: uint32(hr)} }

func addRef(obj uintptr) {
	if obj != 0 {
		comCall(obj, 1)
	}
}

func release(obj uintptr) {
	if obj != 0 {
		comCall(obj, 2)
	}
}

func queryInterface(obj uintptr, iid *windows.GUID) uintptr {
	if obj == 0 {
		return 0
	}
	var out uintptr
	if failed(comCall(obj, 0, uintptr(unsafe.Pointer(iid)), uintptr(unsafe.Pointer(&out)))) {
		return 0
	}
	return out
}

func mustGUID(s string) windows.GUID {
	g, err := windows.GUIDFromString(s)
	if err != nil {
		panic(err)
	}
	return g
}

var (
	iidIUnknown                  = mustGUID("{00000000-0000-0000-C000-000000000046}")
	iidIDispatch                 = mustGUID("{00020400-0000-0000-C000-000000000046}")
	iidIConnectionPointContainer = mustGUID("{B196B284-BAB4-101A-B69C-00AA00341D07}")
	iidNull                      = windows.GUID{}
)

func bstr(s string) uintptr {
	p, err := windows.UTF16PtrFromString(s)
	if err != nil {
		p, _ = windows.UTF16PtrFromString("")
	}
	r, _, _ := procSysAllocString.Call(uintptr(unsafe.Pointer(p)))
	return r
}

func freeBSTR(b uintptr) {
	if b != 0 {
		procSysFreeString.Call(b)
	}
}

func bstrString(b uintptr) string {
	if b == 0 {
		return ""
	}
	n, _, _ := procSysStringLen.Call(b)
	return windows.UTF16ToString(unsafe.Slice((*uint16)(native(b)), n))
}

// Variant types.
const (
	vtEmpty    = 0
	vtI2       = 2
	vtI4       = 3
	vtBSTR     = 8
	vtDispatch = 9
	vtBool     = 11
	vtUnknown  = 13
	vtUI4      = 19
	vtByRef    = 0x4000
)

// variant is VARIANT on 64-bit Windows.
type variant struct {
	VT   uint16
	_    [3]uint16
	Val  uintptr
	Val2 uintptr
}

func (v *variant) clear() { procVariantClear.Call(uintptr(unsafe.Pointer(v))) }

// int returns an integer variant's value, also through a reference.
func (v *variant) int() int64 {
	switch v.VT {
	case vtI2, vtBool:
		return int64(int16(uint16(v.Val)))
	case vtI4, vtUI4:
		return int64(int32(uint32(v.Val)))
	case vtI4 | vtByRef, vtUI4 | vtByRef:
		if v.Val != 0 {
			return int64(*(*int32)(native(v.Val)))
		}
	}
	return 0
}

func (v *variant) string() string {
	if v.VT == vtBSTR {
		return bstrString(v.Val)
	}
	return ""
}

// toVariant makes a variant of a Go value. The variant owns any BSTR it
// holds: clear it after the call.
func toVariant(x any) (variant, error) {
	var v variant
	switch x := x.(type) {
	case bool:
		v.VT = vtBool
		if x {
			v.Val = 0xffff // VARIANT_TRUE
		}
	case int:
		v.VT, v.Val = vtI4, uintptr(uint32(int32(x)))
	case int32:
		v.VT, v.Val = vtI4, uintptr(uint32(x))
	case uint32:
		v.VT, v.Val = vtUI4, uintptr(x)
	case string:
		v.VT, v.Val = vtBSTR, bstr(x)
	default:
		return v, fmt.Errorf("unsupported variant type %T", x)
	}
	return v, nil
}

type dispParams struct {
	Args      *variant
	NamedArgs *int32
	NArgs     uint32
	NNamed    uint32
}

type excepInfo struct {
	Code        uint16
	_           uint16
	Source      uintptr
	Description uintptr
	HelpFile    uintptr
	HelpContext uint32
	_           uintptr
	_           uintptr
	SCode       int32
}

const (
	dispatchMethod      = 1
	dispatchPropertyGet = 2
	dispatchPropertyPut = 4
	dispidPropertyPut   = -3
	dispEException      = 0x80020009
)

// Dispatch is an IDispatch interface, called by member names.
type Dispatch struct {
	p   uintptr
	ids map[string]int32
}

func newDispatch(p uintptr) *Dispatch { return &Dispatch{p: p, ids: map[string]int32{}} }

// Release releases the interface.
func (d *Dispatch) Release() {
	if d != nil && d.p != 0 {
		release(d.p)
		d.p = 0
	}
}

func (d *Dispatch) id(name string) (int32, error) {
	if id, ok := d.ids[name]; ok {
		return id, nil
	}
	n, err := windows.UTF16PtrFromString(name)
	if err != nil {
		return 0, err
	}
	var id int32
	names := [1]*uint16{n}
	hr := comCall(d.p, 5, uintptr(unsafe.Pointer(&iidNull)), uintptr(unsafe.Pointer(&names[0])), 1, 0x0400, uintptr(unsafe.Pointer(&id)))
	if failed(hr) {
		return 0, hrError(name, hr)
	}
	d.ids[name] = id
	return id, nil
}

func (d *Dispatch) invoke(name string, flags uint16, args []any) (variant, error) {
	var result variant
	if d == nil || d.p == 0 {
		return result, errors.New(name + ": 控件已释放")
	}
	id, err := d.id(name)
	if err != nil {
		return result, err
	}
	vars := make([]variant, len(args))
	for i, a := range args {
		v, err := toVariant(a)
		if err != nil {
			return result, err
		}
		vars[len(args)-1-i] = v // in reverse order
	}
	defer func() {
		for i := range vars {
			vars[i].clear()
		}
	}()
	params := dispParams{NArgs: uint32(len(vars))}
	if len(vars) > 0 {
		params.Args = &vars[0]
	}
	named := int32(dispidPropertyPut)
	if flags == dispatchPropertyPut {
		params.NamedArgs, params.NNamed = &named, 1
	}
	var ex excepInfo
	var argErr uint32
	hr := comCall(d.p, 6, uintptr(id), uintptr(unsafe.Pointer(&iidNull)), 0x0400, uintptr(flags),
		uintptr(unsafe.Pointer(&params)), uintptr(unsafe.Pointer(&result)), uintptr(unsafe.Pointer(&ex)), uintptr(unsafe.Pointer(&argErr)))
	if failed(hr) {
		e := &HRESULTError{What: name, HR: uint32(hr)}
		if uint32(hr) == dispEException {
			e.Desc = bstrString(ex.Description)
			if ex.SCode != 0 {
				e.HR = uint32(ex.SCode)
			}
		}
		freeBSTR(ex.Source)
		freeBSTR(ex.Description)
		freeBSTR(ex.HelpFile)
		return result, e
	}
	return result, nil
}

// Put sets a property.
func (d *Dispatch) Put(name string, value any) error {
	_, err := d.invoke(name, dispatchPropertyPut, []any{value})
	return err
}

// GetInt reads an integer property.
func (d *Dispatch) GetInt(name string) (int64, error) {
	v, err := d.invoke(name, dispatchPropertyGet, nil)
	defer v.clear()
	return v.int(), err
}

// GetDispatch reads a property holding an object.
func (d *Dispatch) GetDispatch(name string) (*Dispatch, error) {
	v, err := d.invoke(name, dispatchPropertyGet, nil)
	if err != nil {
		return nil, err
	}
	if (v.VT != vtDispatch && v.VT != vtUnknown) || v.Val == 0 {
		v.clear()
		return nil, fmt.Errorf("%s: 不是对象", name)
	}
	p := v.Val
	if v.VT == vtUnknown {
		p = queryInterface(v.Val, &iidIDispatch)
		v.clear()
		if p == 0 {
			return nil, fmt.Errorf("%s: 不支持 IDispatch", name)
		}
	}
	return newDispatch(p), nil // the variant's reference is ours now
}

// Call calls a method and returns its result as a string, if it is one.
func (d *Dispatch) Call(name string, args ...any) (string, error) {
	v, err := d.invoke(name, dispatchMethod, args)
	defer v.clear()
	return v.string(), err
}
