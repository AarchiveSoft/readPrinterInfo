import os
import ctypes

# Adjust if your DLL lives elsewhere relative to this script
DLL_PATH = os.path.join(os.path.dirname(__file__), "data", "CspStat.dll")
dll = ctypes.WinDLL(DLL_PATH)

# Helper: bind an exported function if present
def bind(name, argtypes=(), restype=ctypes.c_int):
    try:
        fn = getattr(dll, name)
        fn.argtypes = list(argtypes)
        fn.restype = restype
        return fn
    except AttributeError:
        return None

# Bind functions we care about
GetPrinterPortNum     = bind("GetPrinterPortNum")             # () -> int
GetMediaCounter       = bind("GetMediaCounter", [ctypes.c_int])       # (port) -> int
GetInitialMediaCount  = bind("GetInitialMediaCount", [ctypes.c_int])  # (port) -> int
GetStatus             = bind("GetStatus", [ctypes.c_int])             # (port) -> int  (optional)

if GetMediaCounter is None or GetInitialMediaCount is None:
    raise RuntimeError("DLL is missing required exports (GetMediaCounter / GetInitialMediaCount).")

def find_port():
    # 1) Try vendor helper first
    if GetPrinterPortNum is not None:
        try:
            p = GetPrinterPortNum()
            if isinstance(p, int) and p >= 0:
                return p
        except OSError:
            pass
    # 2) Fallback: probe a reasonable range
    for p in range(0, 16):
        try:
            r = GetMediaCounter(p)
            if isinstance(r, int) and r >= 0:
                return p
        except OSError:
            continue
    return None

port = find_port()
if port is None:
    raise RuntimeError("Could not locate DS620 via CspStat.dll (make sure the printer is idle and connected).")

remaining = GetMediaCounter(port)
total     = GetInitialMediaCount(port) if GetInitialMediaCount is not None else -1
status    = GetStatus(port) if GetStatus is not None else None

print(f"Vendor port: {port}")
print(f"Remaining prints: {remaining}")
if total >= 0:
    print(f"Total capacity (current roll): {total}")
    if total > 0:
        pct = (remaining / total) * 100.0
        print(f"Remaining percent: {pct:.1f}%")
if status is not None:
    print(f"Raw status code: {status}")
