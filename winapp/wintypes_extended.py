from ctypes import c_uint64, c_ulong, c_long, c_longlong, c_int, c_int64, WINFUNCTYPE, CFUNCTYPE, Structure
from ctypes.wintypes import HANDLE, HWND, LPARAM, WPARAM, LONG, WORD, DWORD, LPCWSTR, LPWSTR, UINT, BOOL, HMODULE, BYTE

import sys
is_64_bit = sys.maxsize > 2**32

LONG_PTR = c_longlong if is_64_bit else c_long
ULONG_PTR = c_uint64 if is_64_bit else c_ulong
DWORD_PTR = ULONG_PTR
INT_PTR = c_int64 if is_64_bit else c_int
UINT_PTR = WPARAM

WNDPROC = WINFUNCTYPE(LONG_PTR, HWND, UINT, WPARAM, LPARAM)
WNDENUMPROC = WINFUNCTYPE(BOOL, HWND, LPARAM)
WINEVENTPROCTYPE = WINFUNCTYPE(None, HANDLE, DWORD, HWND, LONG, LONG, DWORD, DWORD)
ENUMRESNAMEPROCW = CFUNCTYPE(BOOL, HMODULE, LPCWSTR, LPWSTR, LONG_PTR)

class ACCEL(Structure):
    _fields_ = [
        ("fVirt", BYTE),
        ("key", WORD),
        ("cmd", WORD),
    ]

# Macros
def MAKELONG(wLow, wHigh):
    return LONG(wLow | wHigh << 16).value

def MAKELPARAM(l, h):
    return LPARAM(MAKELONG(l, h)).value

def LOWORD(l):
    return WORD(l & 0xFFFF).value

def HIWORD(l):
    return WORD((l >> 16) & 0xFFFF).value

def MAKEINTRESOURCEW(x):
    return LPCWSTR(x)
