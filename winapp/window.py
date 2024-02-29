from ctypes import windll, Structure, sizeof, c_int, c_uint, byref, c_voidp
from ctypes.wintypes import LPCWSTR, HANDLE, RECT, POINT, HINSTANCE, DWORD, INT, HWND, HMENU, LPVOID

from winapp.const import *
from winapp.wintypes_extended import WNDPROC, MAKELONG, MAKELPARAM
from winapp.dlls import gdi32, kernel32, user32, uxtheme
from winapp.controls.common import *
from winapp.themes import *

# tmp
#import winapp.const as _tmp
#WM = dict((v, k) for k, v in _tmp.__dict__.items() if k.startswith("WM_"))
#def DBG_MSG(msg, *args):
#    print(WM[msg] if msg in WM else hex(msg), *args)


hdc = user32.GetDC(0)
DPI_X = gdi32.GetDeviceCaps(hdc, LOGPIXELSX)  # 96
DPI_Y = gdi32.GetDeviceCaps(hdc, LOGPIXELSY)  # 96
user32.ReleaseDC(0, hdc)


class WNDCLASSEX(Structure):
    def __init__(self, *args, **kwargs):
        super(WNDCLASSEX, self).__init__(*args, **kwargs)
        self.cbSize = sizeof(self)
    _fields_ = [
        ("cbSize", c_uint),
        ("style", c_uint),
        ("lpfnWndProc", WNDPROC),
        ("cbClsExtra", c_int),
        ("cbWndExtra", c_int),
        ("hInstance", HANDLE),
        ("hIcon", HANDLE),
        ("hCursor", HANDLE),
        ("hBrush", HANDLE),
        ("lpszMenuName", LPCWSTR),
        ("lpszClassName", LPCWSTR),
        ("hIconSm", HANDLE)
        ]

class MINMAXINFO(Structure):
    _fields_ = [
        ("ptReserved", POINT),
        ("ptMaxSize", POINT),
        ("ptMaxPosition", POINT),
        ("ptMinTrackSize", POINT),
        ("ptMaxTrackSize", POINT),
        ]


class Window(object):

    def __init__(self, window_class, style=WS_CHILD | WS_VISIBLE, ex_style=0,
            left=0, top=0, width=0, height=0, window_title=0, hmenu=0, parent_window=None, wrap_hwnd=None):

        self.parent_window = parent_window
        self.is_dark = False
        self.__old_proc = None
        self.__new_proc = None
        self.__message_map = {}

        if wrap_hwnd:
            self.hwnd = wrap_hwnd
        else:
            self.hwnd = user32.CreateWindowExW(
                    ex_style,
                    window_class,
                    window_title,
                    style,
                    left, top,
                    width, height,
                    parent_window.hwnd if parent_window else 0,
                    hmenu,
                    0,  # hInstance
                    0   # lpParam
                    )

    def destroy_window(self):
        if self.__old_proc:
            user32.SetWindowLongPtrW(self.hwnd, GWL_WNDPROC, self.__old_proc)
            self.__message_map = {}
            self.__old_proc = None
        user32.DestroyWindow(self.hwnd)

    def window_proc_callback(self, hwnd, msg, wparam, lparam):
        if msg in self.__message_map:
            for callback in self.__message_map[msg]:
                res = callback(hwnd, wparam, lparam)
                if res is not None:
                    return res
        return self.__old_proc(hwnd, msg, wparam, lparam)

#    def register_message_callback(self, msg, callback, overwrite=False):
#        if overwrite:
#            self.__message_map[msg] = [callback]
#        else:
#            if msg not in self.__message_map:
#                self.__message_map[msg] = []
#            self.__message_map[msg].append(callback)

    def register_message_callback(self, msg, callback):
        if msg not in self.__message_map:
            self.__message_map[msg] = []
        self.__message_map[msg].append(callback)
        if self.__new_proc is None:
            self.__new_proc = WNDPROC(self.window_proc_callback)
            self.__old_proc = user32.SetWindowLongPtrW(self.hwnd, GWL_WNDPROC, self.__new_proc)

    def unregister_message_callback(self, msg, callback=None):
        if msg in self.__message_map:
            if callback is None:
                del self.__message_map[msg]
            elif callback in self.__message_map[msg]:
                self.__message_map[msg].remove(callback)
                if len(self.__message_map[msg]) == 0:
                    del self.__message_map[msg]

    def set_window_text(self, txt):
        user32.SetWindowTextW(self.hwnd, txt)

    def set_stayontop(self, flag=True):
        user32.SetWindowPos(self.hwnd, HWND_TOPMOST if flag else HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE)

    def set_layered(self):
        # Set WS_EX_LAYERED on this window
        user32.SetWindowLongPtrA(self.hwnd, GWL_EXSTYLE,
                user32.GetWindowLongPtrA(self.hwnd, GWL_EXSTYLE) | WS_EX_LAYERED)

    def set_alpha(self, alpha):
        ''' only works if WS_EX_LAYERED was passed as ex_style to window_create'''
        user32.SetLayeredWindowAttributes(self.hwnd, 0, alpha, LWA_ALPHA)
        user32.RedrawWindow(self.hwnd, 0, 0, RDW_ERASE | RDW_INVALIDATE | RDW_FRAME | RDW_ALLCHILDREN)

    def resize_window(self, w, h):
        user32.SetWindowPos(self.hwnd, 0, 0, 0, w, h, SWP_NOMOVE)

    def activate_window(self):
        user32.SetActiveWindow(self.hwnd)

    def show(self, cmd_show=SW_SHOW):
        user32.ShowWindow(self.hwnd, cmd_show)

    def enable_window(self, flag):
        user32.EnableWindow(self.hwnd, flag)

    def set_window_pos(self, x, y, w, h, hwnd_insert_after=0, flags=0):
        user32.SetWindowPos(self.hwnd, hwnd_insert_after, x, y, w, h, flags)

    def set_foreground_window(self):
        user32.SetForegroundWindow(self.hwnd)

    def hide_focus_rects(self):
        user32.SendMessageW(self.hwnd, WM_CHANGEUISTATE, MAKELONG(UIS_SET, UISF_HIDEFOCUS), 0)

    def get_window_rect(self):
        rc = RECT()
        user32.GetWindowRect(self.hwnd, byref(rc))
        return rc

    def get_client_rect(self):
        rc = RECT()
        user32.GetClientRect(self.hwnd, byref(rc))
        return rc

    def set_font(self, font_name='MS Shell Dlg', font_size=8, font_weight=FW_DONTCARE, font_italic=FALSE, hfont=None):
        if not hfont:
            cHeight = -kernel32.MulDiv(font_size, DPI_Y, 72)
            hfont = gdi32.CreateFontW(cHeight, 0, 0, 0, font_weight, font_italic, FALSE, FALSE, ANSI_CHARSET, OUT_TT_PRECIS,
                        CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, font_name)
        user32.SendMessageW(self.hwnd, WM_SETFONT, hfont, MAKELPARAM(1, 0))

    def set_parent(self, win=None):
        user32.SetParent(self.hwnd, win.hwnd if win else 0)

    def move_window(self, x, y, width, height, repaint=1):
        return user32.MoveWindow(self.hwnd, x, y, width, height, repaint)

    def update_window(self):
        user32.UpdateWindow(self.hwnd)

    def redraw_window(self):
        user32.RedrawWindow(self.hwnd, 0, 0, RDW_ERASE | RDW_INVALIDATE | RDW_FRAME | RDW_ALLCHILDREN)

    def force_redraw_window(self):
        user32.SetWindowPos(self.hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED)

    def apply_theme(self, is_dark):
        self.is_dark = is_dark
        uxtheme.SetWindowTheme(self.hwnd, 'DarkMode_Explorer' if is_dark else 'Explorer', None)