from ctypes import (Structure, create_unicode_buffer, c_voidp, windll, cast, byref, sizeof,
        c_wchar_p, c_ubyte, POINTER)
from ctypes.wintypes import (SHORT, WORD, DWORD, HWND, HINSTANCE, LPWSTR, LPCWSTR, LPVOID,
        HANDLE, INT, WCHAR, BYTE, COLORREF, HDC, UINT, WPARAM, LPARAM, LONG, HGLOBAL)
from .wintypes_extended import WINFUNCTYPE, UINT_PTR
from .dlls import shell32, comdlg32, gdi32
from .const import *

from .window import *
from .themes import *
from .controls.button import *
from .controls.edit import *
from .controls.static import *


DIALOG_CLASS = '#32770'

BUTTON    = 0x0080
EDIT      = 0x0081
STATIC    = 0x0082
LISTBOX   = 0x0083
SCROLLBAR = 0x0084
COMBOBOX  = 0x0085

BUTTON_COMMAND_IDS = {
    MB_OK: (IDOK,),
    MB_OKCANCEL: (IDOK, IDCANCEL),
    MB_ABORTRETRYIGNORE: (IDABORT, IDRETRY, IDIGNORE),
    MB_YESNOCANCEL: (IDYES, IDNO, IDCANCEL),
    MB_YESNO: (IDYES, IDNO),
    MB_RETRYCANCEL: (IDRETRY, IDCANCEL)
}

class SHSTOCKICONINFO(Structure):
    def __init__(self, *args, **kwargs):
        super(SHSTOCKICONINFO, self).__init__(*args, **kwargs)
        self.cbSize = sizeof(self)
    _fields_ = (
        ('cbSize', DWORD),
        ('hIcon', HANDLE),
        ('iSysImageIndex', INT),
        ('iIcon', INT),
        ('szPath', WCHAR * MAX_PATH),
    )

class DLGTEMPLATEEX_PARTIAL(Structure):
    _pack_ = 2
    _fields_ = (
        ('dlgVer', WORD),
        ('signature', WORD),
        ('helpID', DWORD),
        ('exStyle', DWORD),
        ('style', DWORD),
        ('cDlgItems', WORD),
        ('x', SHORT),
        ('y', SHORT),
        ('cx', SHORT),
        ('cy', SHORT),
        ('menu', WORD),
    )

class DLGITEMTEMPLATEEX_PARTIAL(Structure):
    _pack_ = 2
    _fields_ = (
        ('helpID', DWORD),
        ('exStyle', DWORD),
        ('style', DWORD),
        ('x', SHORT),
        ('y', SHORT),
        ('cx', SHORT),
        ('cy', SHORT),
        ('id', DWORD),
        ('windowClass', DWORD), # array of 2 WORDs
    )

OFNHOOKPROC = WINFUNCTYPE(UINT_PTR, HWND, UINT, WPARAM, LPARAM)
LPOFNHOOKPROC = OFNHOOKPROC #POINTER(OFNHOOKPROC) #c_voidp # TODO

class OPENFILENAMEW(Structure):
    _fields_ = (
        ('lStructSize', DWORD),
        ('hwndOwner', HWND),
        ('hInstance', HINSTANCE),
        ('lpstrFilter', LPWSTR),
        ('lpstrCustomFilter', LPWSTR),
        ('nMaxCustFilter', DWORD),
        ('nFilterIndex', DWORD),
        ('lpstrFile', LPWSTR),
        ('nMaxFile', DWORD),
        ('lpstrFileTitle', LPWSTR),
        ('nMaxFileTitle', DWORD),
        ('lpstrInitialDir', LPCWSTR),
        ('lpstrTitle', LPCWSTR),
        ('Flags', DWORD),
        ('nFileOffset', WORD),
        ('nFileExtension', WORD),
        ('lpstrDefExt', LPCWSTR),
        ('lCustData', LPARAM),
        ('lpfnHook', LPOFNHOOKPROC),
        ('lpTemplateName', LPCWSTR),
        ('pvReserved', LPVOID),
        ('dwReserved', DWORD),
        ('FlagsEx', DWORD),
    )

class LOGFONTW(Structure):
    def __str__(self):
        return  "('%s' %d)" % (self.lfFaceName, self.lfHeight)
#    def __repr__(self):
#        return "<LOGFONTW '%s' %d>" % (self.lfFaceName, self.lfHeight)
    _fields_ = [
        # C:/PROGRA~1/MIAF9D~1/VC98/Include/wingdi.h 1090
        ('lfHeight', LONG),
        ('lfWidth', LONG),
        ('lfEscapement', LONG),
        ('lfOrientation', LONG),
        ('lfWeight', LONG),
        ('lfItalic', BYTE),
        ('lfUnderline', BYTE),
        ('lfStrikeOut', BYTE),
        ('lfCharSet', BYTE),
        ('lfOutPrecision', BYTE),
        ('lfClipPrecision', BYTE),
        ('lfQuality', BYTE),
        ('lfPitchAndFamily', BYTE),
        ('lfFaceName', WCHAR * LF_FACESIZE),
    ]

class CHOOSEFONTW(Structure):
    def __init__(self, *args, **kwargs):
        super(CHOOSEFONTW, self).__init__(*args, **kwargs)
        self.lStructSize = sizeof(self)
    _fields_ = [
        ('lStructSize',                 DWORD),
        ('hwndOwner',                   HWND),
        ('hDC',                         HDC),
        ('lpLogFont',                   POINTER(LOGFONTW)),
        ('iPointSize',                  INT),
        ('Flags',                       DWORD),
        ('rgbColors',                   COLORREF),
        ('lCustData',                   LPARAM),
        ('lpfnHook',                    c_voidp),  # LPCFHOOKPROC
        ('lpTemplateName',              LPCWSTR),
        ('hInstance',                   HINSTANCE),
        ('lpszStyle',                   LPWSTR),
        ('nFontType',                   WORD),
        ('___MISSING_ALIGNMENT__',      WORD),
        ('nSizeMin',                    INT),
        ('nSizeMax',                    INT),
    ]

class PRINTDLGW(Structure):
    def __init__(self, *args, **kwargs):
        super(PRINTDLGW, self).__init__(*args, **kwargs)
        self.lStructSize = sizeof(self)
    _fields_ = [
        ('lStructSize',                 DWORD),
        ('hwndOwner',                   HWND),
        ('hDevMode',                    HGLOBAL),
        ('hDevNames',                   HGLOBAL),
        ('hDC',                         HDC),
        ('Flags',                       DWORD),
        ('nFromPage',                   WORD),
        ('nToPage',                     WORD),
        ('nMinPage',                    WORD),
        ('nMaxPage',                    WORD),
        ('nCopies',                     WORD),
        ('hInstance',                   HINSTANCE),
        ('lCustData',                   LPARAM),
        ('lpfnPrintHook',               LPVOID),  # LPPRINTHOOKPROC
        ('lpfnSetupHook',               LPVOID),  # LPSETUPHOOKPROC
        ('lpPrintTemplateName',         LPCWSTR),
        ('lpSetupTemplateName',         LPCWSTR),
        ('hPrintTemplate',              HGLOBAL),
        ('hSetupTemplate',              HGLOBAL),
    ]
LPPRINTDLGW = POINTER(PRINTDLGW)

comdlg32.PrintDlgW.argtypes = (LPPRINTDLGW,)

class DOCINFOW(Structure):
    def __init__(self, *args, **kwargs):
        super(DOCINFOW, self).__init__(*args, **kwargs)
        self.cbSize = sizeof(self)
    _fields_ = [
        ('cbSize', INT),
        ('lpszDocName', LPCWSTR),
        ('lpszOutput', LPCWSTR),
        ('lpszDatatype', LPCWSTR),
        ('fwType', DWORD),
    ]

gdi32.StartDocW.argtypes = (HDC, POINTER(DOCINFOW))

gdi32.EndDoc.argtypes = (HDC,)
gdi32.AbortDoc.argtypes = (HDC,)

gdi32.StartPage.argtypes = (HDC,)
gdi32.EndPage.argtypes = (HDC,)

class PAGESETUPDLGW(Structure):
    def __init__(self, *args, **kwargs):
        super(PAGESETUPDLGW, self).__init__(*args, **kwargs)
        self.lStructSize = sizeof(self)
    _fields_ = [
        ('lStructSize',                 DWORD),
        ('hwndOwner',                   HWND),
        ('hDevMode',                    HGLOBAL),
        ('hDevNames',                   HGLOBAL),
        ('Flags',                       DWORD),
        ('ptPaperSize',                 POINT),
        ('rtMinMargin',                 RECT),
        ('rtMargin',                    RECT),
        ('hInstance',                   HINSTANCE),
        ('lCustData',                   LPARAM),
        ('lpfnPageSetupHook',           LPVOID),  # LPPAGESETUPHOOK
        ('lpfnPagePaintHook',           LPVOID),  # LPPAGEPAINTHOOK
        ('lpPageSetupTemplateName',     LPCWSTR),
        ('hPageSetupTemplate',          HGLOBAL),
    ]

LPPAGESETUPDLGW = POINTER(PAGESETUPDLGW)

comdlg32.PageSetupDlgW.argtypes = (LPPAGESETUPDLGW,)

# https://learn.microsoft.com/en-us/windows/win32/api/commdlg/ns-commdlg-choosecolorw
class CHOOSECOLORW(Structure):
    def __init__(self, *args, **kwargs):
        super(CHOOSECOLORW, self).__init__(*args, **kwargs)
        self.lStructSize = sizeof(self)
    _fields_ = [
        ("lStructSize",                 DWORD),
        ("hwndOwner",                   HWND),
        ("hInstance",                   HWND),
        ("rgbResult",                   COLORREF),
        ("lpCustColors",                POINTER(COLORREF)),
        ("Flags",                       DWORD),
        ("lCustData",                   LPARAM),
        ("lpfnHook",                    LPVOID),  # LPCCHOOKPROC
        ("lpTemplateName",              LPCWSTR),
    ]

LPCHOOSECOLORW = POINTER(CHOOSECOLORW)

comdlg32.ChooseColorW.argtypes = (LPCHOOSECOLORW,)

# modern icon (flat)
def get_stock_icon(siid):
    sii = SHSTOCKICONINFO()
    SHGSI_ICON = 0x000000100
    shell32.SHGetStockIconInfo(siid, SHGSI_ICON, byref(sii))
    return sii.hIcon

def calculate_text_size(text, font_name='MS Shell Dlg', font_size=8, hfont=None):
    hdc = user32.GetDC(0)
    no_hfont = hfont is None
    if no_hfont:
        hfont = gdi32.CreateFontW(font_size, 0, 0, 0, FW_DONTCARE, FALSE, FALSE, FALSE, ANSI_CHARSET, OUT_TT_PRECIS,
                CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, font_name)
    gdi32.SelectObject(hdc, hfont)
    rc = RECT(0, 0, 0, 0)
    user32.DrawTextW(hdc, text, -1, byref(rc), DT_CALCRECT | DT_LEFT | DT_NOPREFIX)
    user32.ReleaseDC(0, hdc)
    if no_hfont:
        gdi32.DeleteObject(hfont)
    return rc.right, rc.bottom

# logical coordinates, not pixels
def calculate_multiline_text_height(text, text_width, font_name='MS Shell Dlg', font_size=8, font_weight=FW_NORMAL, font_italic=FALSE, hfont=None):
    hdc = user32.GetDC(0)
    no_hfont = hfont is None
    if no_hfont:
        hfont = gdi32.CreateFontW(
            -kernel32.MulDiv(font_size, gdi32.GetDeviceCaps(hdc, LOGPIXELSX), 72),
            0, 0, 0,
            font_weight,
            font_italic,
            FALSE,
            FALSE,
            DEFAULT_CHARSET,
            OUT_OUTLINE_PRECIS,  #OUT_TT_PRECIS,
            CLIP_DEFAULT_PRECIS,
            #CLEARTYPE_QUALITY,
            DEFAULT_QUALITY,
            #VARIABLE_PITCH | FF_SWISS,  #
            DEFAULT_PITCH | FF_DONTCARE,
            font_name
        )

    gdi32.SelectObject(hdc, hfont)

    # If the dialog box template has the DS_SETFONT or DS_SHELLFONT style, the base units are the average
    # width and height, in pixels, of the characters in the font specified by the template.
    tm = TEXTMETRICW()
    gdi32.GetTextMetricsW(hdc, byref(tm))
    text_width  = kernel32.MulDiv(text_width, tm.tmAveCharWidth, 4)

    h = user32.DrawTextW(hdc, text, -1, byref(RECT(0, 0, text_width, 0)), DT_CALCRECT | DT_WORDBREAK | DT_NOPREFIX | DT_LEFT | DT_TOP)
    user32.ReleaseDC(0, hdc)
    if no_hfont:
        gdi32.DeleteObject(hfont)
    return kernel32.MulDiv(h, 8, tm.tmHeight)


class DialogTemplate(object):

    def __init__(self):
        self.__num_controls = 0
        self.__control_id_counter = 1000
        self.__dialog_item_data = b''

    def add_control(self, control_class, text, x, y, w, h, control_id=-1, style=WS_CHILD | WS_VISIBLE, exstyle=0):
        self.__control_id_counter +=1
        self.__num_controls += 1
        dlg_item_data = DLGITEMTEMPLATEEX_PARTIAL(
                0,
                exstyle,
                style,
                x, y,
                w, h,
                control_id,
                (control_class << 16) | 0xffff
                )
        dlg_item_data = bytes(dlg_item_data)
        if type(text) == int:
            dlg_item_data += bytes(MAKEINTRESOURCEW(text))
        else:
            dlg_item_data += bytes(create_unicode_buffer(text))
        dlg_item_data += bytes(WORD(0))
        if len(dlg_item_data) % 4:
            dlg_item_data += b'\x00' * (4 - len(dlg_item_data) % 4)
        self.__dialog_item_data += bytes(dlg_item_data)
        return control_id

    def create(self, x, y, w, h, dialog_title='', font='MS Shell Dlg', font_height=8, show_icon=False,
            style=WS_CAPTION | WS_SYSMENU | DS_CENTER | DS_NOIDLEMSG | DS_SETFONT,
            exstyle=0,
            ):

        if not show_icon:
            exstyle |= WS_EX_DLGMODALFRAME
        # https://learn.microsoft.com/en-us/windows/win32/dlgbox/dlgtemplateex
        dlg_data = DLGTEMPLATEEX_PARTIAL(
                1,
                0xffff,
                0,
                exstyle,
                style,
                self.__num_controls,
                x, y, w, h,
                0,
                )
        dlg_data = bytes(dlg_data)
        dlg_data += bytes(WORD(0))
        dlg_data += bytes(create_unicode_buffer(dialog_title))
        dlg_data += bytes(WORD(font_height))
        dlg_data += bytes(WORD(400)) # weight
        dlg_data += b'\x00'
        dlg_data += b'\x01'
        dlg_data += bytes(create_unicode_buffer(font)) # 14

        if len(dlg_data) % 4:
            dlg_data += b'\x00' * (4 - len(dlg_data) % 4)
        return dlg_data + self.__dialog_item_data


class Dialog(Window):

    def __init__(self, parent_window, dialog_dict, dialog_proc_callback, handle_static=True):

        super().__init__(DIALOG_CLASS, parent_window=parent_window, wrap_hwnd=0)

        dialog = DialogTemplate()
        for control in dialog_dict['controls']:
            dialog.add_control(
                    eval(control['class']),
                    control['caption'] if 'caption' in control else '',
                    *control['rect'],
                    control_id=control['id'],
                    style=control['style'] if 'style' in control else 0,
                    )
        dialog_data = dialog.create(
                *dialog_dict['rect'],
                dialog_dict['caption'],
                *dialog_dict['font'],
                False,
                style=dialog_dict['style'] if 'style' in dialog_dict else 0,
                exstyle=dialog_dict['exstyle'] if 'exstyle' in dialog_dict else 0
                )

        self.__dialog_data = (c_ubyte * len(dialog_data))(*dialog_data)

        def _dialog_proc_callback(hwnd, msg, wparam, lparam):

            if msg == WM_CLOSE:
                if self.__is_async:
                    user32.DestroyWindow(self.hwnd)
                    self.hwnd = None
                    self.children = []
                    self.controls = []
                    self.parent_window._dialog_remove(self)
                else:
                    user32.EndDialog(hwnd, 0)

            elif msg == WM_INITDIALOG:
                self.hwnd = hwnd

                dwm_use_dark_mode(hwnd, self.is_dark)
                hfont = user32.SendMessageW(hwnd, WM_GETFONT, 0, 0)
                self.controls = self.get_children()
                for hwnd_control in self.controls:
                    buf = create_unicode_buffer(32)
                    user32.GetClassNameW(hwnd_control, buf, 32)
                    window_class = buf.value

                    if window_class == 'Button':
                        uxtheme.SetWindowTheme(hwnd_control, 'DarkMode_Explorer' if self.is_dark else 'Explorer', None)

                        style = user32.GetWindowLongPtrA(hwnd_control, GWL_STYLE)

                        if style & BS_TYPEMASK == BS_AUTOCHECKBOX or style & BS_TYPEMASK == BS_AUTORADIOBUTTON:

                            rc = RECT()
                            user32.GetClientRect(hwnd_control, byref(rc))

                            buf = create_unicode_buffer(64)
                            user32.GetWindowTextW(hwnd_control, buf, 64)
                            window_title = buf.value.replace('&', '')
                            user32.SetWindowTextW(hwnd_control, window_title)

                            # wrap to prevent garbage collection while dialog exists
                            window_checkbox = Window('Button', parent_window=self, wrap_hwnd=hwnd_control)

                            text_width, text_height = calculate_text_size(window_title, hfont=hfont)
                            static = Static(parent_window=window_checkbox,
                                    style=WS_CHILD | SS_SIMPLE | WS_VISIBLE,
                                    left=16, top=3, width=text_width, height=text_height,
                                    window_title=window_title)
                            static.set_font(hfont=hfont)

                            def _on_WM_CTLCOLORSTATIC(hwnd, wparam, lparam):
                                gdi32.SetTextColor(wparam, TEXT_COLOR_DARK if self.is_dark else COLOR_WINDOWTEXT)
                                gdi32.SetBkColor(wparam, BG_COLOR_DARK if self.is_dark else user32.GetSysColor(COLOR_3DFACE))
                                return gdi32.GetStockObject(DC_BRUSH)
                            window_checkbox.register_message_callback(WM_CTLCOLORSTATIC, _on_WM_CTLCOLORSTATIC)

                        elif style & BS_TYPEMASK == BS_GROUPBOX:
                            rc = RECT()
                            user32.GetWindowRect(hwnd_control, byref(rc))
                            user32.MapWindowPoints(None, hwnd, byref(rc), 2)
                            buf = create_unicode_buffer(64)
                            user32.GetWindowTextW(hwnd_control, buf, 64)

                            static = Static(parent_window=self,
                                    style=WS_CHILD | SS_SIMPLE | WS_VISIBLE,
                                    ex_style=WS_EX_TRANSPARENT,
                                    left=rc.left + 10, top=rc.top, width=rc.right - rc.left - 16, height=rc.bottom - rc.top,
                                    window_title=buf.value)
                            static.set_font(hfont=hfont)

                    elif window_class == 'Edit' and self.is_dark:
                        # check parent
                        user32.GetClassNameW(user32.GetParent(hwnd_control), buf, 32)
                        if buf.value != 'ComboBox':
                            user32.SetWindowLongPtrA(hwnd_control, GWL_EXSTYLE,
                                    user32.GetWindowLongPtrA(hwnd_control, GWL_EXSTYLE) & ~WS_EX_CLIENTEDGE)
                            user32.SetWindowLongPtrA(hwnd_control, GWL_STYLE,
                                    user32.GetWindowLongPtrA(hwnd_control, GWL_STYLE) | WS_BORDER)

                            rc = RECT()
                            user32.GetWindowRect(hwnd_control, byref(rc))
                            w, h = rc.right - rc.left, rc.bottom - rc.top
                            user32.SendMessageW(hwnd_control, EM_SETMARGINS, EC_LEFTMARGIN, 2)
                            user32.MapWindowPoints(None, user32.GetParent(hwnd_control), byref(rc), 1)
                            user32.SetWindowPos(hwnd_control, 0, rc.left, rc.top + 1, w, h - 2, SWP_FRAMECHANGED)

                    elif window_class == 'ComboLBox' and self.is_dark:
                        uxtheme.SetWindowTheme(hwnd_control, 'DarkMode_CFD', None)

                        user32.SetWindowLongPtrA(hwnd_control, GWL_EXSTYLE,
                                user32.GetWindowLongPtrA(hwnd_control, GWL_EXSTYLE) & ~WS_EX_CLIENTEDGE)
                        user32.SetWindowLongPtrA(hwnd_control, GWL_STYLE,
                                user32.GetWindowLongPtrA(hwnd_control, GWL_STYLE) | WS_BORDER)
                        user32.SetWindowPos(hwnd_control, 0, 0, 0, 0, 0,
                                SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED)

                    elif window_class == 'ComboBox' and self.is_dark:
                        uxtheme.SetWindowTheme(hwnd_control, 'DarkMode_CFD', None)

            elif msg == WM_THEMECHANGED:
                dwm_use_dark_mode(hwnd, self.is_dark)
                for hwnd_control in self.controls:
                    buf = create_unicode_buffer(32)
                    user32.GetClassNameW(hwnd_control, buf, 32)
                    window_class = buf.value
                    if window_class == 'Button':
                        uxtheme.SetWindowTheme(hwnd_control, 'DarkMode_Explorer' if self.is_dark else 'Explorer', None)

                    elif window_class == 'Edit':
                        rc = RECT()
                        user32.GetWindowRect(hwnd_control, byref(rc))
                        w, h = rc.right - rc.left, rc.bottom - rc.top

                        if self.is_dark:
                            user32.SetWindowLongPtrA(hwnd_control, GWL_EXSTYLE,
                                    user32.GetWindowLongPtrA(hwnd_control, GWL_EXSTYLE) & ~WS_EX_CLIENTEDGE)
                            user32.SetWindowLongPtrA(hwnd_control, GWL_STYLE,
                                    user32.GetWindowLongPtrA(hwnd_control, GWL_STYLE) | WS_BORDER)

                            user32.SendMessageW(hwnd_control, EM_SETMARGINS, EC_LEFTMARGIN, 2)
                            user32.MapWindowPoints(None, user32.GetParent(hwnd_control), byref(rc), 1)
                            user32.SetWindowPos(hwnd_control, 0, rc.left, rc.top + 1, w, h - 2, 0)  #, SWP_FRAMECHANGED)
                        else:
                            user32.SetWindowLongPtrA(hwnd_control, GWL_EXSTYLE,
                                    user32.GetWindowLongPtrA(hwnd_control, GWL_EXSTYLE) | WS_EX_CLIENTEDGE)
                            user32.SetWindowLongPtrA(hwnd_control, GWL_STYLE,
                                    user32.GetWindowLongPtrA(hwnd_control, GWL_STYLE) & ~WS_BORDER)

                            user32.SendMessageW(hwnd_control, EM_SETMARGINS, EC_LEFTMARGIN, 0)
                            user32.MapWindowPoints(None, user32.GetParent(hwnd_control), byref(rc), 1)
                            user32.SetWindowPos(hwnd_control, 0, rc.left, rc.top - 1, w, h + 2, 0)  #, SWP_FRAMECHANGED)

                self.force_redraw_window()

            elif self.is_dark:
                if msg == WM_CTLCOLORDLG or (msg == WM_CTLCOLORSTATIC and handle_static):
                    gdi32.SetTextColor(wparam, TEXT_COLOR_DARK)
                    gdi32.SetBkColor(wparam, BG_COLOR_DARK)
                    return BG_BRUSH_DARK

                elif msg == WM_CTLCOLORBTN:
                    gdi32.SetDCBrushColor(wparam, BG_COLOR_DARK)
                    return gdi32.GetStockObject(DC_BRUSH)

                elif msg == WM_CTLCOLOREDIT or msg == WM_CTLCOLORLISTBOX:
                    gdi32.SetTextColor(wparam, TEXT_COLOR_DARK)
                    gdi32.SetBkColor(wparam, CONTROL_BG_COLOR_DARK)
                    gdi32.SetDCBrushColor(wparam, CONTROL_BG_COLOR_DARK)
                    return gdi32.GetStockObject(DC_BRUSH)

            return dialog_proc_callback(hwnd, msg, wparam, lparam)

        self.__dialogproc = WNDPROC(_dialog_proc_callback)

    def _show_async(self, is_dark=False):
        self.__is_async = True
        self.is_dark = self.parent_window.is_dark
        self.hwnd = user32.CreateDialogIndirectParamW(
                0,
                byref(self.__dialog_data),
                self.parent_window.hwnd,
                self.__dialogproc,
                1
                )
        user32.ShowWindow(self.hwnd, SW_SHOW)

    def _show_sync(self, is_dark=False, lparam=0):
        self.__is_async = False
        self.is_dark = self.parent_window.is_dark
        res = user32.DialogBoxIndirectParamW(
            0,
            byref(self.__dialog_data),
            self.parent_window.hwnd,
            self.__dialogproc,
            lparam
        )
        self.hwnd = None
        self.children = []
        return res

    def apply_theme(self, is_dark):
        self.is_dark = is_dark
        if self.hwnd:
            user32.SendMessageW(self.hwnd, WM_THEMECHANGED, 0, 0)
