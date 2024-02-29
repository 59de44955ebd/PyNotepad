from ctypes import Structure, create_unicode_buffer, c_voidp, windll, cast, byref, sizeof, c_wchar_p, c_ubyte, POINTER
from ctypes.wintypes import SHORT, WORD, DWORD, HWND, HINSTANCE, LPWSTR, LPCWSTR, LPVOID, HANDLE, INT, WCHAR, BYTE, COLORREF, HDC, UINT, WPARAM, LPARAM, LONG
from winapp.wintypes_extended import WINFUNCTYPE, UINT_PTR
from winapp.dlls import shell32
from winapp.const import *

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

    def __str__(self):
        return  "('%s' %d)" % (self.lfFaceName, self.lfHeight)

    def __repr__(self):
        return "<LOGFONTW '%s' %d>" % (self.lfFaceName, self.lfHeight)

class CHOOSEFONTW(Structure):
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

    def __init__(self, *args, **kwargs):
        super(CHOOSEFONTW, self).__init__(*args, **kwargs)
        self.lStructSize = sizeof(self)

# modern icon (flat)
def get_stock_icon(siid):
    sii = SHSTOCKICONINFO()
    SHGSI_ICON = 0x000000100
    shell32.SHGetStockIconInfo(siid, SHGSI_ICON, byref(sii))
    return sii.hIcon


class DialogData(object):

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
                0,  # menu, 0x0000 for none
                #0,  # windowClass, 0x0000 for none
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
