import json
import os
import re
import sys
import time

from winapp.mainwin import *
from winapp.dialog import *
from winapp.dlls import *
from winapp.controls.edit import *
from winapp.controls.statusbar import *

from platform import win32_ver
WIN_VERSION = win32_ver()[0]
DARK_SUPPORTED = float(WIN_VERSION) >= 10

import locale
locale.setlocale(locale.LC_TIME, '')  # use system locale for formatting date/time

from resources.const import *

APP_NAME = 'PyNotepad'
APP_VERSION = 1
APP_COPYRIGHT = '2024 github.com/59de44955ebd'
APP_DIR = os.path.dirname(os.path.abspath(__file__))

LANG = locale.windows_locale[kernel32.GetUserDefaultUILanguage()]
if not os.path.isdir(os.path.join(APP_DIR, 'resources', LANG)):
    LANG = 'en_US'

with open(os.path.join(APP_DIR, 'resources', LANG, 'StringTable1.json'), 'rb') as f:
    __ = json.loads(f.read())

def _(s):
    return __[s] if s in __ else s

ENCODINGS = {
        IDM_ANSI:       'ansi',
        IDM_UTF_16_LE:  'utf-16-le',
        IDM_UTF_16_BE:  'utf-16-be',
        IDM_UTF_8:      'utf-8',
        IDM_UTF_8_BOM:  'utf-8-sig',
        }

EOL_MODES = {
        IDM_EOL_CRLF:   '\r\n',
        IDM_EOL_LF:     '\n',
        IDM_EOL_CR:     '\r',
        }

EDIT_MAX_TEXT_LEN = 0x80000  # 512 KB (Edit control's default: 30.000)

STATUSBAR_PART_CARET = 1
STATUSBAR_PART_ZOOM = 2
STATUSBAR_PART_EOL = 3
STATUSBAR_PART_ENCODING = 4


class App(MainWin):

    def __init__(self, args=[]):
        # defaults
        self._show_statusbar = True
        self._dark_mode = reg_should_use_dark_mode() if DARK_SUPPORTED else False
        self._word_wrap = True

        self._font = ['Consolas', 11, 400, FALSE]
        self._search_term = ''
        self._saved_search_term = ''
        self._replace_term = ''
        self._match_case = FALSE
        self._wrap_arround = FALSE
        self._search_up = FALSE

        left, top, width, height = self._load_state()

        self._zoom = 100
        self._filename = None
        self._is_dirty = False
        self._saved_text_len = 1
        self._saved_text = ''
        self._eol_mode_id = IDM_EOL_CRLF
        self._encoding_id = IDM_UTF_8

        # load menu resource
        with open(os.path.join(APP_DIR, 'resources', LANG, 'Menu1.json'), 'rb') as f:
            menu_data = json.loads(f.read())

        self.COMMAND_MESSAGE_MAP = {
            IDM_NEW:                  self.action_new,
            IDM_NEW_WINDOW:           self.action_new_window,
            IDM_OPEN:                 self.action_open,
            IDM_SAVE:                 self.action_save,
            IDM_SAVE_AS:              self.action_save_as,
            IDM_PRINT:                self.action_print,
            IDM_EXIT:                 self.action_exit,
            IDM_UNDO:                 self.action_undo,
            IDM_CUT:                  self.action_cut,
            IDM_COPY:                 self.action_copy,
            IDM_PASTE:                self.action_paste,
            IDM_DELETE:               self.action_delete,
            IDM_FIND:                 self.action_find,
            IDM_FIND_NEXT:            self.action_find_next,
            IDM_FIND_PREVIOUS:        self.action_find_previous,
            IDM_REPLACE:              self.action_replace,
            IDM_GO_TO:                self.action_go_to,
            IDM_SELECT_ALL:           self.action_select_all,
            IDM_TIME_DATE:            self.action_time_date,
            IDM_WORD_WRAP:            self.action_word_wrap,
            IDM_FONT:                 self.action_font,
            IDM_ZOOM_IN:              self.action_zoom_in,
            IDM_ZOOM_OUT:             self.action_zoom_out,
            IDM_RESTORE_DEFAULT_ZOOM: self.action_restore_default_zoom,
            IDM_STATUS_BAR:           self.action_status_bar,
            IDM_DARK_MODE:            self.action_dark_mode,
            IDM_ABOUT_NOTEPAD:        self.action_about_notepad,
            IDM_ABOUT_WINDOWS:        self.action_about_windows,
        }

        for item_id in ENCODINGS.keys():
            self.COMMAND_MESSAGE_MAP[item_id] = lambda id=item_id: self.action_encoding(id)

        for item_id in EOL_MODES.keys():
            self.COMMAND_MESSAGE_MAP[item_id] = lambda id=item_id: self.action_eol_mode(id)

        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            hicon = user32.LoadIconW(kernel32.GetModuleHandleW(None), MAKEINTRESOURCEW(IDI_APPICON))
        else:
            hicon = user32.LoadImageW(0, os.path.join(APP_DIR, 'app.ico'), IMAGE_ICON, 16, 16, LR_LOADFROMFILE)

        # create main window
        super().__init__(
                self._get_caption(),
                left=left, top=top, width=width, height=height,
                menu_data=menu_data,
                menu_mod_translation_table=__,
                hicon=hicon
                )

        if self._word_wrap:
            user32.CheckMenuItem(self.hmenu, IDM_WORD_WRAP, MF_BYCOMMAND | MF_CHECKED)
            user32.EnableMenuItem(self.hmenu, IDM_GO_TO, MF_BYCOMMAND | MF_GRAYED)

        self._create_statusbar()
        self._create_dialogs()
        self._create_edit()

        def _on_WM_SIZE(hwnd, wparam, lparam):
            width, height = lparam & 0xFFFF, (lparam >> 16) & 0xFFFF
            self.statusbar.update_size()  # Reposition and resize the statusbar
            self.statusbar.right_align_parts(width)  # keep statusbar parts right aligned
            user32.SetWindowPos(self.edit.hwnd, 0, 0, 0, width,
                    height - self.statusbar.height if self._show_statusbar else height, 0)
        self.register_message_callback(WM_SIZE, _on_WM_SIZE)

        def _on_WM_COMMAND(hwnd, wparam, lparam):
            command = HIWORD(wparam)
            if lparam == 0:
                self.COMMAND_MESSAGE_MAP[LOWORD(wparam)]()
            elif lparam == self.edit.hwnd and command == EN_CHANGE:
                is_dirty = self._check_if_text_changed()
                if is_dirty != self._is_dirty:
                    self._is_dirty = is_dirty
                    self.set_window_text(self._get_caption())
            return FALSE
        self.register_message_callback(WM_COMMAND, _on_WM_COMMAND)

        def _on_WM_DROPFILES(hwnd, wparam, lparam):
            dropped_items = self.get_dropped_items(wparam)
            if os.path.isfile(dropped_items[0]) and self._handle_dirty():
                self._load_file(dropped_items[0])
        self.register_message_callback(WM_DROPFILES, _on_WM_DROPFILES)

        self.register_message_callback(WM_CLOSE, self.action_exit, True)

        if DARK_SUPPORTED:
            def _on_WM_SETTINGCHANGE(hwnd, wparam, lparam):
                if lparam and cast(lparam, LPCWSTR).value == 'ImmersiveColorSet':
                    if reg_should_use_dark_mode() != self._dark_mode:
                        self.action_dark_mode()
            self.register_message_callback(WM_SETTINGCHANGE, _on_WM_SETTINGCHANGE)

            if self._dark_mode:
                self.apply_theme(True)
                user32.CheckMenuItem(self.hmenu, IDM_DARK_MODE, MF_BYCOMMAND | MF_CHECKED)
        else:
            user32.EnableMenuItem(self.hmenu, IDM_DARK_MODE, MF_BYCOMMAND | MF_GRAYED)

        if len(args) > 0:
            self._load_file(args[0])

        self.show()
        user32.SetFocus(self.edit.hwnd)

    ########################################
    #
    ########################################
    def _create_dialogs(self):
        with open(os.path.join(APP_DIR, 'resources', LANG, 'Dialog1540.json'), 'rb') as f:
            dialog_dict = json.loads(f.read())

        def _dialog_proc_find(hwnd, msg, wparam, lparam):

            if msg == WM_INITDIALOG:
                hwnd_edit = user32.GetDlgItem(hwnd, ID_EDIT_FIND)

                # Limit search input to 127 chars
                user32.SendMessageW(hwnd_edit, EM_SETLIMITTEXT, 127, 0)

                # check if something is selected
                pos_start, pos_end = DWORD(), DWORD()
                user32.SendMessageW(self.edit.hwnd, EM_GETSEL, byref(pos_start), byref(pos_end))
                if pos_end.value > pos_start.value:
                    text_len = user32.SendMessageW(self.edit.hwnd, WM_GETTEXTLENGTH, 0, 0) + 1
                    text_buf = create_unicode_buffer(text_len)
                    user32.SendMessageW(self.edit.hwnd, WM_GETTEXT, text_len, text_buf)
                    self._search_term = text_buf.value[pos_start.value:pos_end.value][:127]
                elif self._search_term == '':
                    self._search_term = self._saved_search_term

                # update button states
                if self._search_term:
                    user32.SendMessageW(hwnd_edit, WM_SETTEXT, 0, create_unicode_buffer(self._search_term))
                else:
                    user32.EnableWindow(user32.GetDlgItem(hwnd, ID_FIND_NEXT), 0)
                if self._search_up:
                    user32.SendMessageW(user32.GetDlgItem(hwnd, ID_UP), BM_SETCHECK, BST_CHECKED, 0)
                else:
                    user32.SendMessageW(user32.GetDlgItem(hwnd, ID_DOWN), BM_SETCHECK, BST_CHECKED, 0)
                if self._match_case:
                    user32.SendMessageW(user32.GetDlgItem(hwnd, ID_MATCH_CASE), BM_SETCHECK, BST_CHECKED, 0)
                if self._wrap_arround:
                    user32.SendMessageW(user32.GetDlgItem(hwnd, ID_WRAP_AROUND), BM_SETCHECK, BST_CHECKED, 0)

            elif msg == WM_COMMAND:
                control_id = LOWORD(wparam)
                command = HIWORD(wparam)

                if control_id == ID_EDIT_FIND:
                    if command == EN_UPDATE:
                        text_len = user32.SendMessageW(user32.GetDlgItem(hwnd, ID_EDIT_FIND), WM_GETTEXTLENGTH, 0, 0)
                        user32.EnableWindow(user32.GetDlgItem(hwnd, ID_FIND_NEXT), int(text_len > 0))

                elif command == BN_CLICKED:
                    if control_id == ID_FIND_NEXT:
                        self._match_case = user32.SendMessageW(user32.GetDlgItem(hwnd, ID_MATCH_CASE), BM_GETCHECK, 0, 0)
                        self._wrap_arround = user32.SendMessageW(user32.GetDlgItem(hwnd, ID_WRAP_AROUND), BM_GETCHECK, 0, 0)
                        self._search_up = user32.SendMessageW(user32.GetDlgItem(hwnd, ID_UP), BM_GETCHECK, 0, 0)
                        hwnd_edit = user32.GetDlgItem(hwnd, ID_EDIT_FIND)
                        text_len = user32.SendMessageW(hwnd_edit, WM_GETTEXTLENGTH, 0, 0) + 1
                        text_buf = create_unicode_buffer(text_len)
                        user32.SendMessageW(hwnd_edit, WM_GETTEXT, text_len, text_buf)
                        self._search_term = text_buf.value
                        self._find()

                    elif control_id == ID_CANCEL:
                        user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)

            elif msg == WM_CLOSE:
                user32.SetFocus(self.edit.hwnd)

            return FALSE

        self.dialog_find = Dialog(self, dialog_dict, _dialog_proc_find)

        with open(os.path.join(APP_DIR, 'resources', LANG, 'Dialog1541.json'), 'rb') as f:
            dialog_dict = json.loads(f.read())

        def _dialog_proc_replace(hwnd, msg, wparam, lparam):
            if msg == WM_INITDIALOG:
                # Limit search and replace input to 127 chars
                hwnd_search_edit = user32.GetDlgItem(hwnd, ID_EDIT_FIND)
                user32.SendMessageW(hwnd_search_edit, EM_SETLIMITTEXT, 127, 0)

                hwnd_replace_edit = user32.GetDlgItem(hwnd, ID_EDIT_REPLACE)
                user32.SendMessageW(hwnd_replace_edit, EM_SETLIMITTEXT, 127, 0)

                # check if something is selected
                pos_start, pos_end = DWORD(), DWORD()
                user32.SendMessageW(self.edit.hwnd, EM_GETSEL, byref(pos_start), byref(pos_end))
                if pos_end.value > pos_start.value:
                    text_len = user32.SendMessageW(self.edit.hwnd, WM_GETTEXTLENGTH, 0, 0) + 1
                    text_buf = create_unicode_buffer(text_len)
                    user32.SendMessageW(self.edit.hwnd, WM_GETTEXT, text_len, text_buf)
                    self._search_term = text_buf.value[pos_start.value:pos_end.value][:127]
                # update button states
                if self._search_term:
                    user32.SendMessageW(hwnd_search_edit, WM_SETTEXT, 0, create_unicode_buffer(self._search_term))
                else:
                    user32.EnableWindow(user32.GetDlgItem(hwnd, ID_FIND_NEXT), 0)
                    user32.EnableWindow(user32.GetDlgItem(hwnd, ID_REPLACE), 0)
                    user32.EnableWindow(user32.GetDlgItem(hwnd, ID_REPLACE_ALL), 0)
                if self._replace_term:
                    user32.SendMessageW(hwnd_replace_edit, WM_SETTEXT, 0, create_unicode_buffer(self._search_term))
                if self._match_case:
                    user32.SendMessageW(user32.GetDlgItem(hwnd, ID_MATCH_CASE), BM_SETCHECK, BST_CHECKED, 0)
                if self._wrap_arround:
                    user32.SendMessageW(user32.GetDlgItem(hwnd, ID_WRAP_AROUND), BM_SETCHECK, BST_CHECKED, 0)

            elif msg == WM_COMMAND:
                control_id = LOWORD(wparam)
                command = HIWORD(wparam)

                if control_id == ID_EDIT_FIND:
                    if command == EN_UPDATE:
                        text_len = user32.SendMessageW(user32.GetDlgItem(hwnd, ID_EDIT_FIND), WM_GETTEXTLENGTH, 0, 0)
                        user32.EnableWindow(user32.GetDlgItem(hwnd, ID_FIND_NEXT), int(text_len > 0))
                        user32.EnableWindow(user32.GetDlgItem(hwnd, ID_REPLACE), int(text_len > 0))
                        user32.EnableWindow(user32.GetDlgItem(hwnd, ID_REPLACE_ALL), int(text_len > 0))

                elif command == BN_CLICKED:
                    if control_id in (ID_FIND_NEXT, ID_REPLACE, ID_REPLACE_ALL):
                        self._match_case = user32.SendMessageW(user32.GetDlgItem(hwnd, ID_MATCH_CASE), BM_GETCHECK, 0, 0)
                        self._wrap_arround = user32.SendMessageW(user32.GetDlgItem(hwnd, ID_WRAP_AROUND), BM_GETCHECK, 0, 0)

                        hwnd_search_edit = user32.GetDlgItem(hwnd, ID_EDIT_FIND)
                        text_len = user32.SendMessageW(hwnd_search_edit, WM_GETTEXTLENGTH, 0, 0) + 1
                        text_buf = create_unicode_buffer(text_len)
                        user32.SendMessageW(hwnd_search_edit, WM_GETTEXT, text_len, text_buf)
                        self._search_term = text_buf.value

                        hwnd_replace_edit = user32.GetDlgItem(hwnd, ID_EDIT_REPLACE)
                        text_len = user32.SendMessageW(hwnd_replace_edit, WM_GETTEXTLENGTH, 0, 0) + 1
                        text_buf = create_unicode_buffer(text_len)
                        user32.SendMessageW(hwnd_replace_edit, WM_GETTEXT, text_len, text_buf)
                        self._replace_term = text_buf.value

                        if control_id == ID_FIND_NEXT:
                            self._find()
                        elif control_id == ID_REPLACE:
                            self._replace()
                        elif control_id == ID_REPLACE_ALL:
                            self._replace_all()

                    elif control_id == ID_CANCEL:
                        user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)

            elif msg == WM_CLOSE:
                user32.SetFocus(self.edit.hwnd)

            return FALSE

        self.dialog_replace = Dialog(self, dialog_dict, _dialog_proc_replace)

        with open(os.path.join(APP_DIR, 'resources', LANG, 'Dialog14.json'), 'rb') as f:
            dialog_dict = json.loads(f.read())

        def _dialog_proc_goto(hwnd, msg, wparam, lparam):
            if msg == WM_INITDIALOG:
                hwnd_edit = user32.GetDlgItem(hwnd, ID_EDIT_GOTO)
                line_idx = user32.SendMessageW(self.edit.hwnd, EM_LINEFROMCHAR, -1, 0)
                user32.SendMessageW(hwnd_edit, WM_SETTEXT, 0,
                        create_unicode_buffer(str(line_idx + 1)))
                user32.SendMessageW(hwnd_edit, EM_SETSEL, 0, -1)

            elif msg == WM_COMMAND:
                control_id = LOWORD(wparam)
                command = HIWORD(wparam)
                if command == BN_CLICKED:
                    if control_id == ID_FIND_NEXT:
                        hwnd_edit = user32.GetDlgItem(hwnd, ID_EDIT_GOTO)
                        text_len = user32.SendMessageW(hwnd_edit, WM_GETTEXTLENGTH, 0, 0) + 1
                        if text_len > 1:
                            text_buf = create_unicode_buffer(text_len)
                            user32.SendMessageW(hwnd_edit, WM_GETTEXT, text_len, text_buf)
                            user32.EndDialog(hwnd, int(text_buf.value))
                        else:
                            user32.EndDialog(hwnd, 0)
                    elif control_id == ID_CANCEL:
                        user32.EndDialog(hwnd, 0)

            return FALSE

        self.dialog_goto = Dialog(self, dialog_dict, _dialog_proc_goto)

    ########################################
    #
    ########################################
    def _create_edit(self):
        self.edit = Edit(
                self,
                style=WS_VISIBLE | WS_CHILD | ES_MULTILINE | WS_TABSTOP | WS_VSCROLL |
                ES_AUTOVSCROLL | ES_NOHIDESEL | (0 if self._word_wrap else WS_HSCROLL)
                )
        self.edit.set_font(*self._font)
        user32.SendMessageW(self.edit.hwnd, EM_SETLIMITTEXT, EDIT_MAX_TEXT_LEN, 0)

        def _on_WM_KEYUP(hwnd, wparam, lparam):
            self._check_caret_pos()
            self._check_if_text_selected()
        self.edit.register_message_callback(WM_KEYUP, _on_WM_KEYUP)

        def _on_WM_MOUSEMOVE(hwnd, wparam, lparam):
            if wparam & MK_LBUTTON:
                pos = user32.SendMessageW(self.edit.hwnd, EM_CHARFROMPOS, 0, lparam)
                if pos > -1:
                    char_idx, line_idx = LOWORD(pos), HIWORD(pos)
                    line_char_idx = user32.SendMessageW(self.edit.hwnd, EM_LINEINDEX, -1, 0)
                    self._show_caret_pos(line_idx, char_idx - line_char_idx)
        self.edit.register_message_callback(WM_MOUSEMOVE, _on_WM_MOUSEMOVE)

        def _on_WM_LBUTTONUP(hwnd, wparam, lparam):
            self._check_caret_pos()
            self._check_if_text_selected()
        self.edit.register_message_callback(WM_LBUTTONUP, _on_WM_LBUTTONUP)

    ########################################
    #
    ########################################
    def _create_statusbar(self):
        self.statusbar = StatusBar(self, parts=(0, 140, 50, 120, 120), parts_right_aligned=True)
        self._show_caret_pos()
        self.statusbar.set_text('100%', STATUSBAR_PART_ZOOM)

        buf = create_unicode_buffer(24)
        user32.GetMenuStringW(self.hmenu, self._eol_mode_id, buf, 24, MF_BYCOMMAND)
        self.statusbar.set_text(buf.value, STATUSBAR_PART_EOL)

        user32.GetMenuStringW(self.hmenu, self._encoding_id, buf, 24, MF_BYCOMMAND)
        self.statusbar.set_text(buf.value, STATUSBAR_PART_ENCODING)

        if self._show_statusbar:
            user32.CheckMenuItem(self.hmenu, IDM_STATUS_BAR, MF_BYCOMMAND | MF_CHECKED)
        else:
            self.statusbar.show(SW_HIDE)

    ########################################
    # Load saved state from registry
    ########################################
    def _load_state(self):
        left, top, width, height=CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT
        hkey = HKEY()
        if advapi32.RegOpenKeyW(HKEY_CURRENT_USER, f'Software\\{APP_NAME}' , byref(hkey)) == ERROR_SUCCESS:
            data = (BYTE * sizeof(DWORD))()
            cbData = DWORD(sizeof(data))
            if advapi32.RegQueryValueExW(hkey, 'iWindowPosX', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                left = cast(data, POINTER(DWORD)).contents.value
            if advapi32.RegQueryValueExW(hkey, 'iWindowPosY', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                top = cast(data, POINTER(DWORD)).contents.value
            if advapi32.RegQueryValueExW(hkey, 'iWindowPosDX', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                width = cast(data, POINTER(DWORD)).contents.value
            if advapi32.RegQueryValueExW(hkey, 'iWindowPosDY', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                height = cast(data, POINTER(DWORD)).contents.value
            if advapi32.RegQueryValueExW(hkey, 'StatusBar', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                self._show_statusbar = cast(data, POINTER(DWORD)).contents.value == 1
            # Logic: if system uses dark mode, always use dark mode as well, otherwise use saved setting
            if DARK_SUPPORTED and not self._dark_mode and advapi32.RegQueryValueExW(hkey, 'DarkMode', None, None,
                    byref(data), byref(cbData)) == ERROR_SUCCESS:
                self._dark_mode = cast(data, POINTER(DWORD)).contents.value == 1

            if advapi32.RegQueryValueExW(hkey, 'iPointSize', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                self._font[1] = cast(data, POINTER(DWORD)).contents.value
            if advapi32.RegQueryValueExW(hkey, 'lfWeight', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                self._font[2] = cast(data, POINTER(DWORD)).contents.value
            if advapi32.RegQueryValueExW(hkey, 'lfItalic', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                self._font[3] = cast(data, POINTER(DWORD)).contents.value  #== 1

            if advapi32.RegQueryValueExW(hkey, 'fWrap', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                self._word_wrap = cast(data, POINTER(DWORD)).contents.value == 1

            if advapi32.RegQueryValueExW(hkey, 'fMatchCase', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                self._match_case = cast(data, POINTER(DWORD)).contents.value == 1
            if advapi32.RegQueryValueExW(hkey, 'fWrapAround', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                self._wrap_arround = cast(data, POINTER(DWORD)).contents.value == 1
            if advapi32.RegQueryValueExW(hkey, 'fReverse', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                self._search_up = cast(data, POINTER(DWORD)).contents.value == 1
            data = (BYTE * 128)()
            cbData = DWORD(sizeof(data))
            if advapi32.RegQueryValueExW(hkey, 'lfFaceName', None, None, data, byref(cbData)) == ERROR_SUCCESS:
                self._font[0] = cast(data, LPWSTR).value
            if advapi32.RegQueryValueExW(hkey, 'searchString', None, None, data, byref(cbData)) == ERROR_SUCCESS:
                self._search_term = cast(data, LPWSTR).value
            if advapi32.RegQueryValueExW(hkey, 'replaceString', None, None, data, byref(cbData)) == ERROR_SUCCESS:
                self._replace_term = cast(data, LPWSTR).value
        else:
            advapi32.RegCreateKeyW(HKEY_CURRENT_USER, f'Software\\{APP_NAME}' , byref(hkey))
        advapi32.RegCloseKey(hkey)
        return left, top, width, height

    ########################################
    # Save state to registry
    ########################################
    def _save_state(self):
        hkey = HKEY()
        if advapi32.RegOpenKeyW(HKEY_CURRENT_USER, f'Software\\{APP_NAME}' , byref(hkey)) == ERROR_SUCCESS:
            dwsize = sizeof(DWORD)
            # font
            buf = create_unicode_buffer(self._font[0])
            advapi32.RegSetValueExW(hkey, 'lfFaceName', 0, REG_SZ, buf, sizeof(buf))
            advapi32.RegSetValueExW(hkey, 'iPointSize', 0, REG_DWORD, byref(DWORD(self._font[1])), dwsize)
            advapi32.RegSetValueExW(hkey, 'lfWeight', 0, REG_DWORD, byref(DWORD(int(self._font[2]))), dwsize)
            advapi32.RegSetValueExW(hkey, 'lfItalic', 0, REG_DWORD, byref(DWORD(int(self._font[3]))), dwsize)

            advapi32.RegSetValueExW(hkey, 'fWrap', 0, REG_DWORD, byref(DWORD(int(self._word_wrap))), dwsize)
            # search
            buf = create_unicode_buffer(self._saved_search_term)
            advapi32.RegSetValueExW(hkey, 'searchString', 0, REG_SZ, buf, sizeof(buf))
            buf = create_unicode_buffer(self._replace_term)
            advapi32.RegSetValueExW(hkey, 'replaceString', 0, REG_SZ, buf, sizeof(buf))
            advapi32.RegSetValueExW(hkey, 'fMatchCase', 0, REG_DWORD, byref(DWORD(int(self._match_case))), dwsize)
            advapi32.RegSetValueExW(hkey, 'fWrapAround', 0, REG_DWORD, byref(DWORD(int(self._wrap_arround))), dwsize)
            advapi32.RegSetValueExW(hkey, 'fReverse', 0, REG_DWORD, byref(DWORD(int(self._search_up))), dwsize)
            # window
            advapi32.RegSetValueExW(hkey, 'DarkMode', 0, REG_DWORD, byref(DWORD(int(self._dark_mode))), dwsize)
            advapi32.RegSetValueExW(hkey, 'StatusBar', 0, REG_DWORD, byref(DWORD(int(self._show_statusbar))), dwsize)
            self.show(SW_RESTORE)
            rc = self.get_window_rect()
            advapi32.RegSetValueExW(hkey, 'iWindowPosX', 0, REG_DWORD, byref(DWORD(rc.left)), dwsize)
            advapi32.RegSetValueExW(hkey, 'iWindowPosY', 0, REG_DWORD, byref(DWORD(rc.top)), dwsize)
            advapi32.RegSetValueExW(hkey, 'iWindowPosDX', 0, REG_DWORD, byref(DWORD(rc.right - rc.left)), dwsize)
            advapi32.RegSetValueExW(hkey, 'iWindowPosDY', 0, REG_DWORD, byref(DWORD(rc.bottom - rc.top)), dwsize)

            advapi32.RegCloseKey(hkey)

    ########################################
    #
    ########################################
    def _find(self, search_up=None):
        sel_start_pos = DWORD()
        sel_end_pos = DWORD()
        user32.SendMessageW(self.edit.hwnd, EM_GETSEL, byref(sel_start_pos), byref(sel_end_pos))

        text_len = user32.SendMessageW(self.edit.hwnd, WM_GETTEXTLENGTH, 0, 0) + 1
        text_buf = create_unicode_buffer(text_len)
        user32.SendMessageW(self.edit.hwnd, WM_GETTEXT, text_len, text_buf)
        txt = text_buf.value

        haystack = txt if self._match_case else txt.lower()
        needle = self._search_term if self._match_case else self._search_term.lower()

        if search_up is None:
            search_up = self._search_up

        if search_up:
            pos = haystack.rfind(needle, 0, sel_start_pos.value)
        else:
            pos = haystack.find(needle, sel_end_pos.value)

        if pos < 0 and self._wrap_arround:
            if search_up:
                pos = haystack.rfind(needle, sel_start_pos.value)
                self.statusbar.set_text('Found next from the bottom' if pos > -1 else '')
            else:
                pos = haystack.find(needle, 0, sel_end_pos.value)
                self.statusbar.set_text('Found next from the top' if pos > -1 else '')
        else:
            self.statusbar.set_text()

        if pos > -1:
            user32.SendMessageW(self.edit.hwnd, EM_SETSEL, pos, pos + len(self._search_term))
            self._check_caret_pos()
            return True
        else:
             self.show_message_box(_("CANNOT_FIND").format(self._search_term), APP_NAME)
             return False

    ########################################
    #
    ########################################
    def _replace(self):
        if self._find():
            user32.SendMessageW(self.edit.hwnd, EM_REPLACESEL, 1, create_unicode_buffer(self._replace_term))

    ########################################
    #
    ########################################
    def _replace_all(self):
        text_len = user32.SendMessageW(self.edit.hwnd, WM_GETTEXTLENGTH, 0, 0) + 1
        text_buf = create_unicode_buffer(text_len)
        user32.SendMessageW(self.edit.hwnd, WM_GETTEXT, text_len, text_buf)
        txt = text_buf.value
        if self._match_case:
            txt_new = txt.replace(self._search_term, self._replace_term)
        else:
            txt_new = re.compile(re.escape(self._search_term), re.IGNORECASE).sub(self._replace_term, txt)
        if txt_new != txt:
            user32.SendMessageW(self.edit.hwnd, EM_SETSEL, 0, -1)
            user32.SendMessageW(self.edit.hwnd, EM_REPLACESEL, 1, create_unicode_buffer(txt_new))
            self._check_caret_pos()

    ########################################
    # Tries to detect the EOL mode of the specified bytes
    ########################################
    def _detect_eol(self, data):
        if b'\n' in data and not b'\r' in data:
            return IDM_EOL_LF
        if b'\r' in data and not b'\n' in data:
            return IDM_EOL_CR
        return IDM_EOL_CRLF

    ########################################
    # Tries to detect the encoding of the specified bytes
    ########################################
    def _detect_encoding(self, data):
        if len(data) == 0:
            return IDM_UTF_8
        bom = self._get_bom(data)
        if bom:
            return bom
        if data[0] == 0:
            return IDM_UTF_16_BE
        if len(data) > 1 and data[1] == 0:
            return IDM_UTF_16_LE
        if self._is_utf8(data):
            return IDM_UTF_8
        return IDM_ANSI

    ########################################
    # Checks bytes for Byte-Order-Mark (BOM), returns BOM type or None
    ########################################
    def _get_bom(self, data):
    	if len(data) >= 3 and data[:3] == b'\xEF\xBB\xBF':
    		return IDM_UTF_8_BOM
    	elif len(data) >= 2:
            if data[:2] == b'\xFF\xFE':
                return IDM_UTF_16_LE
            elif data[:2] == b'\xFE\xFF':
                return IDM_UTF_16_BE

    ########################################
    # Checks bytes for invalid UTF-8 sequences
    # Notice: since ASCII is a UTF-8 subset, function also returns True for pure ASCII data
    ########################################
    def _is_utf8(self, data):
    	data_len = len(data)
    	i = -1
    	while True:
    		i += 1
    		if i >= data_len:
    		    break
    		o = data[i]
    		if (o < 128):
    		    continue
    		elif o & 224 == 192 and o > 193:
    		    n = 1
    		elif o & 240 == 224:
    		    n = 2
    		elif o & 248 == 240 and o < 245:
    		    n = 3
    		else:
    		    return False # invalid UTF-8 sequence
    		for c in range(n):
    			i += 1
    			if i > data_len:
    			    return False # invalid UTF-8 sequence
    			if data[i] & 192 != 128:
    			    return False # invalid UTF-8 sequence
    	return True

    ########################################
    #
    ########################################
    def _check_caret_pos(self):
        pt = POINT()
        ok = user32.GetCaretPos(byref(pt))
        pos = user32.SendMessageW(self.edit.hwnd, EM_CHARFROMPOS, 0, MAKELONG(pt.x, pt.y))
        if pos > -1:
            char_idx, line_idx = LOWORD(pos), HIWORD(pos)
            line_char_idx = user32.SendMessageW(self.edit.hwnd, EM_LINEINDEX, -1, 0)
            self._show_caret_pos(line_idx, char_idx - line_char_idx)

    ########################################
    #
    ########################################
    def _show_caret_pos(self, line_idx=0, col_idx=0):
        self.statusbar.set_text(_('LINE_COLUMN').format(line_idx + 1, col_idx + 1), STATUSBAR_PART_CARET)

    ########################################
    #
    ########################################
    def _load_file(self, filename):
        if os.stat(filename).st_size > EDIT_MAX_TEXT_LEN:
            return self.show_message_box(_('FILE_TOO_BIG'), APP_NAME, MB_ICONWARNING | MB_OK)
        with open(filename, 'rb') as f:
            data = f.read()
        self._show_caret_pos()
        buf = create_unicode_buffer(24)
        eol_mode_id = self._detect_eol(data)
        if eol_mode_id != self._eol_mode_id:
            user32.CheckMenuItem(self.hmenu, self._eol_mode_id, MF_BYCOMMAND | MF_UNCHECKED)
            self._eol_mode_id = eol_mode_id
            user32.CheckMenuItem(self.hmenu, self._eol_mode_id, MF_BYCOMMAND | MF_CHECKED)
        user32.GetMenuStringW(self.hmenu, self._eol_mode_id, buf, 24, MF_BYCOMMAND)
        self.statusbar.set_text(buf.value, STATUSBAR_PART_EOL)
        encoding_id = self._detect_encoding(data)
        if encoding_id != self._encoding_id:
            user32.CheckMenuItem(self.hmenu, self._encoding_id, MF_BYCOMMAND | MF_UNCHECKED)
            self._encoding_id = encoding_id
            user32.CheckMenuItem(self.hmenu, self._encoding_id, MF_BYCOMMAND | MF_CHECKED)
        user32.GetMenuStringW(self.hmenu, self._encoding_id, buf, 24, MF_BYCOMMAND)
        self.statusbar.set_text(buf.value, STATUSBAR_PART_ENCODING)
        if self._eol_mode_id != IDM_EOL_CRLF:
            data = data.replace(EOL_MODES[self._eol_mode_id].encode(), b'\r\n')
        text_buf = create_unicode_buffer(data.decode(ENCODINGS[self._encoding_id]))
        user32.SendMessageW(self.edit.hwnd, WM_SETTEXT, 0, text_buf)
        self._saved_text_len = user32.SendMessageW(self.edit.hwnd, WM_GETTEXTLENGTH, 0, 0) + 1
        text_buf = create_unicode_buffer(self._saved_text_len)
        user32.SendMessageW(self.edit.hwnd, WM_GETTEXT, self._saved_text_len, text_buf)
        flag = MF_BYCOMMAND | (MF_ENABLED if self._saved_text_len > 1 else MF_GRAYED)
        for item_id in (IDM_FIND, IDM_FIND_NEXT, IDM_FIND_PREVIOUS, IDM_REPLACE):
            user32.EnableMenuItem(self.hmenu, item_id, flag)
        self._saved_text = text_buf.value
        self._is_dirty = False
        self._filename = os.path.basename(filename)
        self.set_window_text(self._get_caption())
        user32.SetFocus(self.edit.hwnd)

    ########################################
    #
    ########################################
    def _save_file(self, filename):
        text_len = user32.SendMessageW(self.edit.hwnd, WM_GETTEXTLENGTH, 0, 0) + 1
        text_buf = create_unicode_buffer(text_len)
        user32.SendMessageW(self.edit.hwnd, WM_GETTEXT, text_len, text_buf)
        txt = text_buf.value
        if self._eol_mode_id != IDM_EOL_CRLF:
            txt = txt.replace('\r\n', EOL_MODES[self._eol_mode_id])
        try:
            with open(filename, 'wb') as f:
                f.write(txt.encode(ENCODINGS[self._encoding_id]))
            self._saved_text_len = text_len
            self._saved_text = text_buf.value
            return True
        except Exception as e:
            self.statusbar.set_text(f'{e.strerror}: {e.filename}')
            return False

    ########################################
    #
    ########################################
    def _handle_dirty(self):
        if not self._is_dirty:
            return True
        res = self.show_message_box(
                _('SAVE_CHANGES').format(self._filename if self._filename else _('Untitled')),
                APP_NAME, MB_YESNOCANCEL)
        if res == IDCANCEL:
            return False
        elif res == IDNO:
            return True
        elif res == IDYES:
            return self.action_save()

    ########################################
    #
    ########################################
    def _get_caption(self):
        return ('*' if self._is_dirty else '') + (self._filename if self._filename else _('Untitled')) + ' - ' + APP_NAME

    ########################################
    #
    ########################################
    def _check_if_text_changed(self):
        text_len = user32.SendMessageW(self.edit.hwnd, WM_GETTEXTLENGTH, 0, 0) + 1
        # enable/disable menu items
        flag = MF_BYCOMMAND | (MF_ENABLED if text_len > 1 else MF_GRAYED)
        for item_id in (IDM_FIND, IDM_FIND_NEXT, IDM_FIND_PREVIOUS, IDM_REPLACE):
            user32.EnableMenuItem(self.hmenu, item_id, flag)
        if text_len != self._saved_text_len:
            return True
        text_buf = create_unicode_buffer(text_len)
        user32.SendMessageW(self.edit.hwnd, WM_GETTEXT, text_len, text_buf)
        return text_buf.value != self._saved_text

    ########################################
    #
    ########################################
    def _check_if_text_selected(self):
        char_pos_start = DWORD()
        char_pos_end = DWORD()
        user32.SendMessageW(self.edit.hwnd, EM_GETSEL, byref(char_pos_start), byref(char_pos_end))
        # enable/disable menu items
        flag = MF_BYCOMMAND | (MF_ENABLED if char_pos_end.value > char_pos_start.value else MF_GRAYED)
        for item_id in (IDM_CUT, IDM_COPY, IDM_DELETE):
            user32.EnableMenuItem(self.hmenu, item_id, flag)

    ########################################
    #
    ########################################
    def action_encoding(self, encoding_id):
        if encoding_id == self._encoding_id:
            return
        user32.CheckMenuItem(self.hmenu, self._encoding_id, MF_BYCOMMAND | MF_UNCHECKED)
        self._encoding_id = encoding_id
        user32.CheckMenuItem(self.hmenu, self._encoding_id, MF_BYCOMMAND | MF_CHECKED)
        buf = create_unicode_buffer(24)
        user32.GetMenuStringW(self.hmenu, self._encoding_id, buf, 24, MF_BYCOMMAND)
        self.statusbar.set_text(buf.value, STATUSBAR_PART_ENCODING)

    ########################################
    #
    ########################################
    def action_eol_mode(self, eol_mode_id):
        if eol_mode_id == self._eol_mode_id:
            return
        user32.CheckMenuItem(self.hmenu, self._eol_mode_id, MF_BYCOMMAND | MF_UNCHECKED)
        self._eol_mode_id = eol_mode_id
        user32.CheckMenuItem(self.hmenu, self._eol_mode_id, MF_BYCOMMAND | MF_CHECKED)
        buf = create_unicode_buffer(24)
        user32.GetMenuStringW(self.hmenu, self._eol_mode_id, buf, 24, MF_BYCOMMAND)
        self.statusbar.set_text(buf.value, STATUSBAR_PART_EOL)

    ########################################
    #
    ########################################
    def action_exit(self, *args):
        if not self._handle_dirty():
            user32.SetFocus(self.edit.hwnd)
            return 1
        self._save_state()
        super().quit()

    ########################################
    #
    ########################################
    def action_new(self):
        if not self._handle_dirty():
            user32.SetFocus(self.edit.hwnd)
            return
        self._filename = None
        self._is_dirty = False
        self._saved_text_len = 1
        self._saved_text = ''

        self._show_caret_pos()

        self._eol_mode_id = IDM_EOL_CRLF
        self._encoding_id = IDM_UTF_8
        buf = create_unicode_buffer(24)

        user32.GetMenuStringW(self.hmenu, self._eol_mode_id, buf, 24, MF_BYCOMMAND)
        self.statusbar.set_text(buf.value, STATUSBAR_PART_EOL)

        user32.GetMenuStringW(self.hmenu, self._encoding_id, buf, 24, MF_BYCOMMAND)
        self.statusbar.set_text(buf.value, STATUSBAR_PART_ENCODING)

        user32.SendMessageW(self.edit.hwnd, WM_SETTEXT, 0, create_unicode_buffer(''))
        self.set_window_text(self._get_caption())
        user32.SetFocus(self.edit.hwnd)

    ########################################
    #
    ########################################
    def action_new_window(self):
        os.startfile(sys.executable, arguments=f'"{sys.argv[0]}"')

    ########################################
    #
    ########################################
    def action_open(self):
        if not self._handle_dirty():
            return
        filename = self.get_open_filename(_('Open'), '.txt',
                _('Text Documents') + ' (*.txt)\0*.txt\0' + _('All Files') + ' (*.*)\0*.*\0\0')
        if filename:
            self._load_file(filename)

    ########################################
    #
    ########################################
    def action_save(self):
        if not self._filename:
            return self.action_save_as()
        if not self._save_file(self._filename):
            return False
        self._is_dirty = False
        self.set_window_text(self._get_caption())
        return True

    ########################################
    #
    ########################################
    def action_save_as(self):
        filename = self.get_save_filename(_('Save As'),
                '.txt', _('Text Documents') + ' (*.txt)\0*.txt\0' + _('All Files') + ' (*.*)\0*.*\0\0',
                self._filename if self._filename else '')
        if not filename or not  self._save_file(filename):
            return False
        self._filename = filename
        self._is_dirty = False
        self.set_window_text(self._get_caption())
        return True

    ########################################
    #
    ########################################
    def action_print(self):
        tmp_dir = os.path.join(os.environ['TMP'], APP_NAME)
        if not os.path.isdir(tmp_dir):
            os.mkdir(tmp_dir)
        tmp_file = os.path.join(tmp_dir, self._filename if self._filename else _('Untitled'))
        text_len = user32.SendMessageW(self.edit.hwnd, WM_GETTEXTLENGTH, 0, 0) + 1
        text_buf = create_unicode_buffer(text_len)
        user32.SendMessageW(self.edit.hwnd, WM_GETTEXT, text_len, text_buf)
        with open(tmp_file, 'wb') as f:
            f.write(text_buf.value.encode('utf-8-sig'))
        shell32.ShellExecuteW(self.hwnd, 'open', 'rundll32.exe',
                f'mshtml.dll,PrintHTML "{tmp_file}"',
                None,
                SW_SHOW)

    ########################################
    #
    ########################################
    def action_undo(self):
        user32.SendMessageW(self.edit.hwnd, EM_UNDO, 0, 0)

    ########################################
    #
    ########################################
    def action_cut(self):
        user32.SendMessageW(self.edit.hwnd, WM_CUT, 0, 0)

    ########################################
    #
    ########################################
    def action_copy(self):
        user32.SendMessageW(self.edit.hwnd, WM_COPY, 0, 0)

    ########################################
    # we can't use WM_PASTE because it can't handle Unix/macOS EOL
    ########################################
    def action_paste(self):
        user32.OpenClipboard(0)
        try:
            if user32.IsClipboardFormatAvailable(CF_UNICODETEXT):
                data = user32.GetClipboardData(CF_UNICODETEXT)
                data_locked = kernel32.GlobalLock(data)
                text = c_wchar_p(data_locked)
                kernel32.GlobalUnlock(data_locked)
                txt = text.value
                if '\n' in txt and not '\r' in txt:
                    txt = txt.replace('\n', '\r\n')
                elif '\r' in txt and not '\n' in txt:
                    txt = txt.replace('\r', '\r\n')
                user32.SendMessageW(self.edit.hwnd, EM_REPLACESEL, 1, create_unicode_buffer(txt))
        finally:
            user32.CloseClipboard()

    ########################################
    #
    ########################################
    def action_delete(self):
        user32.SendMessageW(self.edit.hwnd, WM_CLEAR, 0, 0)

    ########################################
    #
    ########################################
    def action_find(self):
        if self.dialog_find.hwnd:
            user32.SetActiveWindow(self.dialog_find.hwnd)
            return
        elif self.dialog_replace.hwnd:
            user32.SendMessageW(self.dialog_replace.hwnd, WM_CLOSE, 0, 0)
        self.dialog_show_async(self.dialog_find)

    ########################################
    #
    ########################################
    def action_find_next(self):
        if self._search_term == '':
            return self.action_find()
        self._find(False)

    ########################################
    #
    ########################################
    def action_find_previous(self):
        if self._search_term == '':
            return self.action_find()
        self._find(True)

    ########################################
    #
    ########################################
    def action_replace(self):
        if self.dialog_replace.hwnd:
            user32.SetActiveWindow(self.dialog_replace.hwnd)
            return
        elif self.dialog_find.hwnd:
            user32.SendMessageW(self.dialog_find.hwnd, WM_CLOSE, 0, 0)
        self.dialog_show_async(self.dialog_replace)

    ########################################
    #
    ########################################
    def action_go_to(self):
        line_goto = self.dialog_show_sync(self.dialog_goto)
        if line_goto > 0:
            total_lines = user32.SendMessageW(self.edit.hwnd, EM_GETLINECOUNT, 0, 0)
            if line_goto > total_lines:
                self.show_message_box(_('GOTO_BEYOND'), APP_NAME + ' - ' + _('Goto Line'))
            else:
                if line_goto == 0:
                    line_goto = 1
                pos = user32.SendMessageW(self.edit.hwnd, EM_LINEINDEX, line_goto - 1, 0)
                user32.SendMessageW(self.edit.hwnd, EM_SETSEL, pos, pos)
        user32.SetFocus(self.edit.hwnd)

    ########################################
    #
    ########################################
    def action_select_all(self):
        user32.SendMessageW(self.edit.hwnd, EM_SETSEL, 0, -1)
        self._check_if_text_selected()

    ########################################
    #
    ########################################
    def action_time_date(self):
        date_str = time.strftime('%X %x')
        user32.SendMessageW(self.edit.hwnd, EM_REPLACESEL, 1, create_unicode_buffer(date_str))

    ########################################
    #
    ########################################
    def action_status_bar(self):
        self._show_statusbar = not self._show_statusbar
        user32.CheckMenuItem(self.hmenu, IDM_STATUS_BAR, MF_BYCOMMAND | (MF_CHECKED if self._show_statusbar else MF_UNCHECKED))
        self.statusbar.show(SW_SHOW if self._show_statusbar else SW_HIDE)
        rc = self.get_client_rect()
        width = rc.right - rc.left
        height = rc.bottom - rc.top
        if self._show_statusbar:
            user32.SetWindowPos(self.edit.hwnd, 0, 0, 0, width, height - self.statusbar.height, 0)
        else:
            user32.SetWindowPos(self.edit.hwnd, 0, 0, 0, width, height, 0)

    ########################################
    #
    ########################################
    def action_dark_mode(self):
        self._dark_mode = not self._dark_mode
        user32.CheckMenuItem(self.hmenu, IDM_DARK_MODE, MF_BYCOMMAND | (MF_CHECKED if self._dark_mode else MF_UNCHECKED))
        self.apply_theme(self._dark_mode)

    ########################################
    #
    ########################################
    def action_zoom_in(self):
        self._zoom += 10
        font = list(self._font)
        font[1] = int(self._font[1] * self._zoom / 100)
        self.edit.set_font(*font)
        self.statusbar.set_text(f'{self._zoom}%', STATUSBAR_PART_ZOOM)

    ########################################
    #
    ########################################
    def action_zoom_out(self):
        self._zoom -= 10
        font = list(self._font)
        font[1] = int(self._font[1] * self._zoom / 100)
        self.edit.set_font(*font)
        self.statusbar.set_text(f'{self._zoom}%', STATUSBAR_PART_ZOOM)

    ########################################
    #
    ########################################
    def action_restore_default_zoom(self):
        self._zoom = 100
        self.edit.set_font(*self._font)
        self.statusbar.set_text('100%', STATUSBAR_PART_ZOOM)

    ########################################
    # WS_HSCROLL can't be changed after edit control was created, so the control needs to be replaced.
    # (MS Notepad does the same thing).
    ########################################
    def action_word_wrap(self):
        self._word_wrap = not self._word_wrap

        user32.CheckMenuItem(self.hmenu, IDM_WORD_WRAP,
            MF_BYCOMMAND | (MF_CHECKED if self._word_wrap else MF_UNCHECKED))
        user32.EnableMenuItem(self.hmenu, IDM_GO_TO,
            MF_BYCOMMAND | (MF_ENABLED if not self._word_wrap else MF_GRAYED))

        text_len = user32.SendMessageW(self.edit.hwnd, WM_GETTEXTLENGTH, 0, 0) + 1
        text_buf = create_unicode_buffer(text_len)
        user32.SendMessageW(self.edit.hwnd, WM_GETTEXT, text_len, text_buf)
        pos_start, pos_end = DWORD(), DWORD()
        user32.SendMessageW(self.edit.hwnd, EM_GETSEL, byref(pos_start), byref(pos_end))

        style = user32.GetWindowLongA(self.edit.hwnd, GWL_STYLE)
        if self._word_wrap:
            style &= ~WS_HSCROLL
        else:
            style |= WS_HSCROLL

        self.edit.destroy_window()

        self._create_edit()
        if self._dark_mode:
            self.edit.apply_theme(self._dark_mode)

        user32.SendMessageW(self.edit.hwnd, WM_SETTEXT, 0, byref(text_buf))
        user32.SendMessageW(self.edit.hwnd, EM_SETSEL, pos_start.value, pos_end.value)

        rc = self.get_client_rect()
        width = rc.right - rc.left
        height = rc.bottom - rc.top
        user32.SetWindowPos(self.edit.hwnd, 0, 0, 0, width,
                height - self.statusbar.height if self._show_statusbar else height, 0)

        self._check_caret_pos()
        user32.SetFocus(self.edit.hwnd)

    ########################################
    #
    ########################################
    def action_font(self):
        res = self.show_font_dialog(*self._font)
        if res is not None:
            self.edit.set_font(*res)
            self._font = res

    ########################################
    #
    ########################################
    def action_about_notepad(self):
        self.show_message_box(
                _('ABOUT_TEXT').format(APP_NAME, APP_VERSION, APP_COPYRIGHT),
                _('ABOUT_CAPTION').format(APP_NAME))
        user32.SetFocus(self.edit.hwnd)

    ########################################
    #
    ########################################
    def action_about_windows(self):
        shell32.ShellAboutW(self.hwnd, create_unicode_buffer('Windows'), None, None)


if __name__ == "__main__":
    app = App(sys.argv[1:])
    sys.exit(app.run())
