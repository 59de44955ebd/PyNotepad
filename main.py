import contextlib
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
APP_VERSION = 2
APP_COPYRIGHT = '2024 github.com/59de44955ebd'
APP_DIR = os.path.dirname(os.path.abspath(__file__))

LANG = locale.windows_locale[kernel32.GetUserDefaultUILanguage()]
if not os.path.isdir(os.path.join(APP_DIR, 'resources', LANG)):
    LANG = 'en_US'

with open(os.path.join(APP_DIR, 'resources', LANG, 'strings.pson'), 'rb') as f:
    __ = eval(f.read())

def tr(s):
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

TAB_SIZES = {
    IDM_TAB_SIZE_2: 2,
    IDM_TAB_SIZE_4: 4,
    IDM_TAB_SIZE_8: 8,
}

EDIT_MAX_TEXT_LEN = 0x80000  # 512 KB (Edit control's default: 30.000)

STATUSBAR_PART_CARET = 1
STATUSBAR_PART_ZOOM = 2
STATUSBAR_PART_EOL = 3
STATUSBAR_PART_ENCODING = 4

INCHES_PER_UNIT = 1 / 2540

@contextlib.contextmanager
def no_redraw(edit):
    edit.send_message(WM_SETREDRAW, FALSE, 0)
    try:
        yield
    finally:
        edit.send_message(WM_SETREDRAW, TRUE, 0)


class App(MainWin):

    def __init__(self, args=[]):
        # defaults
        self._show_statusbar = True
        self._dark_mode = reg_should_use_dark_mode() if DARK_SUPPORTED else False
        self._word_wrap = True
        self._margin = 4
        self._font = ['Consolas', 11, 400, FALSE]
        self._search_term = ''
        self._saved_search_term = ''
        self._replace_term = ''
        self._match_case = FALSE
        self._wrap_arround = FALSE
        self._search_up = FALSE
        self._tab_size = 4
        self._use_spaces = False
        self._zoom = 100
        self._filename = None
        self._is_dirty = False
        self._saved_text_len = 1
        self._saved_text = ''
        self._eol_mode_id = IDM_EOL_CRLF
        self._encoding_id = IDM_UTF_8
        self._print_paper_size = [21000, 29700]  # in mm/100
        self._print_margins = [1000, 1500, 1000, 1500]  # in mm/100

        self._last_sel = [0, 0]
        self._current_sel = [0, 0]

        left, top, width, height = self._load_state()

        # load menu resource
        with open(os.path.join(APP_DIR, 'resources', LANG, 'menu.pson'), 'rb') as f:
            menu_data = eval(f.read())

        self.COMMAND_MESSAGE_MAP = {
            IDM_NEW:                self.action_new,
            IDM_NEW_WINDOW:         self.action_new_window,
            IDM_OPEN:               self.action_open,
            IDM_SAVE:               self.action_save,
            IDM_SAVE_AS:            self.action_save_as,
            IDM_PAGE_SETUP:         self.action_page_setup,
            IDM_PRINT:              self.action_print,
            IDM_EXIT:               self.quit,
            IDM_UNDO:               self.action_undo,
            IDM_CUT:                self.action_cut,
            IDM_COPY:               self.action_copy,
            IDM_PASTE:              self.action_paste,
            IDM_FIND:               self.action_find,
            IDM_FIND_NEXT:          self.action_find_next,
            IDM_FIND_PREVIOUS:      self.action_find_previous,
            IDM_REPLACE:            self.action_replace,
            IDM_GO_TO:              self.action_go_to,
            IDM_SELECT_ALL:         self.action_select_all,
            IDM_TIME_DATE:          self.action_insert_time_date,
            IDM_WORD_WRAP:          self.action_word_wrap,
            IDM_FONT:               self.action_set_font,
            IDM_TABS_AS_SPACES:     self.action_toggle_use_spaces,
            IDM_ZOOM_IN:            self.action_zoom_in,
            IDM_ZOOM_OUT:           self.action_zoom_out,
            IDM_ZOOM_RESET:         self.action_zoom_reset,
            IDM_STATUS_BAR:         self.action_toggle_statusbar,
            IDM_DARK_MODE:          self.action_toggle_dark_mode,
            IDM_ABOUT_NOTEPAD:      self.action_about_notepad,
            IDM_ABOUT_WINDOWS:      self.action_about_windows,
        }

        for item_id in ENCODINGS.keys():
            self.COMMAND_MESSAGE_MAP[item_id] = lambda id=item_id: self.action_set_encoding(id)

        for item_id in EOL_MODES.keys():
            self.COMMAND_MESSAGE_MAP[item_id] = lambda id=item_id: self.action_set_eol_mode(id)

        for item_id in TAB_SIZES.keys():
            self.COMMAND_MESSAGE_MAP[item_id] = lambda id=item_id: self.action_set_tab_size(id)

        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            hicon = user32.LoadIconW(kernel32.GetModuleHandleW(None), MAKEINTRESOURCEW(1))
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
        if self._use_spaces:
            user32.CheckMenuItem(self.hmenu, IDM_TABS_AS_SPACES, MF_BYCOMMAND | MF_CHECKED)
        tab_size_id= list(TAB_SIZES.keys())[list(TAB_SIZES.values()).index(self._tab_size)]
        user32.CheckMenuItem(self.hmenu, tab_size_id, MF_BYCOMMAND | MF_CHECKED)

        self._create_statusbar()
        self._create_dialogs()
        self._create_edit()

        ########################################
        #
        ########################################
        def _on_WM_SIZE(hwnd, wparam, lparam):
            width, height = lparam & 0xFFFF, (lparam >> 16) & 0xFFFF
            self.statusbar.update_size(width)
            user32.SetWindowPos(self.edit.hwnd, 0, 0, 0, width,
                    height - self.statusbar.height if self._show_statusbar else height, 0)
        self.register_message_callback(WM_SIZE, _on_WM_SIZE)

        ########################################
        #
        ########################################
        def _on_WM_COMMAND(hwnd, wparam, lparam):
            command = HIWORD(wparam)
            if lparam == 0:
                command_id = LOWORD(wparam)
                if command_id in self.COMMAND_MESSAGE_MAP:
                    self.COMMAND_MESSAGE_MAP[command_id]()
            elif lparam == self.edit.hwnd and command == EN_CHANGE:
                is_dirty = self._check_if_text_changed()
                if is_dirty != self._is_dirty:
                    self._is_dirty = is_dirty
                    self.set_window_text(self._get_caption())
            return FALSE
        self.register_message_callback(WM_COMMAND, _on_WM_COMMAND)

        ########################################
        #
        ########################################
        def _on_WM_DROPFILES(hwnd, wparam, lparam):
            dropped_items = self.get_dropped_items(wparam)
            if os.path.isfile(dropped_items[0]) and self._handle_dirty():
                self._load_file(dropped_items[0])
        self.register_message_callback(WM_DROPFILES, _on_WM_DROPFILES)

        if DARK_SUPPORTED:
            def _on_WM_SETTINGCHANGE(hwnd, wparam, lparam):
                if lparam and cast(lparam, LPCWSTR).value == 'ImmersiveColorSet':
                    if reg_should_use_dark_mode() != self._dark_mode:
                        self.action_toggle_dark_mode()
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
    def quit(self, *args):
        if not self._handle_dirty():
            user32.SetFocus(self.edit.hwnd)
            return 1
        self._save_state()
        super().quit()

    ########################################
    #
    ########################################
    def _create_dialogs(self):
        with open(os.path.join(APP_DIR, 'resources', LANG, 'dialog_find.pson'), 'rb') as f:
            dialog_dict = eval(f.read())

        def _dialog_proc_find(hwnd, msg, wparam, lparam):

            if msg == WM_INITDIALOG:
                hwnd_edit = user32.GetDlgItem(hwnd, ID_EDIT_FIND)

                # limit search input to 127 chars
                user32.SendMessageW(hwnd_edit, EM_SETLIMITTEXT, 127, 0)

                # check if something is selected
                pos_start, pos_end = DWORD(), DWORD()
                self.edit.send_message(EM_GETSEL, byref(pos_start), byref(pos_end))
                if pos_end.value > pos_start.value:
                    text_len = self.edit.send_message(WM_GETTEXTLENGTH, 0, 0) + 1
                    text_buf = create_unicode_buffer(text_len)
                    self.edit.send_message(WM_GETTEXT, text_len, text_buf)
                    self._search_term = text_buf.value[pos_start.value:pos_end.value][:127]
                elif self._search_term == '':
                    self._search_term = self._saved_search_term

                # update button states
                if self._search_term:
                    user32.SendMessageW(hwnd_edit, WM_SETTEXT, 0, create_unicode_buffer(self._search_term))
                else:
                    user32.EnableWindow(user32.GetDlgItem(hwnd, ID_OK), 0)
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
                        user32.EnableWindow(user32.GetDlgItem(hwnd, ID_OK), int(text_len > 0))

                elif command == BN_CLICKED:
                    if control_id == ID_OK:
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

        with open(os.path.join(APP_DIR, 'resources', LANG, 'dialog_replace.pson'), 'rb') as f:
            dialog_dict = eval(f.read())

        def _dialog_proc_replace(hwnd, msg, wparam, lparam):
            if msg == WM_INITDIALOG:
                # limit search and replace input to 127 chars
                hwnd_search_edit = user32.GetDlgItem(hwnd, ID_EDIT_FIND)
                user32.SendMessageW(hwnd_search_edit, EM_SETLIMITTEXT, 127, 0)

                hwnd_replace_edit = user32.GetDlgItem(hwnd, ID_EDIT_REPLACE)
                user32.SendMessageW(hwnd_replace_edit, EM_SETLIMITTEXT, 127, 0)

                # check if something is selected
                pos_start, pos_end = DWORD(), DWORD()
                self.edit.send_message(EM_GETSEL, byref(pos_start), byref(pos_end))
                if pos_end.value > pos_start.value:
                    text_len = self.edit.send_message(WM_GETTEXTLENGTH, 0, 0) + 1
                    text_buf = create_unicode_buffer(text_len)
                    self.edit.send_message(WM_GETTEXT, text_len, text_buf)
                    self._search_term = text_buf.value[pos_start.value:pos_end.value][:127]
                # update button states
                if self._search_term:
                    user32.SendMessageW(hwnd_search_edit, WM_SETTEXT, 0, create_unicode_buffer(self._search_term))
                else:
                    user32.EnableWindow(user32.GetDlgItem(hwnd, ID_OK), 0)
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
                        user32.EnableWindow(user32.GetDlgItem(hwnd, ID_OK), int(text_len > 0))
                        user32.EnableWindow(user32.GetDlgItem(hwnd, ID_REPLACE), int(text_len > 0))
                        user32.EnableWindow(user32.GetDlgItem(hwnd, ID_REPLACE_ALL), int(text_len > 0))

                elif command == BN_CLICKED:
                    if control_id in (ID_OK, ID_REPLACE, ID_REPLACE_ALL):
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

                        if control_id == ID_OK:
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

        with open(os.path.join(APP_DIR, 'resources', LANG, 'dialog_goto.pson'), 'rb') as f:
            dialog_dict = eval(f.read())

        def _dialog_proc_goto(hwnd, msg, wparam, lparam):
            if msg == WM_INITDIALOG:
                hwnd_edit = user32.GetDlgItem(hwnd, ID_EDIT_GOTO)
                line_idx = self.edit.send_message(EM_LINEFROMCHAR, -1, 0)
                user32.SendMessageW(hwnd_edit, WM_SETTEXT, 0,
                        create_unicode_buffer(str(line_idx + 1)))
                user32.SendMessageW(hwnd_edit, EM_SETSEL, 0, -1)

            elif msg == WM_COMMAND:
                control_id = LOWORD(wparam)
                command = HIWORD(wparam)
                if command == BN_CLICKED:
                    if control_id == ID_OK:
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
            ES_AUTOVSCROLL | ES_NOHIDESEL | (0 if self._word_wrap else WS_HSCROLL),
            bg_color_dark=0x000000,  # 0x000000 0x171717 default: 0x383838
#            text_color_dark=0xffffff,  # default: 0xe0e0e0
        )

        self.edit.send_message(EM_SETLIMITTEXT, EDIT_MAX_TEXT_LEN, 0)
        self.edit.send_message(EM_SETMARGINS, EC_LEFTMARGIN | EC_RIGHTMARGIN, MAKELONG(self._margin, self._margin))
        self._set_editor_font(self._font)

        ########################################
        #
        ########################################
        def _on_WM_KEYUP(hwnd, wparam, lparam):
            self._check_caret_pos()
            self._check_if_text_selected()
        self.edit.register_message_callback(WM_KEYUP, _on_WM_KEYUP)

        ########################################
        #
        ########################################
        def _on_WM_CHAR(hwnd, wparam, lparam):
            if wparam == VK_BACK and self._use_spaces:
                return self._handle_back()
            elif wparam == VK_TAB:
                self._handle_tab()
                return 1
        self.edit.register_message_callback(WM_CHAR, _on_WM_CHAR)

        ########################################
        #
        ########################################
        def _on_WM_MOUSEMOVE(hwnd, wparam, lparam):
            if wparam & MK_LBUTTON:
                # Edit control only knows about selections, not about the active caret position.
                # So we have to find out the active (=changing) side of the selection ourself.
                self._last_sel = self._current_sel
                res = self.edit.send_message(EM_GETSEL, 0, 0)
                pos_from, pos_to = res & 0xFFFF, (res >> 16) & 0xFFFF
                self._current_sel = [pos_from, pos_to]
                if self._current_sel == self._last_sel:
                    return
                pos = pos_to if self._current_sel[1] != self._last_sel[1] else pos_from
                line_idx = self.edit.send_message(EM_LINEFROMCHAR, pos, 0)
                res = self.edit.send_message(EM_POSFROMCHAR, pos, 0)
                if res == -1:
                    # special case, position is EOF, so go one char back
                    line_idx_new = self.edit.send_message(EM_LINEFROMCHAR, pos - 1, 0)
                    if line_idx_new < line_idx:
                        col_idx = 0
                    else:
                        res = self.edit.send_message(EM_POSFROMCHAR, pos - 1, 0)
                        col_idx = (res & 0xFFFF - self._margin) // self._char_width + 1
                else:
                    col_idx = (res & 0xFFFF - self._margin) // self._char_width
                self._show_caret_pos(line_idx, col_idx)

        self.edit.register_message_callback(WM_MOUSEMOVE, _on_WM_MOUSEMOVE)

        ########################################
        #
        ########################################
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

            if advapi32.RegQueryValueExW(hkey, 'DarkMode', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                self._dark_mode = cast(data, POINTER(DWORD)).contents.value == 1
            if advapi32.RegQueryValueExW(hkey, 'iPointSize', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                self._font[1] = cast(data, POINTER(DWORD)).contents.value
            if advapi32.RegQueryValueExW(hkey, 'lfWeight', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                self._font[2] = cast(data, POINTER(DWORD)).contents.value
            if advapi32.RegQueryValueExW(hkey, 'lfItalic', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                self._font[3] = cast(data, POINTER(DWORD)).contents.value  #== 1

            if advapi32.RegQueryValueExW(hkey, 'fWrap', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                self._word_wrap = cast(data, POINTER(DWORD)).contents.value == 1
            if advapi32.RegQueryValueExW(hkey, 'fUseSpaces', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                self._use_spaces = cast(data, POINTER(DWORD)).contents.value == 1
            if advapi32.RegQueryValueExW(hkey, 'iTabSize', None, None, byref(data), byref(cbData)) == ERROR_SUCCESS:
                self._tab_size = cast(data, POINTER(DWORD)).contents.value

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
            advapi32.RegSetValueExW(hkey, 'fUseSpaces', 0, REG_DWORD, byref(DWORD(int(self._use_spaces))), dwsize)
            advapi32.RegSetValueExW(hkey, 'iTabSize', 0, REG_DWORD, byref(DWORD(self._tab_size)), dwsize)

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
    # indents/unindents selected block
    ########################################
    def _handle_tab(self):
        tab = ' ' * self._tab_size if self._use_spaces else '\t'
        tab_len = len(tab)
        is_shift = user32.GetAsyncKeyState(VK_SHIFT) > 1
        res = self.edit.send_message(EM_GETSEL, 0, 0)
        pos_from, pos_to = res & 0xFFFF, (res >> 16) & 0xFFFF
        line_from = self.edit.send_message(EM_LINEFROMCHAR, pos_from, 0)
        if pos_to > pos_from:
            line_to = self.edit.send_message(EM_LINEFROMCHAR, pos_to, 0)
            is_multline = line_to > line_from
        else:
            is_multline = False

        if is_multline:
            # get full block
            sel_pos_from = self.edit.send_message(EM_LINEINDEX, line_from, 0)

            line_start_pos = self.edit.send_message(EM_LINEINDEX, line_to, 0)
            sel_pos_to = line_start_pos + self.edit.send_message(EM_LINELENGTH, line_start_pos, 0)

            if is_shift:
                # unindent block
                stripped = 0
                with no_redraw(self.edit):
                    for line_index in range(line_from, line_to + 1):
                        # check the line starts with tab
                        buf = (WCHAR * (tab_len + 1))()
                        buf[0] = chr(tab_len)
                        self.edit.send_message(EM_GETLINE, line_index, byref(buf))
                        if buf.value == tab:  #.startswith(tab):
                            line_start_pos = self.edit.send_message(EM_LINEINDEX, line_index, 0)
                            self.edit.send_message(EM_SETSEL, line_start_pos, line_start_pos + tab_len)
                            self.edit.send_message(WM_CLEAR, 0, 0)
                            stripped += tab_len
                    # update selection to new size
                    self.edit.send_message(EM_SETSEL, sel_pos_from, sel_pos_to - stripped)
            else:
                # indent block
                tab_buf = create_unicode_buffer(tab)
                with no_redraw(self.edit):
                    for line_index in range(line_from, line_to + 1):
                        line_start_pos = self.edit.send_message(EM_LINEINDEX, line_index, 0)
                        self.edit.send_message(EM_SETSEL, line_start_pos, line_start_pos)
                        self.edit.send_message(EM_REPLACESEL, TRUE, tab_buf)  # TRUE for undoable
                    # update selection to new size
                    self.edit.send_message(EM_SETSEL, sel_pos_from, sel_pos_to + tab_len * (line_to - line_from + 1))

        elif is_shift:
            # jump back to preceding tab pos
            res = self.edit.send_message(EM_POSFROMCHAR, pos_from, 0)
            if res == -1:
                # special case, position is EOF
                line_from_new = self.edit.send_message(EM_LINEFROMCHAR, pos_from - 1, 0)
                if line_from_new < line_from:
                    return
                res = self.edit.send_message(EM_POSFROMCHAR, pos_from - 1, 0)
                x, y = res & 0xFFFF + self._char_width, (res >> 16) & 0xFFFF
            else:
                x, y = res & 0xFFFF, (res >> 16) & 0xFFFF
            pos_char_px = (x - self._margin) // self._char_width
            if pos_char_px == 0:
                return
            pos_char_px_new = (pos_char_px - 1) // self._tab_size * self._tab_size
            pos_new = LOWORD(self.edit.send_message(EM_CHARFROMPOS, 0, MAKELONG(self._margin + pos_char_px_new * self._char_width, y)))
            self.edit.send_message(EM_SETSEL, pos_new, pos_new)

        else:
            self.edit.send_message(EM_REPLACESEL, 1, create_unicode_buffer(tab))

    ########################################
    # if _use_spaces is True and there are only spaces before caret,
    # remove according to self._tab_size
    ########################################
    def _handle_back(self):
        res = self.edit.send_message(EM_GETSEL, 0, 0)
        pos_from, pos_to = res & 0xFFFF, (res >> 16) & 0xFFFF
        if pos_from == pos_to:
            line_index = self.edit.send_message(EM_LINEFROMCHAR, pos_from, 0)
            line_start_pos = self.edit.send_message(EM_LINEINDEX, line_index, 0)
            if pos_from - line_start_pos >= self._tab_size:
                buf_len = pos_from - line_start_pos
                buf = (WCHAR * (buf_len + 1))()
                buf[0] = chr(buf_len)
                self.edit.send_message(EM_GETLINE, line_index, byref(buf))
                if buf.value == ' ' * len(buf.value):
                    pos_new = line_start_pos + (len(buf.value) - 1) // self._tab_size * self._tab_size
                    with no_redraw(self.edit):
                        self.edit.send_message(EM_SETSEL, pos_new, pos_from)
                        self.edit.send_message(WM_CLEAR, 0, 0)
                    return 1

    ########################################
    #
    ########################################
    def _set_editor_font(self, font):
        self.edit.set_font(*font)
        self.edit.send_message(EM_SETTABSTOPS, 1, byref(UINT(self._tab_size * 4)))
        self._calculate_char_width()

    ########################################
    #
    ########################################
    def _calculate_char_width(self):
        hdc = user32.GetDC(0)
        gdi32.SelectObject(hdc, self.edit.hfont)
        tm = TEXTMETRICW()
        gdi32.GetTextMetricsW(hdc, byref(tm))
        user32.ReleaseDC(0, hdc)
        self._char_width = tm.tmAveCharWidth

    ########################################
    #
    ########################################
    def _find(self, search_up=None):
        sel_start_pos = DWORD()
        sel_end_pos = DWORD()
        self.edit.send_message(EM_GETSEL, byref(sel_start_pos), byref(sel_end_pos))

        text_len = self.edit.send_message(WM_GETTEXTLENGTH, 0, 0) + 1
        text_buf = create_unicode_buffer(text_len)
        self.edit.send_message(WM_GETTEXT, text_len, text_buf)
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
                self.statusbar.set_text(tr('Found next from the bottom') if pos > -1 else '')
            else:
                pos = haystack.find(needle, 0, sel_end_pos.value)
                self.statusbar.set_text(tr('Found next from the top') if pos > -1 else '')
        else:
            self.statusbar.set_text()

        if pos > -1:
            self.edit.send_message(EM_SETSEL, pos, pos + len(self._search_term))
            self._check_caret_pos()
            return True
        else:
             self.show_message_box(tr('CANNOT_FIND').format(self._search_term), APP_NAME)
             return False

    ########################################
    #
    ########################################
    def _replace(self):
        if self._find():
            self.edit.send_message(EM_REPLACESEL, 1, create_unicode_buffer(self._replace_term))

    ########################################
    #
    ########################################
    def _replace_all(self):
        text_len = self.edit.send_message(WM_GETTEXTLENGTH, 0, 0) + 1
        text_buf = create_unicode_buffer(text_len)
        self.edit.send_message(WM_GETTEXT, text_len, text_buf)
        txt = text_buf.value
        if self._match_case:
            txt_new = txt.replace(self._search_term, self._replace_term)
        else:
            txt_new = re.compile(re.escape(self._search_term), re.IGNORECASE).sub(self._replace_term, txt)
        if txt_new != txt:
            self.edit.send_message(EM_SETSEL, 0, -1)
            self.edit.send_message(EM_REPLACESEL, 1, create_unicode_buffer(txt_new))
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
    		    return False  # invalid UTF-8 sequence
    		for c in range(n):
    			i += 1
    			if i > data_len:
    			    return False  # invalid UTF-8 sequence
    			if data[i] & 192 != 128:
    			    return False  # invalid UTF-8 sequence
    	return True

    ########################################
    #
    ########################################
    def _check_caret_pos(self):
        pt = POINT()
        ok = user32.GetCaretPos(byref(pt))
        pos = self.edit.send_message(EM_CHARFROMPOS, 0, MAKELONG(pt.x, pt.y))
        if pos > -1:
            line_idx = HIWORD(pos)
            col_idx = (pt.x - self._margin) // self._char_width
            self._show_caret_pos(line_idx, col_idx)

    ########################################
    #
    ########################################
    def _show_caret_pos(self, line_idx=0, col_idx=0):
        self.statusbar.set_text(tr('LINE_COLUMN').format(line_idx + 1, col_idx + 1), STATUSBAR_PART_CARET)

    ########################################
    #
    ########################################
    def _load_file(self, filename):
        if os.stat(filename).st_size > EDIT_MAX_TEXT_LEN:
            return self.show_message_box(tr('FILE_TOO_BIG'), APP_NAME, MB_ICONWARNING | MB_OK)
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
        self.edit.send_message(WM_SETTEXT, 0, text_buf)
        self._saved_text_len = self.edit.send_message(WM_GETTEXTLENGTH, 0, 0) + 1
        text_buf = create_unicode_buffer(self._saved_text_len)
        self.edit.send_message(WM_GETTEXT, self._saved_text_len, text_buf)
        flag = MF_BYCOMMAND | (MF_ENABLED if self._saved_text_len > 1 else MF_GRAYED)
        for item_id in (IDM_FIND, IDM_FIND_NEXT, IDM_FIND_PREVIOUS, IDM_REPLACE):
            user32.EnableMenuItem(self.hmenu, item_id, flag)
        self._saved_text = text_buf.value
        self._is_dirty = False
        self._filename = filename  # os.path.basename(filename)
        self.set_window_text(self._get_caption())
        user32.SetFocus(self.edit.hwnd)

    ########################################
    #
    ########################################
    def _save_file(self, filename):
        text_len = self.edit.send_message(WM_GETTEXTLENGTH, 0, 0) + 1
        text_buf = create_unicode_buffer(text_len)
        self.edit.send_message(WM_GETTEXT, text_len, text_buf)
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
        res = self.show_message_box_modern(
            tr('SAVE_CHANGES').format(self._filename if self._filename else tr('Untitled')),
            APP_NAME,
            MB_YESNOCANCEL
        )
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
        return ('*' if self._is_dirty else '') + (self._filename if self._filename else tr('Untitled')) + ' - ' + APP_NAME

    ########################################
    #
    ########################################
    def _check_if_text_changed(self):
        text_len = self.edit.send_message(WM_GETTEXTLENGTH, 0, 0) + 1
        # enable/disable menu items
        flag = MF_BYCOMMAND | (MF_ENABLED if text_len > 1 else MF_GRAYED)
        for item_id in (IDM_FIND, IDM_FIND_NEXT, IDM_FIND_PREVIOUS, IDM_REPLACE):
            user32.EnableMenuItem(self.hmenu, item_id, flag)
        if text_len != self._saved_text_len:
            return True
        text_buf = create_unicode_buffer(text_len)
        self.edit.send_message(WM_GETTEXT, text_len, text_buf)
        return text_buf.value != self._saved_text

    ########################################
    #
    ########################################
    def _check_if_text_selected(self):
        char_pos_start = DWORD()
        char_pos_end = DWORD()
        self.edit.send_message(EM_GETSEL, byref(char_pos_start), byref(char_pos_end))
        # enable/disable menu items
        flag = MF_BYCOMMAND | (MF_ENABLED if char_pos_end.value > char_pos_start.value else MF_GRAYED)
        for item_id in (IDM_CUT, IDM_COPY):
            user32.EnableMenuItem(self.hmenu, item_id, flag)

    ########################################
    #
    ########################################
    def action_set_encoding(self, encoding_id):
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
    def action_set_eol_mode(self, eol_mode_id):
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
    def action_set_tab_size(self, tab_size_id):
        tab_size_id_old = list(TAB_SIZES.keys())[list(TAB_SIZES.values()).index(self._tab_size)]
        user32.CheckMenuItem(self.hmenu, tab_size_id_old, MF_BYCOMMAND | MF_UNCHECKED)
        user32.CheckMenuItem(self.hmenu, tab_size_id, MF_BYCOMMAND | MF_CHECKED)
        self._tab_size = list(TAB_SIZES.values())[list(TAB_SIZES.keys()).index(tab_size_id)]
        self.edit.send_message(EM_SETTABSTOPS, 1, byref(UINT(self._tab_size * 4)))
        self._calculate_char_width()

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

        self.edit.send_message(WM_SETTEXT, 0, create_unicode_buffer(''))
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
        filename = self.get_open_filename(
            tr('Open'),
            '.txt',
            tr('Text Documents') + ' (*.txt)\0*.txt\0' + tr('All Files') + ' (*.*)\0*.*\0\0'
        )
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
        filename = self.get_save_filename(
            tr('Save As'),
            '.txt',
            tr('Text Documents') + ' (*.txt)\0*.txt\0' + tr('All Files') + ' (*.*)\0*.*\0\0',
            self._filename if self._filename else ''
        )
        if not filename or not self._save_file(filename):
            return False
        self._filename = filename
        self._is_dirty = False
        self.set_window_text(self._get_caption())
        return True

    ########################################
    #
    ########################################
    def action_page_setup(self):
        psd = PAGESETUPDLGW()
        psd.hwndOwner = self.hwnd
        psd.Flags = PSD_INHUNDREDTHSOFMILLIMETERS | PSD_MARGINS
        psd.ptPaperSize = POINT(*self._print_paper_size)
        psd.rtMargin = RECT(*self._print_margins)
        ok = comdlg32.PageSetupDlgW(byref(psd))
        if ok:
            self._print_paper_size = [psd.ptPaperSize.x, psd.ptPaperSize.y]
            self._print_margins = [psd.rtMargin.left, psd.rtMargin.top, psd.rtMargin.right, psd.rtMargin.bottom]

    ########################################
    #
    ########################################
    def action_print(self):
        print_paper_size = POINT(*self._print_paper_size)
        print_margins = RECT(*self._print_margins)

        pdlg = PRINTDLGW()
        # hwndOwner = 0 : legacy dialog
        # hwndOwner = desktop hwnd: slightly more modern dialog (non-UWP)
        # hwndOwner = self.hwnd: modern UWP dialog, but dialog only shown for the first time
        pdlg.hwndOwner = user32.GetDesktopWindow()
        pdlg.Flags = PD_RETURNDC | PD_USEDEVMODECOPIES | PD_NOSELECTION  # | PD_PRINTSETUP
        pdlg.nFromPage = 1
        pdlg.nToPage = 1
        pdlg.nMinPage = 1
        pdlg.nMaxPage = 0xffff
        pdlg.nStartPage = 0XFFFFFFFF  # START_PAGE_GENERAL
        if not comdlg32.PrintDlgW(byref(pdlg)):
            return False

        hdc = pdlg.hDC

        di = DOCINFOW()
        di.lpszDocName = os.path.basename(self._filename) if self._filename else tr('Untitled')
        job_id = gdi32.StartDocW(hdc, byref(di))
        if job_id <= 0:
            return False

        # Get printer resolution
        pt_dpi = POINT()
        pt_dpi.x = gdi32.GetDeviceCaps(hdc, LOGPIXELSX)
        pt_dpi.y = gdi32.GetDeviceCaps(hdc, LOGPIXELSY)

        font_name, font_size, font_weight, font_italic = self._font
        cHeight = -kernel32.MulDiv(font_size, pt_dpi.y, 72)
        hfont = gdi32.CreateFontW(cHeight, 0, 0, 0, font_weight, font_italic, FALSE, FALSE, ANSI_CHARSET, OUT_TT_PRECIS,
                CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, font_name)
        gdi32.SelectObject(hdc, hfont)

        # Set page rect in logical units, rcPage is the rectangle {0, 0, maxX, maxY} where
        # maxX+1 and maxY+1 are the number of physically printable pixels in x and y.
        # rc is the rectangle to render the text in (which will, of course, fit within the
        # rectangle defined by rcPage).
        rcPage = RECT()
        rcPage.top = 0
        rcPage.left = 0
        rcPage.right = round(print_paper_size.x * pt_dpi.x * INCHES_PER_UNIT)
        rcPage.bottom = round(print_paper_size.y * pt_dpi.x * INCHES_PER_UNIT)

        rc = RECT()
        rc.left = round(print_margins.left * pt_dpi.x * INCHES_PER_UNIT)
        rc_top = rc.top = round(print_margins.top * pt_dpi.x * INCHES_PER_UNIT)
        rc.right = rcPage.right - round(print_margins.right * pt_dpi.x * INCHES_PER_UNIT)
        rc_bottom = rc.bottom = rcPage.bottom - round(print_margins.bottom * pt_dpi.x * INCHES_PER_UNIT)

        text_len = self.edit.send_message(WM_GETTEXTLENGTH, 0, 0) + 1
        text_buf = create_unicode_buffer(text_len)
        self.edit.send_message(WM_GETTEXT, text_len, text_buf)
        lines = text_buf.value.split('\r\n')

        if pdlg.Flags & PD_PAGENUMS:
            first_page, last_page = pdlg.nFromPage, pdlg.nToPage
        else:
            first_page, last_page = 1, 0xffffffff

        # print with word wrap
        success = gdi32.StartPage(hdc)
        if success:
            page = 1
            for line in lines:
                h = user32.DrawTextW(hdc, line if line else ' ', -1, byref(RECT(rc.left, 0, rc.right, 0)), DT_CALCRECT | DT_WORDBREAK)

                if rc.top + h > rc_bottom:
                    page += 1
                    if page > last_page:
                        break
                    elif page > first_page:
                        gdi32.EndPage(hdc)
                        gdi32.StartPage(hdc)
                    rc.top = rc_top

                if page < first_page:
                    rc.top += h
                    continue

                h = user32.DrawTextW(
                    hdc,
                    line if line else '\r',
                    -1,
                    byref(rc),
                    DT_WORDBREAK if line else DT_SINGLELINE
                )
                rc.top += h

            gdi32.EndPage(hdc)

        if success:
            gdi32.EndDoc(hdc)
        else:
            gdi32.AbortDoc(hdc)

        gdi32.DeleteDC(hdc)

    ########################################
    #
    ########################################
    def action_undo(self):
        self.edit.send_message(EM_UNDO, 0, 0)

    ########################################
    #
    ########################################
    def action_cut(self):
        self.edit.send_message(WM_CUT, 0, 0)

    ########################################
    #
    ########################################
    def action_copy(self):
        self.edit.send_message(WM_COPY, 0, 0)

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
                self.edit.send_message(EM_REPLACESEL, 1, create_unicode_buffer(txt))
        finally:
            user32.CloseClipboard()

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
            total_lines = self.edit.send_message(EM_GETLINECOUNT, 0, 0)
            if line_goto > total_lines:
                self.show_message_box(tr('GOTO_BEYOND'), APP_NAME + ' - ' + tr('Goto Line'))
            else:
                if line_goto == 0:
                    line_goto = 1
                pos = self.edit.send_message(EM_LINEINDEX, line_goto - 1, 0)
                self.edit.send_message(EM_SETSEL, pos, pos)
        user32.SetFocus(self.edit.hwnd)

    ########################################
    #
    ########################################
    def action_select_all(self):
        self.edit.send_message(EM_SETSEL, 0, -1)
        self._check_if_text_selected()

    ########################################
    #
    ########################################
    def action_insert_time_date(self):
        date_str = time.strftime('%X %x')
        self.edit.send_message(EM_REPLACESEL, 1, create_unicode_buffer(date_str))

    ########################################
    #
    ########################################
    def action_toggle_statusbar(self):
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
    def action_toggle_dark_mode(self):
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
        self._set_editor_font(font)
        self.statusbar.set_text(f'{self._zoom}%', STATUSBAR_PART_ZOOM)

    ########################################
    #
    ########################################
    def action_zoom_out(self):
        if self._zoom <= 10:
            return
        self._zoom -= 10
        font = list(self._font)
        font[1] = int(self._font[1] * self._zoom / 100)
        self._set_editor_font(font)
        self.statusbar.set_text(f'{self._zoom}%', STATUSBAR_PART_ZOOM)

    ########################################
    #
    ########################################
    def action_zoom_reset(self):
        self._zoom = 100
        self._set_editor_font(self._font)
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

        text_len = self.edit.send_message(WM_GETTEXTLENGTH, 0, 0) + 1
        text_buf = create_unicode_buffer(text_len)
        self.edit.send_message(WM_GETTEXT, text_len, text_buf)
        pos_start, pos_end = DWORD(), DWORD()
        self.edit.send_message(EM_GETSEL, byref(pos_start), byref(pos_end))

        style = user32.GetWindowLongA(self.edit.hwnd, GWL_STYLE)
        if self._word_wrap:
            style &= ~WS_HSCROLL
        else:
            style |= WS_HSCROLL

        self.edit.destroy_window()

        self._create_edit()
        if self._dark_mode:
            self.edit.apply_theme(self._dark_mode)

        self.edit.send_message(WM_SETTEXT, 0, byref(text_buf))
        self.edit.send_message(EM_SETSEL, pos_start.value, pos_end.value)

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
    def action_set_font(self):
        res = self.show_font_dialog(*self._font)
        if res is None:
            return
        self._font = res
        if self._zoom != 100:
            font = list(self._font)
            font[1] = int(self._font[1] * self._zoom / 100)
            self._set_editor_font(font)
        else:
            self._set_editor_font(self._font)

    ########################################
    #
    ########################################
    def action_zoom_out(self):
        if self._zoom <= 10:
            return
        self._zoom -= 10
        font = list(self._font)
        font[1] = int(self._font[1] * self._zoom / 100)
        self._set_editor_font(font)
        self.statusbar.set_text(f'{self._zoom}%', STATUSBAR_PART_ZOOM)

    ########################################
    #
    ########################################
    def action_zoom_reset(self):
        self._zoom = 100

    ########################################
    #
    ########################################
    def action_toggle_use_spaces(self):
        self._use_spaces = not self._use_spaces
        user32.CheckMenuItem(self.hmenu, IDM_TABS_AS_SPACES,
            MF_BYCOMMAND | (MF_CHECKED if self._use_spaces else MF_UNCHECKED))

        # fix indentation accordingly
        text_len = self.edit.send_message(WM_GETTEXTLENGTH, 0, 0) + 1
        text_buf = create_unicode_buffer(text_len)
        self.edit.send_message(WM_GETTEXT, text_len, text_buf)
        lines = text_buf.value.split('\r\n')
        fixed = False
        if self._use_spaces:
            tab_new = ' ' * self._tab_size
            for i, line in enumerate(lines):
                num_tabs = len(line) - len(line.lstrip('\t'))
                if num_tabs:
                    lines[i] = tab_new * num_tabs + lines[i][num_tabs:]
                    fixed = True
        else:
            for i, line in enumerate(lines):
                num_tabs = (len(line) - len(line.lstrip(' '))) // self._tab_size
                if num_tabs:
                    lines[i] = '\t' * num_tabs + lines[i][num_tabs * self._tab_size:]
                    fixed = True

        if fixed:
            txt_new = '\r\n'.join(lines)
            self.edit.send_message(EM_SETSEL, 0, -1)
            self.edit.send_message(EM_REPLACESEL, 1, create_unicode_buffer(txt_new))
            self._check_caret_pos()

    ########################################
    #
    ########################################
    def action_about_notepad(self):
        self.show_message_box(
            tr('ABOUT_TEXT').format(APP_NAME, APP_VERSION, APP_COPYRIGHT),
            tr('ABOUT_CAPTION').format(APP_NAME)
        )
        user32.SetFocus(self.edit.hwnd)

    ########################################
    #
    ########################################
    def action_about_windows(self):
        shell32.ShellAboutW(self.hwnd, create_unicode_buffer('Windows'), None, None)


if __name__ == "__main__":
    app = App(sys.argv[1:])
    sys.exit(app.run())
