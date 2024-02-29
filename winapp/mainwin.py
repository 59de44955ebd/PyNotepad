from ctypes import (windll, WINFUNCTYPE, c_int64, c_int, c_uint, c_uint64, c_long, c_ulong, c_longlong, c_voidp, c_wchar_p, Structure,
        sizeof, byref, create_string_buffer, create_unicode_buffer, cast,  c_char_p, pointer)
from ctypes.wintypes import (HWND, WORD, DWORD, LONG, HICON, WPARAM, LPARAM, HANDLE, LPCWSTR, MSG, UINT, LPWSTR, HINSTANCE,
        LPVOID, INT, RECT, POINT, BYTE, BOOL, COLORREF, LPPOINT)

from winapp.const import *
from winapp.wintypes_extended import * #LONG_PTR, WNDPROC, MAKELONG, MAKEINTRESOURCEW
from winapp.dlls import comdlg32, gdi32, shell32, user32
from winapp.window import *
from winapp.menu import *
from winapp.themes import *
from winapp.dialog import *  # DialogData
from winapp.controls.button import *


class ACCEL(Structure):
    _fields_ = [
        ("fVirt", BYTE),
        ("key", WORD),
        ("cmd", WORD),
    ]

VKEY_NAME_MAP = {
    'Del': VK_DELETE,
    'Plus': VK_OEM_PLUS,
    'Minus': VK_OEM_MINUS,
}


class MainWin(Window):

    def __init__(self,
            window_title='MyPythonApp',
            window_class='MyPythonAppClass',
            hicon=0,
            left=CW_USEDEFAULT, top=CW_USEDEFAULT, width=CW_USEDEFAULT, height=CW_USEDEFAULT,
            style=WS_OVERLAPPEDWINDOW,
            ex_style=0,
            color=None,
            hbrush=None,
            menu_data=None,
            menu_mod_translation_table=None,
            accelerators=None,
            cursor=None,
            ):

        self.hicon = hicon

        self.__window_title = window_title
        self.__has_app_menus = menu_data is not None
        self.__popup_menus = {}
        self.__timers = {}
        self.__timer_id_counter = 1000
        self.__die = False
        # For asnyc dialog (only one at a time supported)
        self.__hwnd_dialog = None
        self.__dialogproc = None

        def _on_WM_TIMER(hwnd, wparam, lparam):
            if wparam in self.__timers:
                callback = self.__timers[wparam][0]
                if self.__timers[wparam][1]:
                    user32.KillTimer(self.hwnd, wparam)
                    del self.__timers[wparam]
                callback()
            # An application should return zero if it processes this message.
            return 0

        self.__message_map = {
            WM_TIMER:        [_on_WM_TIMER],
            WM_CLOSE:        [self.quit],
        }

        def _window_proc_callback(hwnd, msg, wparam, lparam):
            if msg in self.__message_map:
                for callback in self.__message_map[msg]:
                    res = callback(hwnd, wparam, lparam)
                    if res is not None:
                        return res
            return user32.DefWindowProcW(hwnd, msg, wparam, lparam)

        self.windowproc = WNDPROC(_window_proc_callback)

        if hbrush is None:
            hbrush = COLOR_WINDOW + 1
        elif type(color) == int:
            hbrush = color + 1
        elif type(color) == COLORREF:
            hbrush = gdi32.CreateSolidBrush(color)
        self.bg_brush_light = hbrush

        newclass = WNDCLASSEX()
        newclass.lpfnWndProc = self.windowproc
        newclass.style = CS_VREDRAW | CS_HREDRAW
        newclass.lpszClassName = window_class
        newclass.hBrush = hbrush
        newclass.hCursor = user32.LoadCursorW(0, cursor if cursor else IDC_ARROW)
        newclass.hIcon = self.hicon

        accels = []

        if menu_data:
            self.hmenu = user32.CreateMenu()
            MainWin.__handle_menu_items(self.hmenu, menu_data['items'], accels, menu_mod_translation_table)
        else:
            self.hmenu = 0

        user32.RegisterClassExW(byref(newclass))

        super().__init__(
            newclass.lpszClassName,
            style=style,
            ex_style=ex_style,
            left=left, top=top, width=width, height=height,
            window_title=window_title,
            hmenu=self.hmenu
        )

        if accelerators:
            accels += accelerators

        if len(accels):
            acc_table = (ACCEL * len(accels))()
            for (i, acc) in enumerate(accels):
                acc_table[i] = ACCEL(TRUE | acc[0], acc[1], acc[2])
            self.__haccel = user32.CreateAcceleratorTableW(acc_table[0], len(accels))
        else:
            self.__haccel = None

    def make_popup_menu(self, menu_data):
        hmenu = user32.CreatePopupMenu()
        MainWin.__handle_menu_items(hmenu, menu_data['items'])
        return hmenu

    def get_dropped_items(self, hdrop):
        dropped_items = []
        cnt = shell32.DragQueryFileW(hdrop, 0xFFFFFFFF, None, 0)
        for i in range(cnt):
            file_buffer = create_unicode_buffer('', MAX_PATH)
            shell32.DragQueryFileW(hdrop, i, file_buffer, MAX_PATH)
            dropped_items.append(file_buffer[:].split('\0', 1)[0])
        shell32.DragFinish(hdrop)
        return dropped_items

    def create_timer(self, callback, ms, is_singleshot=False, timer_id=None):
        if timer_id is None:
            timer_id = self.__timer_id_counter
            self.__timer_id_counter += 1
        self.__timers[timer_id] = (callback, is_singleshot)
        user32.SetTimer(self.hwnd, timer_id, ms, 0)
        return timer_id

    def kill_timer(self, timer_id):
        if timer_id in self.__timers:
            user32.KillTimer(self.hwnd, timer_id)
            del self.__timers[timer_id]

    def register_message_callback(self, msg, callback, overwrite=False):
        if overwrite:
            self.__message_map[msg] = [callback]
        else:
            if msg not in self.__message_map:
                self.__message_map[msg] = []
            self.__message_map[msg].append(callback)
        if msg == WM_DROPFILES:
            shell32.DragAcceptFiles(self.hwnd, True)

    def unregister_message_callback(self, msg, callback=None):
        if msg in self.__message_map:
            if callback is None:  # was: == True
                del self.__message_map[msg]
            elif callback in self.__message_map[msg]:
                self.__message_map[msg].remove(callback)
                if len(self.__message_map[msg]) == 0:
                    del self.__message_map[msg]

    def run(self):
        msg = MSG()
        while not self.__die and user32.GetMessageW(byref(msg), 0, 0, 0) != 0:

            # unfortunately this disables global accelerators while a dialog is shown
            if self.__hwnd_dialog and user32.IsDialogMessageW(self.__hwnd_dialog, byref(msg)):
                continue

            if not user32.TranslateAcceleratorW(self.hwnd, self.__haccel, byref(msg)):
                user32.TranslateMessage(byref(msg))
                user32.DispatchMessageW(byref(msg))

        if self.__haccel:
            user32.DestroyAcceleratorTable(self.__haccel)
        user32.DestroyWindow(self.hwnd)
        user32.DestroyIcon(self.hicon)
        return 0

    def quit(self, *args):
        self.__die = True
        user32.PostMessageW(self.hwnd, WM_NULL, 0, 0)

    def apply_theme(self, is_dark):
        self.is_dark = is_dark

        # Update colors of window titlebar
        dwm_use_dark_mode(self.hwnd, is_dark)

        user32.SetClassLongPtrW(self.hwnd, GCL_HBRBACKGROUND, BG_BRUSH_DARK if is_dark else self.bg_brush_light)

        # Updating existing dialog is too complicate, so trigger decent close
        if self.__hwnd_dialog:
            user32.PostMessageW(self.__hwnd_dialog, WM_CLOSE, 0, 0)

        if self.__has_app_menus:
            # Update colors of menus
            uxtheme.SetPreferredAppMode(PreferredAppMode.ForceDark if is_dark else PreferredAppMode.ForceLight)
            uxtheme.FlushMenuThemes()
            self.redraw_window()

            # Update colors of menubar
            if is_dark:
                def _on_WM_UAHDRAWMENU(hwnd, wparam, lparam):
                    pUDM = cast(lparam, POINTER(UAHMENU)).contents
                    mbi = MENUBARINFO()
                    ok = user32.GetMenuBarInfo(hwnd, OBJID_MENU, 0, byref(mbi))
                    rc_win = RECT()
                    user32.GetWindowRect(hwnd, byref(rc_win))
                    rc = mbi.rcBar
                    user32.OffsetRect(byref(rc), -rc_win.left, -rc_win.top)
                    res = user32.FillRect(pUDM.hdc, byref(rc), MENUBAR_BG_BRUSH_DARK)
                    return TRUE
                self.register_message_callback(WM_UAHDRAWMENU, _on_WM_UAHDRAWMENU)

                def _on_WM_UAHDRAWMENUITEM(hwnd, wparam, lparam):
                    pUDMI = cast(lparam, POINTER(UAHDRAWMENUITEM)).contents
                    mii = MENUITEMINFOW()
                    mii.fMask = MIIM_STRING
                    buf = create_unicode_buffer('', 256)
                    mii.dwTypeData = cast(buf, LPWSTR)
                    mii.cch = 256
                    ok = user32.GetMenuItemInfoW(pUDMI.um.hmenu, pUDMI.umi.iPosition, TRUE, byref(mii))
                    if pUDMI.dis.itemState & ODS_HOTLIGHT or pUDMI.dis.itemState & ODS_SELECTED:
                        user32.FillRect(pUDMI.um.hdc, byref(pUDMI.dis.rcItem), MENU_BG_BRUSH_HOT)
                    else:
                        user32.FillRect(pUDMI.um.hdc, byref(pUDMI.dis.rcItem), MENUBAR_BG_BRUSH_DARK)
                    gdi32.SetBkMode(pUDMI.um.hdc, TRANSPARENT)
                    gdi32.SetTextColor(pUDMI.um.hdc, TEXT_COLOR_DARK)
                    user32.DrawTextW(pUDMI.um.hdc, mii.dwTypeData, len(mii.dwTypeData), byref(pUDMI.dis.rcItem), DT_CENTER | DT_SINGLELINE | DT_VCENTER)
                    return TRUE
                self.register_message_callback(WM_UAHDRAWMENUITEM, _on_WM_UAHDRAWMENUITEM)

                def UAHDrawMenuNCBottomLine(hwnd, wparam, lparam):
                    rcClient = RECT()
                    user32.GetClientRect(hwnd, byref(rcClient))
                    user32.MapWindowPoints(hwnd, None, byref(rcClient), 2)
                    rcWindow = RECT()
                    user32.GetWindowRect(hwnd, byref(rcWindow))
                    user32.OffsetRect(byref(rcClient), -rcWindow.left, -rcWindow.top)
                    # the rcBar is offset by the window rect
                    rcAnnoyingLine = rcClient
                    rcAnnoyingLine.bottom = rcAnnoyingLine.top
                    rcAnnoyingLine.top -= 1
                    hdc = user32.GetWindowDC(hwnd)
                    user32.FillRect(hdc, byref(rcAnnoyingLine), BG_BRUSH_DARK)
                    user32.ReleaseDC(hwnd, hdc)

                def _on_WM_NCPAINT(hwnd, wparam, lparam):
                    user32.DefWindowProcW(hwnd, WM_NCPAINT, wparam, lparam)
                    UAHDrawMenuNCBottomLine(hwnd, wparam, lparam)
                    return TRUE
                self.register_message_callback(WM_NCPAINT, _on_WM_NCPAINT)

                def _on_WM_NCACTIVATE(hwnd, wparam, lparam):
                    user32.DefWindowProcW(hwnd, WM_NCACTIVATE, wparam, lparam)
                    UAHDrawMenuNCBottomLine(hwnd, wparam, lparam)
                    return TRUE
                self.register_message_callback(WM_NCACTIVATE, _on_WM_NCACTIVATE)

            else:
                self.unregister_message_callback(WM_UAHDRAWMENU)
                self.unregister_message_callback(WM_UAHDRAWMENUITEM)
                self.unregister_message_callback(WM_NCPAINT)
                self.unregister_message_callback(WM_NCACTIVATE)

    def load_dialog_data(self, data):
        dialog = DialogData()
        for control in data['controls']:
            dialog.add_control(
                    eval(control['class']),
                    control['caption'],
                    *control['rect'],
                    control_id=control['id'],
                    style=control['style']
                    )
        dlg_data = dialog.create(
                *data['rect'],
                data['caption'],
                *data['font'],
                False,
                style=data['style'],
                exstyle=data['exstyle'] if 'exstyle' in data else 0
                )
        return (c_ubyte * len(dlg_data))(*dlg_data)

    def show_dialog_async(self, dlg_data, dialog_proc_callback):

        def _dialog_proc_callback(hwnd, msg, wparam, lparam):
            if self.is_dark:
                res = self.__dialog_handle_dark(hwnd, msg, wparam, lparam)
                if res is not None:
                    return res
            if msg == WM_CLOSE:
                self.close_dialog_async()
            return dialog_proc_callback(hwnd, msg, wparam, lparam)

        self.__dialogproc = WNDPROC(_dialog_proc_callback)
        self.__hwnd_dialog = user32.CreateDialogIndirectParamW(
                0,
                byref(dlg_data),
                self.hwnd,
                self.__dialogproc,
                0
                )
        user32.SendMessageW(self.__hwnd_dialog, WM_CHANGEUISTATE, MAKELONG(UIS_CLEAR, UISF_HIDEFOCUS), 0)
        user32.ShowWindow(self.__hwnd_dialog, SW_SHOW)

    def close_dialog_async(self):
        user32.DestroyWindow(self.__hwnd_dialog)
        self.__hwnd_dialog = None
        self.__dialogproc = None

    def show_dialog_sync(self, dlg_data, dialog_proc_callback):

        def _dialog_proc_callback(hwnd, msg, wparam, lparam):
            if self.is_dark:
                res = self.__dialog_handle_dark(hwnd, msg, wparam, lparam)
                if res is not None:
                    return res
            if msg == WM_CLOSE:
                 user32.EndDialog(hwnd, 0)
            return dialog_proc_callback(hwnd, msg, wparam, lparam)

        return user32.DialogBoxIndirectParamW(
                0,
                byref(dlg_data),
                self.hwnd,
                WNDPROC(_dialog_proc_callback),
                0
                )

    def get_open_filename(self, title, default_extension='',
                filter_string='All Files (*.*)\0*.*\0\0', initial_path=''):
        file_buffer = create_unicode_buffer(initial_path, MAX_PATH)
        ofn = OPENFILENAMEW()
        ofn.hwndOwner = self.hwnd
        ofn.lStructSize = sizeof(OPENFILENAMEW)
        ofn.lpstrTitle = title
        ofn.lpstrFile = cast(file_buffer, LPWSTR)
        ofn.nMaxFile = MAX_PATH
        ofn.lpstrDefExt = default_extension
        ofn.lpstrFilter = cast(create_unicode_buffer(filter_string), c_wchar_p)
        ofn.Flags = OFN_ENABLESIZING | OFN_PATHMUSTEXIST #| OFN_EXPLORER # | OFN_ENABLEHOOK
        if comdlg32.GetOpenFileNameW(byref(ofn)):
            return file_buffer[:].split('\0', 1)[0]
        else:
            return None

    def get_save_filename(self, title, default_extension='',
                filter_string='All Files (*.*)\0*.*\0\0', initial_path=''):
        file_buffer = create_unicode_buffer(initial_path, MAX_PATH)
        ofn = OPENFILENAMEW()
        ofn.hwndOwner = self.hwnd
        ofn.lStructSize = sizeof(OPENFILENAMEW)
        ofn.lpstrTitle = title
        ofn.lpstrFile = cast(file_buffer, LPWSTR)
        ofn.nMaxFile = MAX_PATH
        ofn.lpstrDefExt = default_extension
        ofn.lpstrFilter = cast(create_unicode_buffer(filter_string), c_wchar_p)
        ofn.Flags = OFN_ENABLESIZING | OFN_OVERWRITEPROMPT
        if comdlg32.GetSaveFileNameW(byref(ofn)):
            return file_buffer[:].split('\0', 1)[0]
        else:
            return None

    def show_message_box(self, text, caption='', utype=MB_ICONINFORMATION | MB_OK):
        font = ['MS Shell Dlg', 8]

#        if not self.is_dark:
#            return user32.MessageBoxW(self.hwnd, text, caption, utype)

        if utype & MB_ICONINFORMATION:
            hicon = get_stock_icon(SIID_INFO)
        elif utype & MB_ICONWARNING:
            hicon = get_stock_icon(SIID_WARNING)
        elif utype & MB_ICONERROR:
            hicon = get_stock_icon(SIID_ERROR)
        elif utype & MB_ICONQUESTION:
            hicon = get_stock_icon(SIID_HELP)
        else:
            hicon = None

        btn_ids = BUTTON_COMMAND_IDS[utype & 0xf]

        dialog_width = 222 if hicon else (219 if len(btn_ids) > 2 else 189)  # 219 => 197
        dialog_min_height = 74 if hicon else 62

        text_width = dialog_width - 60 if hicon else dialog_width - 27
        text_x, text_y = 7, 14
        button_width, button_height, button_dist = 53, 14, 5
        margin_right, margin_bottom = 10, 20

        # calulate required height for message text
        text_height = self.dialog_calculate_text_height(text, text_width, *font)

        # if there is an icon and text height is smaller thsn icon height, center text vertically
        if hicon and text_height < 20:
            text_y += (20 - text_height) // 2

        dialog_height = max(dialog_min_height, 54 + text_height)

        dialog_dict = {'class': '#32770', 'caption': caption, 'font': font,
                'rect': [0, 0, dialog_width, dialog_height],
                'style': 2496137669, 'exstyle': 65793, 'controls': []
                }

        # add icon
        if hicon:
            text_x = 41
            dialog_dict['controls'].append({
                    'id': 20, 'class': 'STATIC',
                    'caption': '', 'rect': [14, 14, 21, 20],
                    'style': 1342308355, 'exstyle': 4})

        # add text
        dialog_dict['controls'].append({
                'id': 65535, 'class': 'STATIC', 'caption': text,
                'rect': [text_x, text_y, text_width, text_height],
                'style': 1342316672, 'exstyle': 4})

        # add button(s)
        x = dialog_width - margin_right - len(btn_ids) * button_width - (len(btn_ids) - 1) * button_dist
        for i in range(len(btn_ids)):
            dialog_dict['controls'].append({
                'id': btn_ids[i], 'class': 'BUTTON',
                'caption': user32.MB_GetString(btn_ids[i] - 1),
                'rect': [x, dialog_height - margin_bottom, button_width, button_height],
                'style': WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP | BS_TEXT | (BS_DEFPUSHBUTTON if i == 0 else BS_PUSHBUTTON),
                'exstyle': 4
            })
            x += button_width + button_dist

        dlg_data = self.load_dialog_data(dialog_dict)

        def _dialog_proc_callback(hwnd, msg, wparam, lparam):
            if msg == WM_INITDIALOG:
                if hicon:
                    user32.SendMessageW(user32.GetDlgItem(hwnd, 20), STM_SETICON, hicon, 0)
                if self.is_dark:
                    dwm_use_dark_mode(hwnd, True)
                    for btn_id in btn_ids:
                        uxtheme.SetWindowTheme(user32.GetDlgItem(hwnd, btn_id), 'DarkMode_Explorer', None)
                user32.SetFocus(hwnd)
            elif msg == WM_COMMAND:
                control_id = LOWORD(wparam)
                command = HIWORD(wparam)
                if command == BN_CLICKED:
                    user32.EndDialog(hwnd, control_id)
            elif msg == WM_CLOSE:
                user32.EndDialog(hwnd, 0)
            elif self.is_dark:
                if msg == WM_CTLCOLORDLG or msg == WM_CTLCOLORSTATIC:
                    hdc = wparam
                    gdi32.SetTextColor(hdc, TEXT_COLOR_DARK)
                    gdi32.SetBkColor(hdc, BG_COLOR_DARK)
                    return BG_BRUSH_DARK
                elif msg == WM_CTLCOLORBTN:
                    gdi32.SetDCBrushColor(wparam, BG_COLOR_DARK)
                    return gdi32.GetStockObject(DC_BRUSH)
            return FALSE

        res = user32.DialogBoxIndirectParamW(
                0,
                byref(dlg_data),
                self.hwnd,
                WNDPROC(_dialog_proc_callback),
                0
                )

        user32.SetActiveWindow(self.hwnd)
        return IDCANCEL if res == 0 else res

    def show_font_dialog(self, font_name, font_size, font_weight=FW_DONTCARE, font_italic=False):
        lf = LOGFONTW(
                lfFaceName = font_name,
                lfHeight = -kernel32.MulDiv(font_size, DPI_Y, 72),
                lfCharSet = ANSI_CHARSET,
                lfWeight = font_weight,
                lfItalic = int(font_italic))
        cf = CHOOSEFONTW(
                hwndOwner = self.hwnd,
                lpLogFont = pointer(lf),
                Flags = CF_INITTOLOGFONTSTRUCT)
        if comdlg32.ChooseFontW(byref(cf)):
            return (
                    lf.lfFaceName,
                    kernel32.MulDiv(-lf.lfHeight, 72, DPI_Y) if lf.lfHeight < 0 else lf.lfHeight,
                    lf.lfWeight,
                    lf.lfItalic)

    # logical coordinates, not pixels
    def dialog_calculate_text_height(self, text, text_width, font_name='MS Shell Dlg', font_size=8):
        hdc = user32.GetDC(0)
        hfont = gdi32.CreateFontW(font_size, 0, 0, 0, FW_DONTCARE, FALSE, FALSE, FALSE, ANSI_CHARSET, OUT_TT_PRECIS,
                CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, font_name)
        gdi32.SelectObject(hdc, hfont)
        rc = RECT(0, 0, text_width, 0)
        user32.DrawTextW(hdc, text, -1, byref(rc), DT_CALCRECT | DT_LEFT | DT_TOP | DT_WORDBREAK | DT_NOPREFIX)
        user32.ReleaseDC(0, hdc)
        return rc.bottom

    def __dialog_handle_dark(self, hwnd_dialog, msg, wparam, lparam):
        if msg == WM_INITDIALOG:
            dwm_use_dark_mode(hwnd_dialog, True)

            controls = []
            def _enum_child_func(hwnd, lparam):
                controls.append(hwnd)
                return TRUE

            EnumWindowsProc = WINFUNCTYPE(BOOL, HWND, LPARAM)
            user32.EnumChildWindows(hwnd_dialog, EnumWindowsProc(_enum_child_func), 0)

            h_font = user32.SendMessageW(hwnd_dialog, WM_GETFONT, 0, 0)

            for hwnd in controls:
                buf = create_unicode_buffer(32)
                user32.GetClassNameW(hwnd, buf, 32)
                window_class = buf.value

                if window_class == 'Button':
                    uxtheme.SetWindowTheme(hwnd, 'DarkMode_Explorer', None)

                    style = user32.GetWindowLongPtrA(hwnd, GWL_STYLE)

                    if style & BS_TYPEMASK == BS_AUTOCHECKBOX or style & BS_TYPEMASK == BS_AUTORADIOBUTTON:
                        rc = RECT()
                        user32.GetClientRect(hwnd, byref(rc))

                        buf = create_unicode_buffer(64)
                        user32.GetWindowTextW(hwnd, buf, 64)
                        window_title = buf.value.replace('&', '')
                        user32.SetWindowTextW(hwnd, window_title)

                        window_checkbox = Window('Button', wrap_hwnd=hwnd)

                        checkbox_static = Static(
                                window_checkbox,
                                style=WS_CHILD | SS_SIMPLE | WS_VISIBLE,
                                ex_style=WS_EX_TRANSPARENT,
                                left=16,
                                top=3,
                                width=rc.right - 16,
                                height=rc.bottom,
                                window_title=window_title
                                )
                        user32.SendMessageW(checkbox_static.hwnd, WM_SETFONT, h_font, 0)

                        def _on_WM_CTLCOLORSTATIC(hwnd, wparam, lparam):
                            gdi32.SetTextColor(wparam, TEXT_COLOR_DARK)
                            gdi32.SetBkMode(wparam, TRANSPARENT)
                            return gdi32.GetStockObject(DC_BRUSH)
                        window_checkbox.register_message_callback(WM_CTLCOLORSTATIC, _on_WM_CTLCOLORSTATIC)

                    elif style & BS_TYPEMASK == BS_GROUPBOX:
                        rc = RECT()
                        user32.GetWindowRect(hwnd, byref(rc))
                        user32.ScreenToClient(hwnd_dialog, byref(rc))

                        buf = create_unicode_buffer(64)
                        user32.GetWindowTextW(hwnd, buf, 64)
                        window_title = buf.value

                        hwnd_static = user32.CreateWindowExW(
                                WS_EX_TRANSPARENT,
                                WC_STATIC,
                                window_title,
                                WS_CHILD | SS_SIMPLE | WS_VISIBLE,
                                rc.left + 10, rc.top,
                                rc.right - rc.left - 16, rc.bottom - rc.top,
                                hwnd_dialog,
                                0, 0, 0
                                )
                        user32.SendMessageW(hwnd_static, WM_SETFONT, h_font, 0)

                elif window_class == 'Edit':
                    user32.SetWindowLongPtrA(hwnd, GWL_EXSTYLE,
                            user32.GetWindowLongPtrA(hwnd, GWL_EXSTYLE) & ~WS_EX_CLIENTEDGE)
                    user32.SetWindowLongPtrA(hwnd, GWL_STYLE,
                            user32.GetWindowLongPtrA(hwnd, GWL_STYLE) | WS_BORDER)
                    user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                            SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED)

        elif msg == WM_CTLCOLORDLG or msg == WM_CTLCOLORSTATIC:
            hdc = wparam
            gdi32.SetTextColor(hdc, TEXT_COLOR_DARK)
            gdi32.SetBkColor(hdc, BG_COLOR_DARK)
            return BG_BRUSH_DARK

        elif msg == WM_CTLCOLORBTN:
            gdi32.SetDCBrushColor(wparam, BG_COLOR_DARK)
            return gdi32.GetStockObject(DC_BRUSH)

        elif msg == WM_CTLCOLOREDIT:
            gdi32.SetTextColor(wparam, TEXT_COLOR_DARK)
            gdi32.SetBkColor(wparam, CONTROL_COLOR_DARK)
            gdi32.SetDCBrushColor(wparam, CONTROL_COLOR_DARK)
            return gdi32.GetStockObject(DC_BRUSH)

    @staticmethod
    def __handle_menu_items(hmenu, menu_items, accels=None, key_mod_translation=None):
        for row in menu_items:
            if 'items' in row:
                hmenu_child = user32.CreateMenu()
                user32.AppendMenuW(hmenu, MF_POPUP, hmenu_child, row['caption'])
                MainWin.__handle_menu_items(hmenu_child, row['items'], accels, key_mod_translation)
            else:
                if row['caption'] == '-':
                    user32.AppendMenuW(hmenu, MF_SEPARATOR, 0, '-')
                    continue
                id = row['id'] if 'id' in row else None
                flags = MF_STRING
                if 'flags' in row:
                    if 'CHECKED' in row['flags']:
                        flags |= MF_CHECKED
                    if 'GRAYED' in row['flags']:
                        flags |= MF_GRAYED
                if '\t' in row['caption']:
                    parts = row['caption'].split('\t') #[1]
                    vk = parts[1]
                    fVirt = 0
                    if 'Alt+' in vk:
                        fVirt |= FALT
                        vk = vk.replace('Alt+', '')
                        if key_mod_translation and 'ALT' in key_mod_translation:
                            parts[1] = parts[1].replace('Alt', key_mod_translation['ALT'])
                    if 'Ctrl+' in vk:
                        fVirt |= FCONTROL
                        vk = vk.replace('Ctrl+', '')
                        if key_mod_translation and 'CTRL' in key_mod_translation:
                            parts[1] = parts[1].replace('Ctrl', key_mod_translation['CTRL'])
                    if 'Shift+' in vk:
                        fVirt |= FSHIFT
                        vk = vk.replace('Shift+', '')
                        if key_mod_translation and 'SHIFT' in key_mod_translation:
                            parts[1] = parts[1].replace('Shift', key_mod_translation['SHIFT'])
                    if len(vk) > 1:
                        vk = VKEY_NAME_MAP[vk] if vk in VKEY_NAME_MAP else eval('VK_' + vk)
                    else:
                        vk = ord(vk)
                    if accels is not None:
                        accels.append((fVirt, vk, id))
                    row['caption'] = '\t'.join(parts)
                user32.AppendMenuW(hmenu, flags, id, row['caption'])


if __name__ == "__main__":
    import sys
    app = MainWin(window_title='Hello World!')
    app.show()
    sys.exit(app.run())
