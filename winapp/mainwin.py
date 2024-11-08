from ctypes import (windll, WINFUNCTYPE, c_int64, c_int, c_uint, c_uint64, c_long, c_ulong, c_longlong, c_voidp, c_wchar_p, Structure,
        sizeof, byref, create_string_buffer, create_unicode_buffer, cast,  c_char_p, pointer)
from ctypes.wintypes import (HWND, WORD, DWORD, LONG, HICON, WPARAM, LPARAM, HANDLE, LPCWSTR, MSG, UINT, LPWSTR, HINSTANCE,
        LPVOID, INT, RECT, POINT, BYTE, BOOL, COLORREF, LPPOINT)

from .const import *
from .wintypes_extended import *
from .dlls import comdlg32, gdi32, shell32, user32, ACCEL
from .window import *
from .menu import *
from .themes import *
from .dialog import *


VKEY_NAME_MAP = {
    'Del': VK_DELETE,
    'Plus': VK_OEM_PLUS,
    'Minus': VK_OEM_MINUS,
    'Enter': VK_RETURN,
    'Left': VK_LEFT,
    'Right': VK_RIGHT,
}


class MainWin(Window):

    def __init__(self,
        window_title='MyPythonApp',
        window_class='MyPythonAppClass',
        hicon=0,
        left=CW_USEDEFAULT, top=CW_USEDEFAULT, width=CW_USEDEFAULT, height=CW_USEDEFAULT,
        style=WS_OVERLAPPEDWINDOW,
        ex_style=0,
#        color=None,
        hbrush=COLOR_WINDOW + 1,
        bg_brush_dark=BG_BRUSH_DARK,
        menu_data=None,
        menu_mod_translation_table=None,
        accelerators=None,
        cursor=None,
        parent_window=None
    ):

        self.hicon = hicon

        self.__window_title = window_title
        self.__has_app_menus = menu_data is not None
        self.__popup_menus = {}
        self.__timers = {}
        self.__timer_id_counter = 1000
        self.__die = False
        # For asnyc dialogs
        self.__current_dialogs = []

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

#        if type(color) == int:
#            hbrush = color + 1
#        elif type(color) == COLORREF:
#            hbrush = gdi32.CreateSolidBrush(color)
#        elif hbrush is None:
#            hbrush = COLOR_WINDOW + 1

        self.bg_brush_light = hbrush
        self.bg_brush_dark = bg_brush_dark

        newclass = WNDCLASSEX()
        newclass.lpfnWndProc = self.windowproc
        newclass.style = CS_VREDRAW | CS_HREDRAW
        newclass.lpszClassName = window_class
        newclass.hBrush = self.bg_brush_light  #hbrush
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
            hmenu=self.hmenu,
            parent_window=parent_window
        )

        if accelerators:
            accels += accelerators

        if len(accels):
            acc_table = (ACCEL * len(accels))()
            for (i, acc) in enumerate(accels):
                acc_table[i] = ACCEL(TRUE | acc[0], acc[1], acc[2])
            self.__haccel = user32.CreateAcceleratorTableW(acc_table, len(accels))
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
            for dialog in self.__current_dialogs:
                if user32.IsDialogMessageW(dialog.hwnd, byref(msg)):
                    break

            # If the inner loop completes without encountering
            # the break statement then the following else
            # block will be executed and outer loop will continue
            else:

                if not user32.TranslateAcceleratorW(self.hwnd, self.__haccel, byref(msg)):
                    user32.TranslateMessage(byref(msg))
#                    print('>>>', msg.message)
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
        super().apply_theme(is_dark)

        # Update colors of window titlebar
        dwm_use_dark_mode(self.hwnd, is_dark)

        user32.SetClassLongPtrW(self.hwnd, GCL_HBRBACKGROUND, self.bg_brush_dark if is_dark else self.bg_brush_light)

        # Update colors of menus
        uxtheme.SetPreferredAppMode(PreferredAppMode.ForceDark if is_dark else PreferredAppMode.ForceLight)
        uxtheme.FlushMenuThemes()

        if self.__has_app_menus:
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

        self.redraw_window()

    def dialog_show_async(self, dialog):
        self.__current_dialogs.append(dialog)
        dialog._show_async()

    def dialog_show_sync(self, dialog, lparam=0):
        res = dialog._show_sync(lparam=lparam)
        user32.SetActiveWindow(self.hwnd)
        return res

    def _dialog_remove(self, dialog):
        if dialog in self.__current_dialogs:
            self.__current_dialogs.remove(dialog)

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
        ofn.Flags = OFN_ENABLESIZING | OFN_PATHMUSTEXIST
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

#        if not self.is_dark:
#            return user32.MessageBoxW(self.hwnd, text, caption, utype)

        font_name = 'Segoe UI'  # 'MS Shell Dlg'
        font_size = 8

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

        dialog_width = 260 if hicon else (240 if len(btn_ids) > 2 else 190)

        dialog_min_height = 74 if hicon else 62

        text_width = dialog_width - 60 if hicon else dialog_width - 27
        text_x, text_y = 7, 14
        button_width, button_height, button_dist = 53, 14, 5
        margin_right, margin_bottom = 10, 20

        # calulate required height for message text
        text_height = calculate_multiline_text_height(text, text_width, font_name, font_size)

        # if there is an icon and text height is smaller thsn icon height, center text vertically
        if hicon and text_height < 20:
            text_y += (20 - text_height) // 2

        dialog_height = max(dialog_min_height, 54 + text_height)

        dialog_dict = {
            'class': '#32770',
            'caption': caption,
            'font': [font_name, font_size],
            'rect': [0, 0, dialog_width, dialog_height],
            'style': WS_CAPTION | WS_POPUP | WS_VISIBLE | WS_CLIPSIBLINGS | WS_SYSMENU | DS_MODALFRAME | DS_SETFONT | DS_CENTER,
            'exstyle': WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME | WS_EX_WINDOWEDGE,
            'controls': []
        }

        # add icon
        if hicon:
            text_x = 41
            dialog_dict['controls'].append({
                'id': 100,
                'class': 'STATIC',
                'caption': '',
                'rect': [14, 14, 21, 20],
                'style': WS_CHILDWINDOW | WS_GROUP | WS_VISIBLE | SS_ICON,
                'exstyle': WS_EX_NOPARENTNOTIFY
            })

        # add text
        dialog_dict['controls'].append({
            'id': 101,
            'class': 'STATIC',
            'caption': text,
            'rect': [text_x, text_y, text_width, text_height],
            'style': WS_CHILDWINDOW | WS_GROUP | WS_VISIBLE | SS_LEFT | SS_NOPREFIX,  # | SS_EDITCONTROL,
            'exstyle': WS_EX_NOPARENTNOTIFY
        })

        # add button(s)
        x = dialog_width - margin_right - len(btn_ids) * button_width - (len(btn_ids) - 1) * button_dist
        for i in range(len(btn_ids)):
            dialog_dict['controls'].append({
                'id': btn_ids[i],
                'class': 'BUTTON',
                'caption': user32.MB_GetString(btn_ids[i] - 1),
                'rect': [x, dialog_height - margin_bottom, button_width, button_height],
                'style': WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP | BS_TEXT | (BS_DEFPUSHBUTTON if i == 0 else BS_PUSHBUTTON),
                'exstyle': WS_EX_NOPARENTNOTIFY
            })
            x += button_width + button_dist

        def _dialog_proc_callback(hwnd, msg, wparam, lparam):
            if msg == WM_INITDIALOG:
                if hicon:
                    user32.SendMessageW(user32.GetDlgItem(hwnd, 100), STM_SETICON, hicon, 0)

            elif msg == WM_COMMAND:
                control_id = LOWORD(wparam)
                command = HIWORD(wparam)
                if command == BN_CLICKED:
                    user32.EndDialog(hwnd, control_id)
            return FALSE

        return self.dialog_show_sync(Dialog(self, dialog_dict, _dialog_proc_callback))


    def show_message_box_modern(self, text, caption='', utype=MB_ICONINFORMATION | MB_OK):
        font_name = 'Segoe UI'
        font_size = 8

        btn_ids = BUTTON_COMMAND_IDS[utype & 0xf]

        dialog_width = 240 if len(btn_ids) > 2 else 190
        margin_left = 6
        margin_right = 7
        margin_bottom = 20

        text_y = 6
        text_width = dialog_width - margin_left - margin_right
        button_width, button_height, button_dist = 53, 14, 5

        # large font for blue message text
        hfont_message = gdi32.CreateFontW(
            22,
            0, 0, 0,
            FW_NORMAL,
            FALSE,
            FALSE,
            FALSE,
            DEFAULT_CHARSET,
            OUT_OUTLINE_PRECIS,
            CLIP_DEFAULT_PRECIS,
            CLEARTYPE_QUALITY,
            DEFAULT_PITCH | FF_DONTCARE,
            'Segoe UI'
        )

        # calulate required height for message text
        text_height = calculate_multiline_text_height(text, text_width * 8 // 22, hfont=hfont_message) + 2  # ???

        dialog_height = text_height + 49

        dialog_dict = {
            'class': '#32770',
            'caption': caption,
            'font': [font_name, font_size],
            'rect': [0, 0, dialog_width, dialog_height],
            'style': WS_CAPTION | WS_POPUP | WS_VISIBLE | WS_CLIPSIBLINGS | WS_SYSMENU | DS_MODALFRAME | DS_SETFONT | DS_CENTER,
            'exstyle': WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME | WS_EX_WINDOWEDGE,
            'controls': []
        }

        # add text
        dialog_dict['controls'].append({
            'id': 101,
            'class': 'STATIC',
            'caption': text,
            'rect': [margin_left, text_y, text_width, text_height],
            'style': WS_CHILDWINDOW | WS_GROUP | WS_VISIBLE | SS_LEFT | SS_NOPREFIX,
            'exstyle': WS_EX_NOPARENTNOTIFY
        })

        # add button(s)
        x = dialog_width - margin_right - len(btn_ids) * button_width - (len(btn_ids) - 1) * button_dist
        for i in range(len(btn_ids)):
            dialog_dict['controls'].append({
                'id': btn_ids[i],
                'class': 'BUTTON',
                'caption': user32.MB_GetString(btn_ids[i] - 1),
                'rect': [x, dialog_height - margin_bottom, button_width, button_height],
                'style': WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP | BS_TEXT | (BS_DEFPUSHBUTTON if i == 0 else BS_PUSHBUTTON),
                'exstyle': WS_EX_NOPARENTNOTIFY
            })
            x += button_width + button_dist

        def _dialog_proc_callback(hwnd, msg, wparam, lparam):
            if msg == WM_INITDIALOG:
                user32.SendMessageW(user32.GetDlgItem(hwnd, 101), WM_SETFONT, hfont_message, 0)

            elif msg == WM_CTLCOLORSTATIC:
                if self.is_dark:
                    gdi32.SetTextColor(wparam, TEXT_COLOR_DARK)  # light blue: 0xDAA026
                    gdi32.SetBkColor(wparam, BG_COLOR_DARKER)
                    return gdi32.GetStockObject(NULL_BRUSH)
                else:
                    gdi32.SetTextColor(wparam, 0x993300)  # dark blue
                    return gdi32.GetStockObject(DC_BRUSH)

            elif msg == WM_ERASEBKGND:
                rc = RECT()
                user32.GetClientRect(hwnd, byref(rc))
                user32.FillRect(wparam, byref(rc), BG_BRUSH_DARK if self.is_dark else COLOR_3DFACE + 1)
                rc.bottom -= 40
                user32.FillRect(wparam, byref(rc), gdi32.CreateSolidBrush(BG_COLOR_DARKER) if self.is_dark else COLOR_WINDOW + 1)
                rc.top = rc.bottom - 1
                user32.FillRect(wparam, byref(rc), MENU_BG_BRUSH_HOT if self.is_dark else gdi32.CreateSolidBrush(0xdfdfdf))
                return 1

            elif msg == WM_COMMAND:
                control_id = LOWORD(wparam)
                command = HIWORD(wparam)
                if command == BN_CLICKED:
                    user32.EndDialog(hwnd, control_id)
            return FALSE

        res =  self.dialog_show_sync(Dialog(self, dialog_dict, _dialog_proc_callback, handle_static=False))
        gdi32.DeleteObject(hfont_message)
        return res

    def show_font_dialog(self, font_name, font_size, font_weight=FW_DONTCARE, font_italic=False):
        lf = LOGFONTW(
            lfFaceName = font_name,
            lfHeight = -kernel32.MulDiv(font_size, DPI_Y, 72),
            lfCharSet = ANSI_CHARSET,
            lfWeight = font_weight,
            lfItalic = int(font_italic)
        )
        cf = CHOOSEFONTW(
            hwndOwner = self.hwnd,
            lpLogFont = pointer(lf),
            Flags = CF_INITTOLOGFONTSTRUCT | CF_NOSCRIPTSEL
        )
        if comdlg32.ChooseFontW(byref(cf)):
            return (
                lf.lfFaceName,
                kernel32.MulDiv(-lf.lfHeight, 72, DPI_Y) if lf.lfHeight < 0 else lf.lfHeight,
                lf.lfWeight,
                lf.lfItalic > 0
            )

    def show_color_dialog(self, initialColor=COLORREF(0), custom_colors=[]):
        cc = CHOOSECOLORW()
        cc.hwndOwner = self.hwnd
        cc.Flags = CC_SOLIDCOLOR | CC_FULLOPEN | CC_RGBINIT
        cc.lpCustColors = (COLORREF * 16)()
        for i, c in enumerate(custom_colors[:16]):
            cc.lpCustColors[i] = c
        cc.rgbResult = initialColor
        if comdlg32.ChooseColorW(byref(cc)):
            return cc.rgbResult

    @staticmethod
    def __handle_menu_items(hmenu, menu_items, accels=None, key_mod_translation=None):
        for row in menu_items:
            if row is None or row['caption'] == '-':
                user32.AppendMenuW(hmenu, MF_SEPARATOR, 0, '-')
                continue
            if 'items' in row:
                hmenu_child = user32.CreateMenu()
                flags = MF_POPUP
                if 'flags' in row and 'GRAYED' in row['flags']:
                    flags |= MF_GRAYED
                user32.AppendMenuW(hmenu, flags, hmenu_child, row['caption'])

                if 'id' in row or 'hbitmap' in row:
                    info = MENUITEMINFOW()
#                    ok = user32.GetMenuItemInfoW(hmenu, hmenu_child, FALSE, byref(info))
                    info.fMask = 0
                    if 'id' in row:
                        info.wID = row['id'] if 'id' in row else -1
                        info.fMask |= MIIM_ID
                    if 'hbitmap' in row:
                        info.hbmpItem = row['hbitmap']
                        info.fMask |= MIIM_BITMAP
                    user32.SetMenuItemInfoW(hmenu, hmenu_child, FALSE, byref(info))

                MainWin.__handle_menu_items(hmenu_child, row['items'], accels, key_mod_translation)
            else:
#                if row['caption'] == '-':
#                    user32.AppendMenuW(hmenu, MF_SEPARATOR, 0, '-')
#                    continue
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
                        if key_mod_translation and vk.upper() in key_mod_translation:
                            parts[1] = parts[1].replace(vk, key_mod_translation[vk.upper()])
                        vk = VKEY_NAME_MAP[vk] if vk in VKEY_NAME_MAP else eval('VK_' + vk)
                    else:
                        vk = ord(vk)

                    if accels is not None:
                        accels.append((fVirt, vk, id))

                    row['caption'] = '\t'.join(parts)
                user32.AppendMenuW(hmenu, flags, id, row['caption'])

                if 'hbitmap' in row:
                    info = MENUITEMINFOW()
#                    ok = user32.GetMenuItemInfoW(hmenu, id, FALSE, byref(info))
                    info.fMask = MIIM_BITMAP
                    info.hbmpItem = row['hbitmap']
                    user32.SetMenuItemInfoW(hmenu, id, FALSE, byref(info))
