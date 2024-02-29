from ctypes import windll, c_uint, POINTER, c_int, c_void_p, c_wchar_p, Structure
from ctypes.wintypes import *

#from comtypes import GUID, HRESULT

from winapp.wintypes_extended import WNDPROC, LONG_PTR, ENUMRESNAMEPROCW
#from winapp.shellapi import PIDL, IDropTarget, SHFILEINFOW, IShellItem, REFIID

advapi32 = windll.Advapi32
comctl32 = windll.Comctl32
comdlg32 = windll.comdlg32
gdi32 = windll.Gdi32
kernel32 = windll.Kernel32
ole32 = windll.Ole32
shell32 = windll.shell32
shlwapi = windll.Shlwapi
user32 = windll.user32
uxtheme = windll.UxTheme

########################################
# advapi32
########################################
#HKEY_CURRENT_USER = c_void_p(0x80000001)

LSTATUS = LONG
PHKEY = POINTER(HKEY)

#LSTATUS RegOpenKeyW(
#  [in]           HKEY    hKey,
#  [in, optional] LPCWSTR lpSubKey,
#  [out]          PHKEY   phkResult
#);
advapi32.RegOpenKeyW.argtypes = (HKEY, LPCWSTR, PHKEY)
advapi32.RegOpenKeyW.restype = LSTATUS

#LSTATUS RegCreateKeyW(
#  [in]           HKEY    hKey,
#  [in, optional] LPCWSTR lpSubKey,
#  [out]          PHKEY   phkResult
#);
advapi32.RegCreateKeyW.argtypes = (HKEY, LPCWSTR, PHKEY)
advapi32.RegCreateKeyW.restype = LSTATUS

#LSTATUS RegCloseKey(
#  [in] HKEY hKey
#);
advapi32.RegCloseKey.argtypes = (HKEY, )
advapi32.RegCloseKey.restype = LSTATUS

#LSTATUS RegQueryValueExW(
#  [in]                HKEY    hKey,
#  [in, optional]      LPCWSTR lpValueName,
#                      LPDWORD lpReserved,
#  [out, optional]     LPDWORD lpType,
#  [out, optional]     LPBYTE  lpData,
#  [in, out, optional] LPDWORD lpcbData
#);
advapi32.RegQueryValueExW.argtypes = (HKEY, LPCWSTR, POINTER(DWORD), POINTER(DWORD), c_void_p, POINTER(DWORD))
advapi32.RegQueryValueExW.restype = LSTATUS

#LSTATUS RegSetValueExW(
#  [in]           HKEY       hKey,
#  [in, optional] LPCWSTR    lpValueName,
#                 DWORD      Reserved,
#  [in]           DWORD      dwType,
#  [in]           const BYTE *lpData,
#  [in]           DWORD      cbData
#);
advapi32.RegSetValueExW.argtypes = (HKEY, LPCWSTR, DWORD, DWORD, c_void_p, DWORD)  # POINTER(BYTE)
advapi32.RegSetValueExW.restype = LSTATUS

########################################
# comctl32
########################################
#int ImageList_Add(
#  [in]           HIMAGELIST himl,
#  [in]           HBITMAP    hbmImage,
#  [in, optional] HBITMAP    hbmMask
#);
comctl32.ImageList_Add.argtypes = (HANDLE, HANDLE, HANDLE)

comctl32.ImageList_AddIcon.argtypes = (HANDLE, HANDLE)

comctl32.ImageList_BeginDrag.argtypes = (HANDLE, INT, INT, INT)
comctl32.ImageList_Create.restype = HANDLE

comctl32.ImageList_Destroy.argtypes = (HANDLE,)

#BOOL ImageList_Draw(
#  HIMAGELIST himl,
#  int        i,
#  HDC        hdcDst,
#  int        x,
#  int        y,
#  UINT       fStyle
#);
comctl32.ImageList_Draw.argtypes = (HANDLE, INT, HANDLE, INT, INT, UINT)

comctl32.ImageList_GetIcon.argtypes = (HANDLE, INT, UINT)
comctl32.ImageList_GetIcon.restype = HICON

#BOOL ImageList_GetIconSize(
#  HIMAGELIST himl,
#  int        *cx,
#  int        *cy
#);
comctl32.ImageList_GetIconSize.argtypes = (HANDLE, POINTER(INT), POINTER(INT))

comctl32.ImageList_GetImageCount.argtypes = (HANDLE, )

comctl32.ImageList_GetImageInfo.argtypes = (HANDLE, INT, LPVOID) #POINTER(IMAGEINFO)]

#HIMAGELIST ImageList_LoadImageW(
#  HINSTANCE hi,
#  LPCWSTR   lpbmp,
#  int       cx,
#  int       cGrow,
#  COLORREF  crMask,
#  UINT      uType,
#  UINT      uFlags
#);
comctl32.ImageList_LoadImageW.argtypes = (HANDLE, LPCWSTR, INT, INT, COLORREF, UINT, UINT)
comctl32.ImageList_LoadImageW.restype = HANDLE

comctl32.ImageList_Merge.argtypes = (HANDLE, INT, HANDLE, INT, INT, INT)
comctl32.ImageList_Merge.restype = HANDLE

comctl32.ImageList_Remove.argtypes = (HANDLE, INT)

comctl32.ImageList_Replace.argtypes = (HANDLE, INT, HANDLE, HANDLE)

comctl32.ImageList_ReplaceIcon.argtypes = (HANDLE, INT, HICON)

########################################
# gdi32
########################################
#BOOL BitBlt(
#  [in] HDC   hdc,
#  [in] int   x,
#  [in] int   y,
#  [in] int   cx,
#  [in] int   cy,
#  [in] HDC   hdcSrc,
#  [in] int   x1,
#  [in] int   y1,
#  [in] DWORD rop
#);
gdi32.BitBlt.argtypes = (HANDLE, INT, INT, INT, INT, HANDLE, INT, INT, DWORD)

#HBITMAP CreateBitmap(
#  [in] int        nWidth,
#  [in] int        nHeight,
#  [in] UINT       nPlanes,
#  [in] UINT       nBitCount,
#  [in] const VOID *lpBits
#);
gdi32.CreateBitmap.argtypes = (INT, INT, UINT, UINT, LPVOID)
gdi32.CreateBitmap.restype = HANDLE

#HBITMAP CreateBitmapIndirect(
#  [in] const BITMAP *pbm
#);
gdi32.CreateBitmapIndirect.restype = HANDLE

gdi32.CreateCompatibleDC.argtypes = (HANDLE, )
gdi32.CreateCompatibleDC.restype = HANDLE

gdi32.CreateCompatibleBitmap.argtypes = (HANDLE, INT, INT)

#    HBITMAP CreateDIBitmap(
#      [in] HDC                    hdc,
#      [in] const BITMAPINFOHEADER *pbmih,
#      [in] DWORD                  flInit,
#      [in] const VOID             *pjBits,
#      [in] const BITMAPINFO       *pbmi,
#      [in] UINT                   iUsage
#    );
gdi32.CreateDIBitmap.argtypes = (HANDLE, c_void_p, DWORD, LPVOID, c_void_p, UINT)
gdi32.CreateDIBitmap.restype = HANDLE

gdi32.CreateSolidBrush.argtypes = (COLORREF, )
gdi32.CreateSolidBrush.restype = HANDLE

gdi32.DeleteDC.argtypes = (HANDLE, )

gdi32.DeleteObject.argtypes = (HANDLE, )

#BOOL ExtTextOutW(
#  [in] HDC        hdc,
#  [in] int        x,
#  [in] int        y,
#  [in] UINT       options,
#  [in] const RECT *lprect,
#  [in] LPCWSTR    lpString,
#  [in] UINT       c,
#  [in] const INT  *lpDx
#);
gdi32.ExtTextOutW.argtypes = (HANDLE, INT, INT, UINT, POINTER(RECT), LPCWSTR, UINT, POINTER(INT))

gdi32.GetDIBits.argtypes = (HDC, HBITMAP, UINT, UINT, LPVOID, c_void_p, UINT)
gdi32.GetDIBits.restype = c_int

gdi32.GetDeviceCaps.argtypes = (HDC, INT)

#int GetObject(
#  [in]  HANDLE h,
#  [in]  int    c,
#  [out] LPVOID pv
#);
gdi32.GetObjectW.argtypes = (HANDLE, INT, LPVOID)

#gdi32.GetStockBrush

gdi32.GetTextMetricsW.argtypes = (HANDLE, LPVOID)

#BOOL MaskBlt(
#  [in] HDC     hdcDest,
#  [in] int     xDest,
#  [in] int     yDest,
#  [in] int     width,
#  [in] int     height,
#  [in] HDC     hdcSrc,
#  [in] int     xSrc,
#  [in] int     ySrc,
#  [in] HBITMAP hbmMask,
#  [in] int     xMask,
#  [in] int     yMask,
#  [in] DWORD   rop
#);
gdi32.MaskBlt.argtypes = (HANDLE, INT, INT, INT, INT, HANDLE, INT, INT, HANDLE, INT, INT, DWORD)
#(hdcDest, 0, 0, 16, 16, hdcSrc, 0, 0, icon_info.hbmMask, 0, 0,  rop)

gdi32.SelectObject.argtypes = (HANDLE, HANDLE) #(hdcSrc, icon_info.hbmColor)

gdi32.SetBkColor.argtypes = (HANDLE, COLORREF)

gdi32.SetBkMode.argtypes = (HANDLE, INT)

gdi32.SetDCBrushColor.argtypes = (HANDLE, COLORREF)
gdi32.SetDCPenColor.argtypes = (HANDLE, COLORREF)

gdi32.SetTextColor.argtypes = (HANDLE, COLORREF)

#BOOL SetViewportExtEx(
#  [in]  HDC    hdc,
#  [in]  int    x,
#  [in]  int    y,
#  [out] LPSIZE lpsz
#);
gdi32.SetViewportExtEx.argtypes = (HANDLE, INT, INT, POINTER(SIZE))

#BOOL SetWindowExtEx(
#  [in]  HDC    hdc,
#  [in]  int    x,
#  [in]  int    y,
#  [out] LPSIZE lpsz
#);
gdi32.SetWindowExtEx.argtypes = (HANDLE, INT, INT, POINTER(SIZE))

########################################
# kernel32
########################################
#BOOL EnumResourceNamesW(
#  [in, optional] HMODULE          hModule,
#  [in]           LPCWSTR          lpType,
#  [in]           ENUMRESNAMEPROCW lpEnumFunc,
#  [in]           LONG_PTR         lParam
#);
kernel32.EnumResourceNamesW.argtypes = (HMODULE, LPCWSTR, ENUMRESNAMEPROCW, LONG_PTR)
kernel32.EnumResourceNamesW.restype = BOOL

kernel32.FindResourceW.argtypes = (HANDLE, LPCWSTR, LPCWSTR)
kernel32.FindResourceW.restype = HANDLE

kernel32.FreeLibrary.argtypes = (HANDLE, )

kernel32.GetModuleHandleW.argtypes = (LPCWSTR,)
kernel32.GetModuleHandleW.restype = HMODULE

kernel32.GetProcessId.argytypes = (HANDLE,)

#LPVOID GlobalLock(
#  [in] HGLOBAL hMem
#);
kernel32.GlobalLock.argtypes = (HANDLE, )
kernel32.GlobalLock.restype = LPVOID

kernel32.GlobalUnlock.argtypes = (HANDLE,)

kernel32.LoadLibraryW.restype = HANDLE

kernel32.LoadResource.argtypes = (HANDLE, HANDLE)
kernel32.LoadResource.restype = HANDLE

kernel32.LockResource.argtypes = (HANDLE, )
kernel32.LockResource.restype = HANDLE

kernel32.SizeofResource.argtypes = (HANDLE, HANDLE)

########################################
# ole32
########################################
#HRESULT CreateStreamOnHGlobal(
#  [in]  HGLOBAL  hGlobal,
#  [in]  BOOL     fDeleteOnRelease,
#  [out] LPSTREAM *ppstm
#);
#ole32.CreateStreamOnHGlobal.argtypes = (HANDLE, BOOL, LPSTREAM)

#ole32.RegisterDragDrop.argtypes = (HWND, POINTER(IDropTarget))  # POINTER(IDropTarget)

########################################
# shell32
########################################

shell32.DragAcceptFiles.argtypes = (HWND, BOOL)
shell32.DragQueryFileW.argtypes = (WPARAM, UINT, LPWSTR, UINT)
shell32.DragFinish.argtypes = (WPARAM, )

shell32.DragQueryPoint.argtypes = (WPARAM, LPPOINT)

##SHSTDAPI SHBindToObject(
##        IShellFolder       *psf,
##        PCUIDLIST_RELATIVE pidl,
##  [in]  IBindCtx           *pbc,
##        REFIID             riid,
##  [out] void               **ppv
##);
#shell32.SHBindToObject.argtypes = (c_void_p, PIDL, c_void_p, POINTER(GUID), POINTER(HANDLE))
#
##ULONG SHChangeNotifyRegister(
##  [in] HWND                      hwnd,
##       int                       fSources,
##       LONG                      fEvents,
##       UINT                      wMsg,
##       int                       cEntries,
##  [in] const SHChangeNotifyEntry *pshcne
##);
#shell32.SHChangeNotifyRegister.argtypes = (HWND, INT, LONG, UINT, INT, c_void_p)
#shell32.SHChangeNotifyRegister.restype = ULONG
#
#shell32.SHCreateItemFromParsingName.argtypes = (c_wchar_p, c_void_p, POINTER(GUID), POINTER(HANDLE))
#shell32.SHCreateItemFromParsingName.restype = HRESULT
#
#shell32.SHCreateItemFromIDList.argtypes = (PIDL, REFIID, LPVOID)
#shell32.SHCreateItemFromIDList.restype = HRESULT
#
##SHSTDAPI SHCreateShellItemArrayFromShellItem(
##  [in]  IShellItem *psi,
##  [in]  REFIID     riid,
##  [out] void       **ppv
##);
#shell32.SHCreateShellItemArrayFromShellItem.argtypes = (POINTER(IShellItem), REFIID, LPVOID)
#
#shell32.SHGetFileInfoW.argtypes = (c_void_p, DWORD, POINTER(SHFILEINFOW), UINT, UINT)
#shell32.SHGetFileInfoW.restype = HANDLE

shell32.Shell_NotifyIconW.argtypes = (DWORD, LPVOID)

##SHGetFolderLocation(
##  [in]  HWND             hwnd,
##  [in]  int              csidl,
##  [in]  HANDLE           hToken,
##  [in]  DWORD            dwFlags,
##  [out] PIDLIST_ABSOLUTE *ppidl
##);
#shell32.SHGetFolderLocation.argtypes = (HWND, INT, HANDLE, DWORD, POINTER(PIDL))
#
#shell32.SHGetImageList.argtypes = (INT, POINTER(GUID), POINTER(HANDLE))
#
##BOOL SHGetPathFromIDListW(
##  [in]  PCIDLIST_ABSOLUTE pidl,
##  [out] LPWSTR            pszPath
##);
#shell32.SHGetPathFromIDListW.argtypes = (PIDL, LPWSTR)
#
##HRESULT SHGetSpecialFolderLocation(
##  [in]  HWND             hwnd,
##  [in]  int              csidl,
##  [out] PIDLIST_ABSOLUTE *ppidl
##);
#shell32.SHGetSpecialFolderLocation.argtypes = (HANDLE, INT, POINTER(PIDL))

shell32.SHGetStockIconInfo.argtypes = (UINT, UINT, LPVOID)  # POINTER(SHSTOCKICONINFO)

##SHSTDAPI SHOpenFolderAndSelectItems(
##  [in]           PCIDLIST_ABSOLUTE     pidlFolder,
##                 UINT                  cidl,
##  [in, optional] PCUITEMID_CHILD_ARRAY apidl,
##                 DWORD                 dwFlags
##);
#shell32.SHOpenFolderAndSelectItems.argtypes = (c_void_p, UINT, c_void_p, DWORD)
#
##SHSTDAPI SHParseDisplayName(
##  [in]            PCWSTR           pszName,
##  [in, optional]  IBindCtx         *pbc,
##  [out]           PIDLIST_ABSOLUTE *ppidl,
##  [in]            SFGAOF           sfgaoIn,
##  [out, optional] SFGAOF           *psfgaoOut
##);
#shell32.SHParseDisplayName.argtypes = (LPCWSTR, c_void_p, POINTER(PIDL), ULONG, POINTER(ULONG))
#shell32.SHParseDisplayName.restype = HRESULT
#
#shell32.SHShellFolderView_Message.argtypes = (HWND, UINT, PIDL)  #LPARAM
#
##HINSTANCE ShellExecuteW(
##  [in, optional] HWND    hwnd,
##  [in, optional] LPCWSTR lpOperation,
##  [in]           LPCWSTR lpFile,
##  [in, optional] LPCWSTR lpParameters,
##  [in, optional] LPCWSTR lpDirectory,
##  [in]           INT     nShowCmd
##);
#
#shell32.ShellExecuteW.argtypes = (HANDLE, LPCWSTR, LPCWSTR, LPCWSTR, LPCWSTR, INT)
#shell32.ShellExecuteW.restype = HANDLE
#
##            RunFileDlg(
##                IN HWND    hwndOwner,       // Owner window, receives notifications
##                IN HICON   hIcon,           // Dialog icon handle, if NULL default icon is used
##                IN LPCTSTR lpszDirectory,   // Working directory
##                IN LPCTSTR lpszTitle,       // Dialog title, if NULL default is displayed
##                IN LPCTSTR lpszDescription, // Dialog description, if NULL default is displayed
##                IN UINT    uFlags           // Dialog flags (see below)
##            );
#
#shell32.RunFileDlg = shell32[61]
#shell32.RunFileDlg.argtypes = (HWND, HANDLE, LPWSTR, LPWSTR, LPWSTR, UINT)
#
#shell32.SHGetNameFromIDList.argtypes = (PIDL, UINT, LPVOID) # POINTER(WCHAR * MAX_PATH)

########################################
# user32
########################################
#HACCEL CreateAcceleratorTableW(
#  [in] LPACCEL paccel,
#  [in] int     cAccel
#);
#user32.CreateAcceleratorTableW.argtypes = (LPVOID, INT)
user32.CreateAcceleratorTableW.restype = HANDLE


#HWND CreateDialogIndirectParamW(
#  [in, optional] HINSTANCE       hInstance,
#  [in]           LPCDLGTEMPLATEW lpTemplate,
#  [in, optional] HWND            hWndParent,
#  [in, optional] DLGPROC         lpDialogFunc,
#  [in]           LPARAM          dwInitParam
#);
user32.CreateDialogIndirectParamW.argtypes = (HINSTANCE, c_void_p, HWND, WNDPROC, LPARAM)
user32.CreateDialogIndirectParamW.restype = HWND

user32.DestroyAcceleratorTable.restype = HANDLE
user32.TranslateAcceleratorW.argtypes = (HWND, HANDLE, POINTER(MSG))

user32.CopyImage.argtypes = [HANDLE, UINT, INT, INT, UINT]
user32.CopyImage.restype = HANDLE

#HICON CreateIconFromResourceEx(
#  [in] PBYTE presbits,
#  [in] DWORD dwResSize,
#  [in] BOOL  fIcon,
#  [in] DWORD dwVer,
#  [in] int   cxDesired,
#  [in] int   cyDesired,
#  [in] UINT  Flags
#);
user32.CreateIconFromResourceEx.argtypes = (c_void_p, DWORD, BOOL, DWORD, INT, INT, UINT)
user32.CreateIconFromResourceEx.restype = HICON

user32.CreateIconIndirect.argtypes = (HANDLE, )  #[POINTER(ICONINFO)]

#user32.EnableWindow.argytpes = [HWND, BOOL]

user32.DrawEdge.argtypes = (HDC, POINTER(RECT), UINT, UINT)

user32.DrawFocusRect.argtypes = (HANDLE, POINTER(RECT))

user32.DrawTextW.argtypes = (HANDLE, LPCWSTR , INT, POINTER(RECT), UINT)

#user32.DrawTextExW.argtypes = (HANDLE, LPWSTR, INT, POINTER(RECT), UINT, LPVOID)

user32.FillRect.argtypes = (HANDLE, POINTER(RECT), HBRUSH)

user32.FindWindowW.restype = HWND

user32.FindWindowExW.argtypes = (HWND, HWND, LPCWSTR, LPCWSTR)
user32.FindWindowExW.restype = HWND

user32.FrameRect.argtypes = (HDC, POINTER(RECT), HBRUSH)

user32.GetCaretPos.argtypes = (POINTER(POINT),)

user32.GetClassNameW.argtypes = (HWND, LPWSTR, INT)

user32.GetClipboardData.restype = HANDLE

user32.GetDC.argtypes = (HANDLE,)
user32.GetDC.restype = HANDLE

user32.GetDesktopWindow.restype = HANDLE
user32.GetForegroundWindow.restype = HANDLE

#BOOL GetMenuBarInfo(
#  [in]      HWND         hwnd,
#  [in]      LONG         idObject,
#  [in]      LONG         idItem,
#  [in, out] PMENUBARINFO pmbi
#);
user32.GetMenuBarInfo.argtypes = (HWND, LONG, LONG, LPVOID)

#BOOL GetMenuItemInfoW(
#  [in]      HMENU           hmenu,
#  [in]      UINT            item,
#  [in]      BOOL            fByPosition,
#  [in, out] LPMENUITEMINFOW lpmii
#);
user32.GetMenuItemInfoW.argtypes = (HANDLE, UINT, BOOL, LPVOID)

#int GetMenuStringW(
#  [in]            HMENU  hMenu,
#  [in]            UINT   uIDItem,
#  [out, optional] LPWSTR lpString,
#  [in]            int    cchMax,
#  [in]            UINT   flags
#);
user32.GetMenuStringW.argtypes = (HANDLE, UINT, LPWSTR, INT, UINT)

user32.GetWindow.argtypes = (HANDLE, UINT)

user32.GetWindowLongPtrA.argtypes = (HWND, LONG_PTR)
user32.GetWindowLongPtrA.restype = ULONG

user32.GetWindowLongPtrW.argtypes = (HWND, LONG_PTR)
user32.GetWindowLongPtrW.restype = WNDPROC

#user32.GetWindowLongW.restype = ULONG

user32.GetWindowTextW.argtypes = (HWND, LPWSTR, INT)  # LPWSTR

user32.GetWindowThreadProcessId.argtypes = (HANDLE, POINTER(DWORD))

user32.DefDlgProcW.argtypes = (HWND, c_uint, WPARAM, LPARAM)

user32.DefWindowProcW.argtypes = (HWND, c_uint, WPARAM, LPARAM)

#BOOL InvalidateRect(
#  [in] HWND       hWnd,
#  [in] const RECT *lpRect,
#  [in] BOOL       bErase
#);
user32.InvalidateRect.argtypes = (HWND, POINTER(RECT), BOOL)

user32.IsDialogMessageW.argtypes = (HWND, POINTER(MSG))
user32.IsDialogMessageW.restype = BOOL

user32.IsWindowVisible.argtypes = (HWND, )

user32.LoadIconW.argtypes = (HINSTANCE, LPCWSTR)
user32.LoadIconW.restype = HICON

user32.LoadImageW.argtypes = (HINSTANCE, LPCWSTR, UINT, INT, INT, UINT)
user32.LoadImageW.restype = HANDLE

user32.MapDialogRect.argtypes = (HWND, POINTER(RECT))

user32.MB_GetString.restype = LPCWSTR

user32.OffsetRect.argtypes = (POINTER(RECT), INT, INT)

user32.PostMessageW.argtypes = (HWND, UINT, LPVOID, LPVOID)
user32.PostMessageW.restype = LONG_PTR

user32.ReleaseDC.argtypes = (HWND, HANDLE)

user32.SetClassLongPtrW.argtypes = (HWND, INT, LONG_PTR)

user32.SetClipboardData.argtypes = (UINT, HANDLE)
user32.SetClipboardData.restype = HANDLE

#user32.SendMessageW.argtypes = (HWND, UINT, WPARAM, c_void_p)
# allow to send pointers
user32.SendMessageW.argtypes = (HWND, UINT, LPVOID, LPVOID)
user32.SendMessageW.restype = LONG_PTR

user32.SetSysColors.argtypes = (INT, POINTER(INT), POINTER(COLORREF))

user32.SetWindowLongPtrA.argtypes = (HWND, LONG_PTR, ULONG)
user32.SetWindowLongPtrA.restype = LONG

user32.SetWindowLongPtrW.argtypes = (HWND, LONG_PTR, WNDPROC)
user32.SetWindowLongPtrW.restype = WNDPROC

user32.SetWindowPos.argtypes = (HWND, LONG_PTR, INT, INT, INT, INT, UINT)

# https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwineventhook
# https://learn.microsoft.com/en-us/windows/win32/winauto/event-constants?redirectedfrom=MSDN
user32.SetWinEventHook.restype = HANDLE

user32.TrackPopupMenu.argtypes = (HANDLE, UINT, INT, INT, INT, HANDLE, c_void_p)
user32.TrackPopupMenuEx.argtypes = (HANDLE, UINT, INT, INT, HANDLE, c_void_p)

user32.RegisterShellHookWindow.argtypes = (HWND,)

########################################
# UxTheme
########################################

uxtheme.SetWindowTheme.argtypes = (HANDLE, LPCWSTR, LPCWSTR)

uxtheme.ShouldAppsUseDarkMode = uxtheme[136]

uxtheme.ShouldSystemUseDarkMode = uxtheme[138]

# using fnAllowDarkModeForWindow = bool (WINAPI*)(HWND hWnd, bool allow); // ordinal 133
uxtheme.AllowDarkModeForWindow = uxtheme[133]
uxtheme.AllowDarkModeForWindow.argtypes = (HWND, BOOL)

# SetPreferredAppMode = PreferredAppMode(WINAPI*)(PreferredAppMode appMode);
uxtheme.SetPreferredAppMode = uxtheme[135] # ordinal 135, in 1903

uxtheme.FlushMenuThemes = uxtheme[136]


#user32.LoadImageW.argtypes = [HANDLE, LPCWSTR, UINT, INT, INT, UINT]
#user32.LoadImageW.restype = HANDLE
#
#gdi32.CreatePatternBrush.argtypes = [HANDLE]
#
#user32.GetIconInfo.argtypes = [HANDLE, POINTER(ICONINFO)]
#
#gdi32.MaskBlt.argtypes = [HANDLE, INT, INT, INT, INT, HANDLE, INT, INT, HANDLE, INT, INT, DWORD]
#
#user32.GetDC.restype = HANDLE

# https://learn.microsoft.com/en-us/windows/win32/api/uxtheme/nf-uxtheme-openthemedata
#HTHEME OpenThemeData(
#  [in] HWND    hwnd,
#  [in] LPCWSTR pszClassList
#);
uxtheme.OpenThemeData.argtypes = (HWND, LPCWSTR)
uxtheme.OpenThemeData.restype = HANDLE

# https://learn.microsoft.com/en-us/windows/win32/api/uxtheme/nf-uxtheme-getthemepartsize
#HRESULT GetThemePartSize(
#  [in]  HTHEME    hTheme,
#  [in]  HDC       hdc,
#  [in]  int       iPartId,
#  [in]  int       iStateId,
#  [in]  LPCRECT   prc,
#        THEMESIZE eSize,
#  [out] SIZE      *psz
#);
uxtheme.GetThemePartSize.argtypes = (HANDLE, HANDLE, INT, INT, POINTER(RECT), UINT, POINTER(SIZE))

# https://learn.microsoft.com/en-us/windows/win32/api/uxtheme/nf-uxtheme-drawthemebackground
#HRESULT DrawThemeBackground(
#  [in] HTHEME  hTheme,
#  [in] HDC     hdc,
#  [in] int     iPartId,
#  [in] int     iStateId,
#  [in] LPCRECT pRect,
#  [in] LPCRECT pClipRect
#);
uxtheme.DrawThemeBackground.argtypes = (HANDLE, HANDLE, INT, INT, POINTER(RECT), POINTER(RECT))
