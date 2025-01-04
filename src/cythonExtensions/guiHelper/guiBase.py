import ctypes
from ctypes import wintypes
import win32con

# Utility functions
def errcheck(result, func, args):
    """Properly simple error checking for Windows API calls"""
    
    if result is None or result == 0:
        raise ctypes.WinError(ctypes.get_last_error())
    
    return result

def MAKEINTRESOURCEW(x):
    """Convert an integer to a resource string"""
    
    return wintypes.LPCWSTR(x)

def LOWORD(l):
    return l & 0xffff

def HIWORD(l):
    return (l >> 16) & 0xffff


# Types missing from ctypes.wintypes
LRESULT = ctypes.c_int64
HCURSOR = ctypes.c_void_p
WNDPROC = ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)

class WNDCLASSW(ctypes.Structure):
    """
    ## Description:
        Structure for registering a window class.
    
    ## Fields:
        - `style`: The class style(s). Defines the window class attributes.
        - `lpfnWndProc`: Pointer to the window procedure. Defines how messages are processed.
        - `cbClsExtra`: The number of extra bytes to allocate following the window-class structure.
        - `cbWndExtra`: The number of extra bytes to allocate following the window instance.
        - `hInstance`: Handle to the instance that contains the window procedure.
        - `hIcon`: Handle to the class icon.
        - `hCursor`: Handle to the class cursor.
        - `hbrBackground`: Handle to the class background brush.
        - `lpszMenuName`: Pointer to a null-terminated string or resource identifier for the menu.
        - `lpszClassName`: Pointer to a null-terminated string that specifies the class name.
    """
    
    style: wintypes.UINT
    """The class style(s). Defines the window class attributes."""
    
    lpfnWndProc: ctypes.c_void_p
    """Pointer to the window procedure. Defines how messages are processed."""
    
    cbClsExtra: ctypes.c_int
    """The number of extra bytes to allocate following the window-class structure."""
    
    cbWndExtra: ctypes.c_int
    """The number of extra bytes to allocate following the window instance."""
    
    hInstance: wintypes.HINSTANCE
    """Handle to the instance that contains the window procedure."""
    
    hIcon: wintypes.HICON
    """Handle to the class icon."""
    
    hCursor: HCURSOR
    """Handle to the class cursor."""
    
    hbrBackground: wintypes.HBRUSH
    """Handle to the class background brush."""
    
    lpszMenuName: wintypes.LPCWSTR
    """Pointer to a null-terminated string or resource identifier for the menu."""
    
    lpszClassName: wintypes.LPCWSTR
    """Pointer to a null-terminated string that specifies the class name."""
    
    _fields_ = [
        ('style', wintypes.UINT),
        ('lpfnWndProc', WNDPROC),
        ('cbClsExtra', ctypes.c_int),
        ('cbWndExtra', ctypes.c_int),
        ('hInstance', wintypes.HINSTANCE),
        ('hIcon', wintypes.HICON),
        ('hCursor', HCURSOR),
        ('hbrBackground', wintypes.HBRUSH),
        ('lpszMenuName', wintypes.LPCWSTR),
        ('lpszClassName', wintypes.LPCWSTR)
    ]

class POINT(ctypes.Structure):
    """
    Description:
        Structure for defining a point.
    
    Fields:
        - `x`: The x-coordinate of the point.
        - `y`: The y-coordinate of the point.
    """
    
    x: ctypes.c_long
    """The x-coordinate of the point."""
    
    y: ctypes.c_long
    """The y-coordinate of the point."""
    
    _fields_ = [
        ("x", ctypes.c_long),
        ("y", ctypes.c_long)
    ]

class MINMAXINFO(ctypes.Structure):
    """
    Description:
        Structure for handling window resizing restrictions.
    
    Fields:
        - `ptReserved`: Reserved, must be zero.
        - `ptMaxSize`: The maximized width (x) and height (y) of the window.
        - `ptMaxPosition`: The position of the left (x) and top (y) of the window when maximized.
        - `ptMinTrackSize`: The minimum tracking width (x) and height (y) of the window.
        - `ptMaxTrackSize`: The maximum tracking width (x) and height (y) of the window.
    """
    
    ptReserved: POINT
    """Reserved, must be zero."""
    
    ptMaxSize: POINT
    """The maximized width (x) and height (y) of the window."""
    
    ptMaxPosition: POINT
    """The position of the left (x) and top (y) of the window when maximized."""
    
    ptMinTrackSize: POINT
    """The minimum tracking width (x) and height (y) of the window."""
    
    ptMaxTrackSize: POINT
    """The maximum tracking width (x) and height (y) of the window."""
    
    _fields_ = [
        ("ptReserved", POINT),
        ("ptMaxSize", POINT),
        ("ptMaxPosition", POINT),
        ("ptMinTrackSize", POINT),
        ("ptMaxTrackSize", POINT)
    ]

class PAINTSTRUCT(ctypes.Structure):
    """
    Description:
        Structure for handling painting messages.
    
    Fields:
        - `hdc`: Handle to the device context for the client area.
        - `fErase`: Indicates whether the background must be erased.
        - `rcPaint`: The rectangle in which the painting is requested.
        - `fRestore`: Reserved; used internally by the system.
        - `fIncUpdate`: Reserved; used internally by the system.
        - `rgbReserved`: Reserved; used internally by the system.
    """
    
    hdc: wintypes.HDC
    """Handle to the device context for the client area."""
    
    fErase: wintypes.BOOL
    """Indicates whether the background must be erased."""
    
    rcPaint: wintypes.RECT
    """The rectangle in which the painting is requested."""
    
    fRestore: wintypes.BOOL
    """Reserved; used internally by the system."""
    
    fIncUpdate: wintypes.BOOL
    """Reserved; used internally by the system."""
    
    rgbReserved: ctypes.c_byte
    """Reserved; used internally by the system."""
    
    _fields_ = [('hdc', wintypes.HDC),
                ('fErase', wintypes.BOOL),
                ('rcPaint', wintypes.RECT),
                ('fRestore', wintypes.BOOL),
                ('fIncUpdate', wintypes.BOOL),
                ('rgbReserved', wintypes.BYTE * 32)]

class RGBQUAD(ctypes.Structure):
    """
    Description:
        Structure for defining a color in a bitmap.
    
    Fields:
        - `rgbBlue`: The intensity of blue in the color.
        - `rgbGreen`: The intensity of green in the color.
        - `rgbRed`: The intensity of red in the color.
        - `rgbReserved`: Reserved; must be zero.
    """
    
    rgbBlue: ctypes.c_byte
    """The intensity of blue in the color."""
    
    rgbGreen: ctypes.c_byte
    """The intensity of green in the color."""
    
    rgbRed: ctypes.c_byte
    """The intensity of red in the color."""
    
    rgbReserved: ctypes.c_byte
    """Reserved; must be zero."""
    
    _fields_ = [
        ('rgbBlue', ctypes.c_byte),
        ('rgbGreen', ctypes.c_byte),
        ('rgbRed', ctypes.c_byte),
        ('rgbReserved', ctypes.c_byte)
    ]

class BITMAPINFOHEADER(ctypes.Structure):
    """
    Description:
        Structure for defining the dimensions and color format of a DIB.
    
    Fields:
        - `biSize`: The size of the structure.
        - `biWidth`: The width of the bitmap, in pixels.
        - `biHeight`: The height of the bitmap, in pixels.
        - `biPlanes`: The number of color planes for the target device.
        - `biBitCount`: The number of bits per pixel.
        - `biCompression`: The type of compression for the bitmap.
        - `biSizeImage`: The size of the image data, in bytes.
        - `biXPelsPerMeter`: The horizontal resolution, in pixels per meter.
        - `biYPelsPerMeter`: The vertical resolution, in pixels per meter.
        - `biClrUsed`: The number of colors used in the bitmap.
        - `biClrImportant`: The number of important colors in the bitmap.
    """
    
    biSize : wintypes.DWORD
    """The size of the structure."""
    
    biWidth : wintypes.LONG
    """The width of the bitmap, in pixels."""
    
    biHeight : wintypes.LONG
    """The height of the bitmap, in pixels."""
    
    biPlanes : wintypes.WORD
    """The number of color planes for the target device."""
    
    biBitCount : wintypes.WORD
    """The number of bits per pixel."""
    
    biCompression : wintypes.DWORD
    """The type of compression for the bitmap."""
    
    biSizeImage : wintypes.DWORD
    """The size of the image data, in bytes."""
    
    biXPelsPerMeter : wintypes.LONG
    """The horizontal resolution, in pixels per meter."""
    
    biYPelsPerMeter : wintypes.LONG
    """The vertical resolution, in pixels per meter."""
    
    biClrUsed : wintypes.DWORD
    """The number of colors used in the bitmap."""
    
    biClrImportant : wintypes.DWORD
    """The number of important colors in the bitmap."""
    
    
    _fields_ = [
        ('biSize', wintypes.DWORD),
        ('biWidth', wintypes.LONG),
        ('biHeight', wintypes.LONG),
        ('biPlanes', wintypes.WORD),
        ('biBitCount', wintypes.WORD),
        ('biCompression', wintypes.DWORD),
        ('biSizeImage', wintypes.DWORD),
        ('biXPelsPerMeter', wintypes.LONG),
        ('biYPelsPerMeter', wintypes.LONG),
        ('biClrUsed', wintypes.DWORD),
        ('biClrImportant', wintypes.DWORD)
    ]

class BITMAPINFO(ctypes.Structure):
    """
    Description:
        Structure for defining a DIB.
    
    Fields:
        - `bmiHeader`: The dimensions and color format of the DIB.
        - `bmiColors`: The colors in the DIB.
    """
    
    bmiHeader: BITMAPINFOHEADER
    """The dimensions and color format of the DIB."""
    
    bmiColors: RGBQUAD
    """The colors in the DIB."""
    
    _fields_ = [
        ('bmiHeader', BITMAPINFOHEADER),
        ('bmiColors', RGBQUAD * 1)
    ]

class LVCOLUMNW(ctypes.Structure):
    """
    Description:
        Structure for defining a column in a list-view control.
    
    Fields:
        - `mask`: The column attributes to retrieve or set.
        - `fmt`: The column format.
        - `cx`: The width of the column.
        - `pszText`: The column header text.
        - `cchTextMax`: The size of the buffer pointed to by `pszText`.
        - `iSubItem`: The column index.
        - `iImage`: The index of the column's image.
        - `iOrder`: The display order of the column.
    """
    
    mask: wintypes.UINT
    """The column attributes to retrieve or set."""
    
    fmt: wintypes.INT
    """The column format."""
    
    cx: wintypes.INT
    """The width of the column."""
    
    pszText: wintypes.LPWSTR
    """The column header text."""
    
    cchTextMax: wintypes.INT
    """The size of the buffer pointed to by `pszText`."""
    
    iSubItem: wintypes.INT
    """The column index."""
    
    iImage: wintypes.INT
    """The index of the column's image."""
    
    iOrder: wintypes.INT
    """The display order of the column."""
    
    _fields_ = [
        ('mask', wintypes.UINT),
        ('fmt', wintypes.INT),
        ('cx', wintypes.INT),
        ('pszText', wintypes.LPWSTR),
        ('cchTextMax', wintypes.INT),
        ('iSubItem', wintypes.INT),
        ('iImage', wintypes.INT),
        ('iOrder', wintypes.INT)
    ]


class LVITEMW(ctypes.Structure):
    """
    Description:
        Structure for defining an item in a list-view control.
    
    Fields:
        - `mask`: The item attributes to retrieve or set.
        - `iItem`: The index of the item.
        - `iSubItem`: The index of the subitem.
        - `state`: The state of the item.
        - `stateMask`: The state mask.
        - `pszText`: The item text.
        - `cchTextMax`: The size of the buffer pointed to by `pszText`.
        - `iImage`: The index of the item's image.
        - `lParam`: The item data.
        - `iIndent`: The item indentation.
        - `iGroupId`: The group ID.
        - `cColumns`: The number of columns.
        - `puColumns`: The column indices.
        - `piColFmt`: The column formats.
        - `iGroup`: The group index.
    """
    
    mask: wintypes.UINT
    """The item attributes to retrieve or set."""
    
    iItem: wintypes.INT
    """The index of the item."""
    
    iSubItem: wintypes.INT
    """The index of the subitem."""
    
    state: wintypes.UINT
    """The state of the item."""
    
    stateMask: wintypes.UINT
    """The state mask."""
    
    pszText: wintypes.LPWSTR
    """The item text."""
    
    cchTextMax: wintypes.INT
    """The size of the buffer pointed to by `pszText`."""
    
    iImage: wintypes.INT
    """The index of the item's image."""
    
    lParam: wintypes.LPARAM
    """The item data."""
    
    iIndent: wintypes.INT
    """The item indentation."""
    
    iGroupId: wintypes.INT
    """The group ID."""
    
    cColumns: wintypes.UINT
    """The number of columns."""
    
    # puColumns: ctypes.POINTER(wintypes.INT)
    # """The column indices."""
    
    # piColFmt: ctypes.POINTER(wintypes.INT)
    # """The column formats."""
    
    iGroup: wintypes.INT
    """The group index."""
    
    _fields_ = [
        ('mask', wintypes.UINT),
        ('iItem', wintypes.INT),
        ('iSubItem', wintypes.INT),
        ('state', wintypes.UINT),
        ('stateMask', wintypes.UINT),
        ('pszText', wintypes.LPWSTR),
        ('cchTextMax', wintypes.INT),
        ('iImage', wintypes.INT),
        ('lParam', wintypes.LPARAM),
        ('iIndent', wintypes.INT),
        ('iGroupId', wintypes.INT),
        ('cColumns', wintypes.UINT),
        ('puColumns', ctypes.POINTER(wintypes.INT)),
        ('piColFmt', ctypes.POINTER(wintypes.INT)),
        ('iGroup', wintypes.INT)
    ]

# Loading libraries and defining function prototypes
kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
kernel32.GetModuleHandleW.argtypes = wintypes.LPCWSTR,
kernel32.GetModuleHandleW.restype = wintypes.HMODULE
kernel32.GetModuleHandleW.errcheck = errcheck

user32 = ctypes.WinDLL('user32', use_last_error=True)
user32.CreateWindowExW.argtypes = wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, wintypes.HWND, wintypes.HMENU, wintypes.HINSTANCE, wintypes.LPVOID
user32.CreateWindowExW.restype = wintypes.HWND
user32.CreateWindowExW.errcheck = errcheck
user32.LoadIconW.argtypes = wintypes.HINSTANCE, wintypes.LPCWSTR
user32.LoadIconW.restype = wintypes.HICON
user32.LoadIconW.errcheck = errcheck
user32.LoadCursorW.argtypes = wintypes.HINSTANCE, wintypes.LPCWSTR
user32.LoadCursorW.restype = HCURSOR
user32.LoadCursorW.errcheck = errcheck
user32.RegisterClassW.argtypes = ctypes.POINTER(WNDCLASSW),
user32.RegisterClassW.restype = wintypes.ATOM
user32.RegisterClassW.errcheck = errcheck
user32.ShowWindow.argtypes = wintypes.HWND, ctypes.c_int
user32.ShowWindow.restype = wintypes.BOOL
user32.UpdateWindow.argtypes = wintypes.HWND,
user32.UpdateWindow.restype = wintypes.BOOL
user32.UpdateWindow.errcheck = errcheck
user32.GetMessageW.argtypes = ctypes.POINTER(wintypes.MSG), wintypes.HWND, wintypes.UINT, wintypes.UINT
user32.GetMessageW.restype = wintypes.BOOL
user32.TranslateMessage.argtypes = ctypes.POINTER(wintypes.MSG),
user32.TranslateMessage.restype = wintypes.BOOL
user32.DispatchMessageW.argtypes = ctypes.POINTER(wintypes.MSG),
user32.DispatchMessageW.restype = LRESULT
user32.DrawTextW.argtypes = wintypes.HDC, wintypes.LPCWSTR, ctypes.c_int, ctypes.POINTER(wintypes.RECT), wintypes.UINT
user32.DrawTextW.restype = ctypes.c_int
user32.PostQuitMessage.argtypes = ctypes.c_int,
user32.PostQuitMessage.restype = None
user32.DefWindowProcW.argtypes = wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM
user32.DefWindowProcW.restype = LRESULT
user32.SetPropW.argtypes = wintypes.HWND, wintypes.LPCWSTR, wintypes.HANDLE
user32.SetPropW.restype = wintypes.BOOL
user32.GetPropW.argtypes = wintypes.HWND, wintypes.LPCWSTR
user32.GetPropW.restype = wintypes.HANDLE

gdi32 = ctypes.WinDLL('gdi32', use_last_error=True)
gdi32.GetStockObject.argtypes = ctypes.c_int,
gdi32.GetStockObject.restype = wintypes.HGDIOBJ

# For subclassing (required to handle the tab key press in the edit control).
comctl32 = ctypes.WinDLL('ComCtl32.dll')
comctl32.SetWindowSubclass.argtypes = [wintypes.HWND, ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM, ctypes.c_ulonglong, ctypes.c_ulonglong), ctypes.c_ulonglong, ctypes.c_ulonglong]
comctl32.SetWindowSubclass.restype = wintypes.BOOL
comctl32.DefSubclassProc.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
comctl32.DefSubclassProc.restype = LRESULT
comctl32.RemoveWindowSubclass.argtypes = [wintypes.HWND, ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM, ctypes.c_ulonglong, ctypes.c_ulonglong), ctypes.c_ulonglong]
comctl32.RemoveWindowSubclass.restype = wintypes.BOOL

# Windows API Constants
IDI_APPLICATION = MAKEINTRESOURCEW(win32con.IDI_APPLICATION)
IDC_ARROW = MAKEINTRESOURCEW(win32con.IDC_ARROW)

# List view control constants: https://gist.github.com/AHK-just-me/6098109
LVS_REPORT = 0x0001
LVCF_TEXT = 0x0004
LVCF_WIDTH = 0x0002
LVCF_SUBITEM = 0x0008
LVIF_TEXT = 0x0001
LVM_INSERTCOLUMNW = 0x1061 # (LVM_FIRST + 97)
LVM_INSERTITEMW = 0x104D # (LVM_FIRST + 77)
LVM_SETITEMW = 0x104C # (LVM_FIRST + 76)

# Image list control constants: https://learn.microsoft.com/en-us/windows/win32/controls/ilc-constants

# Subclassing function
@ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM, ctypes.c_ulonglong, ctypes.c_ulonglong)
def SubclsProc(hwnd, uMsg, wParam, lParam, uIdSubclass, dwRefData):
    """Procedure used with subclassed window controls to handle key presses not handled by the default window procedure"""
    
    if uMsg == win32con.WM_KEYDOWN:
        # Move focus to the next or previous control in the tab order based on Shift key state.
        if wParam == win32con.VK_TAB:
            # Check if Shift key is pressed
            shift_pressed = (user32.GetAsyncKeyState(win32con.VK_SHIFT) & 0x8000) != 0
            user32.SetFocus(user32.GetNextDlgTabItem(dwRefData, hwnd, shift_pressed))
            
            return 0
        
        # Simulate clicking the OK button when Enter key is pressed.
        elif wParam == win32con.VK_RETURN:
            parent_hwnd = dwRefData
            button_hwnd = user32.GetDlgItem(parent_hwnd, 1)  # The OK button must have an ID of 1
            user32.SendMessageW(button_hwnd, win32con.BM_CLICK, 0, 0)
            
            return 0
        
        elif wParam == win32con.VK_MENU:
            print("Menu key pressed")
            return 0
    
    # if uMsg == win32con.WM_PAINT:
    #     # Handle custom drawing to center text vertically
    #     ps = PAINTSTRUCT()
    #     hdc = user32.BeginPaint(hwnd, ctypes.byref(ps))
        
    #     # Get client rect and text
    #     rect = wintypes.RECT()
    #     user32.GetClientRect(hwnd, ctypes.byref(rect))
        
    #     buffer = ctypes.create_unicode_buffer(256)
    #     user32.GetWindowTextW(hwnd, buffer, len(buffer))
        
    #     # Clear the background
    #     user32.FillRect(hdc, ctypes.byref(rect), gdi32.GetStockObject(0))
        
    #     # Set text alignment
    #     gdi32.SetBkMode(hdc, win32con.TRANSPARENT)
    #     gdi32.SelectObject(hdc, gdi32.GetStockObject(17))
        
    #     # Calculate text size
    #     textSize = wintypes.SIZE()
    #     gdi32.GetTextExtentPoint32W(hdc, buffer.value, len(buffer.value), ctypes.byref(textSize))
        
    #     # Calculate position to center text
    #     rect.top += (rect.bottom - rect.top - textSize.cy) // 2
    #     rect.bottom = rect.top + textSize.cy
        
    #     # Draw the text
    #     user32.DrawTextW(hdc, buffer.value, -1, ctypes.byref(rect), win32con.DT_CENTER | win32con.DT_SINGLELINE)
    #     user32.EndPaint(hwnd, ctypes.byref(ps))
        
        # return 0
    
    return comctl32.DefSubclassProc(hwnd, uMsg, wParam, lParam)

import os, psutil, win32gui, win32ui, win32api, pywintypes, PIL.Image
from typing import Union, Optional, Tuple


class ProcessViewer():
    def __init__(self, title: str):
        self.classname = "ProcessViewer"
        self.title = title
        self.width = 800
        self.height = 600
        self.this_pid = os.getpid()
        self.iconDir = os.path.join(os.path.dirname(__file__), 'icons')
        self.defaultIconPath = os.path.join(self.iconDir, 'default.png')
        self.processes = []

    def WndProc(self, hwnd, message, wParam, lParam):
        if message == win32con.WM_DESTROY:
            self.destroyWindow()
            return 0
        elif message == win32con.WM_CREATE:
            self.initListView(hwnd)
            return 0
        return user32.DefWindowProcW(hwnd, message, wParam, lParam)

    def initListView(self, hwnd):
        print("Initializing list view")
        hWndListView = user32.CreateWindowExW(
            0,
            'SysListView32',
            None,
            win32con.WS_CHILD | win32con.WS_VISIBLE | LVS_REPORT,
            0, 0, self.width, self.height,
            hwnd,
            None,
            kernel32.GetModuleHandleW(None),
            None
        )

        # Insert columns
        columns = ['Process Name', 'PID', 'CPU Usage', 'Memory Usage']
        for i, col in enumerate(columns):
            lvColumn = LVCOLUMNW()
            lvColumn.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM
            lvColumn.pszText = col
            lvColumn.cx = self.width // len(columns)
            user32.SendMessageW(hWndListView, LVM_INSERTCOLUMNW, i, ctypes.byref(lvColumn))

        # Fetch and display process information
        self.fetch_processes()
        self.display_processes(hWndListView)

    def fetch_processes(self):
        self.processes = self.get_running_processes_with_icons()

    def display_processes(self, hWndListView):
        for i, proc in enumerate(self.processes):
            items = [proc['name'], str(proc['pid']), f"{proc['cpu_percent']}%", f"{proc['memory_percent']:.2f}%"]
            for j, item in enumerate(items):
                lvItem = LVITEMW()
                lvItem.mask = LVIF_TEXT
                lvItem.iItem = i
                lvItem.iSubItem = j
                lvItem.pszText = item
                if j == 0:  # Insert the first column which represents a row
                    user32.SendMessageW(hWndListView, LVM_INSERTITEMW, 0, ctypes.byref(lvItem))
                else:  # Insert the rest columns of the row
                    user32.SendMessageW(hWndListView, LVM_SETITEMW, 0, ctypes.byref(lvItem))

        # Update the window to show the list view
        user32.UpdateWindow(hWndListView)

    # def get_version_number(self, exe: str) -> Optional[Tuple[int, int, int, int]]:
    #     try:
    #         info = win32api.GetFileVersionInfo(exe, "\\")
    #         ms = info['FileVersionMS']
    #         ls = info['FileVersionLS']
    #         return win32api.HIWORD(ms), win32api.LOWORD(ms), win32api.HIWORD(ls), win32api.LOWORD(ls)
    #     except pywintypes.error:
    #         return None

    # def github_version_tag(self, xyz: Tuple[int, int, int, int]):
    #     versioned = []
    #     zero_removed = False
    #     for component in xyz[::-1]:
    #         if component >= 1 or zero_removed:
    #             zero_removed = True
    #             versioned.insert(0, str(component))
    #     return 'v' + '.'.join(versioned)

    # def proc_version_tag(self, proc: psutil.Process):
    #     exe = proc.info.get("exe", None)
    #     if exe is None:
    #         try:
    #             exe = proc.exe()
    #         except (psutil.NoSuchProcess, psutil.AccessDenied):
    #             return None
    #     version = self.get_version_number(exe)
    #     if version is not None:
    #         return self.github_version_tag(version)
    #     return None

    def icon_path(self, exe: str, name: str):
        id_file_name = f'{name}.png'
        id_path = os.path.join(self.iconDir, id_file_name)

        if not os.path.exists(id_path):
            ico_x = win32api.GetSystemMetrics(win32con.SM_CXICON)
            try:
                large, small = win32gui.ExtractIconEx(exe, 0)
            except pywintypes.error:
                return self.defaultIconPath

            if not large:
                return self.defaultIconPath

            if small:
                win32gui.DestroyIcon(small[0])

            hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
            h_bmp = win32ui.CreateBitmap()
            h_bmp.CreateCompatibleBitmap(hdc, ico_x, ico_x)
            hdc = hdc.CreateCompatibleDC()

            hdc.SelectObject(h_bmp)
            hdc.DrawIcon((0, 0), large[0])

            bmp_str = h_bmp.GetBitmapBits(True)
            img = PIL.Image.frombuffer(
                'RGBA',
                (32, 32),
                bmp_str, 'raw', 'BGRA', 0, 1
            )

            img.save(id_path)
            win32gui.DestroyIcon(large[0])

        return id_path

    def get_running_processes_with_icons(self):
        processes = []
        for proc in psutil.process_iter(['pid', 'name', 'exe', 'cpu_percent', 'memory_percent']):
            try:
                proc_info = proc.info
                proc_info['icon'] = self.icon_path(proc_info['exe'], proc_info['name'])
                processes.append(proc_info)
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        return processes

    def run(self):
        wndclass = WNDCLASSW()
        wndclass.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
        wndclass.lpfnWndProc = WNDPROC(self.WndProc)
        wndclass.cbClsExtra = wndclass.cbWndExtra = 0
        wndclass.hInstance = kernel32.GetModuleHandleW(None)
        wndclass.hIcon = user32.LoadIconW(None, IDI_APPLICATION)
        wndclass.hCursor = user32.LoadCursorW(None, IDC_ARROW)
        wndclass.hbrBackground = gdi32.GetStockObject(win32con.WHITE_BRUSH)
        wndclass.lpszMenuName = None
        wndclass.lpszClassName = self.classname

        try:
            user32.RegisterClassW(ctypes.byref(wndclass))
        except OSError:
            print("Failed to register window class as it already exists. Continuing...")
        
        rect = wintypes.RECT()
        user32.AdjustWindowRectEx(ctypes.byref(rect), win32con.WS_OVERLAPPEDWINDOW, False, 0)
        border_width = rect.right - rect.left
        border_height = rect.bottom - rect.top

        self.hwnd = user32.CreateWindowExW(
            0,
            wndclass.lpszClassName,
            self.title,
            win32con.WS_OVERLAPPEDWINDOW,
            win32con.CW_USEDEFAULT,
            win32con.CW_USEDEFAULT,
            self.width + border_width,
            self.height + border_height,
            None,
            None,
            wndclass.hInstance,
            None
        )
        
        user32.ShowWindow(self.hwnd, win32con.SW_SHOWNORMAL)
        user32.UpdateWindow(self.hwnd)

        msg = wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))

    def destroyWindow(self) -> None:
        user32.DestroyWindow(self.hwnd)
        user32.UnregisterClassW(self.classname, 0)
        user32.PostQuitMessage(0)

if __name__ == '__main__':
    app = ProcessViewer("Process Viewer")
    app.run()
