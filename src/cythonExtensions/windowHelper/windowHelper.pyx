# cython: language_level = 3str

"""This extension module provides functions for dealing with windows."""

import win32gui, win32con, winsound, pywintypes, ctypes
from time import sleep

def sendWindowsMessage(hwnd: int, message: int, wParam: int, lParam: str) -> int:
    """
    Description:
        Sends a message to the specified window.
    """
    
    encoded_path = lParam.encode('utf-16-le') + b'\0\0'  # Add null terminator
    result = ctypes.windll.user32.SendMessageW(hwnd, message, wParam, ctypes.c_void_p(ctypes.cast(encoded_path, ctypes.c_void_p).value))
    
    return ctypes.get_last_error() if not result else 0

def findHandleByClassName(className: str, check_all=False) -> list[int] | int:
    """
    Description:
        Searches for a window with the specified class name and returns its handle if found. Otherwise, returns `0`.
        - If multiple windows with the same class name exist:
            - If `check_all` is `True`, returns the handles of all the found windows.
            - If `check_all` is `False`, the handle of the top window in terms of z-index will be returned.
        - Unlike `win32ui.FindWindow`, this function does not raise an exception if no window is found.
    ---
    Parameters:
        `className -> str`:
            The class name of the window.
        
        `check_all -> bool`:
            If `True`, all windows with the specified class name will be returned. Otherwise, only the first window will be returned.
    ---
    Returns:
        `list[int]`: The handle to the window(s) with the specified class name if any exists, otherwise an empty list.
    """
    
    cdef int hwnd = win32gui.GetTopWindow(0)
    cdef list[int] output = []
    
    while hwnd:
        if win32gui.GetClassName(hwnd) == className:
            if check_all:
                output.append(hwnd)
            else:
                return hwnd
        
        hwnd = win32gui.GetWindow(hwnd, win32con.GW_HWNDNEXT)
    
    return output


def getHandleByTitle(title: str) -> int:
    """searches for a window with the specified title and returns its handle if found. Otherwise, returns `0`."""
    
    cdef int hwnd = win32gui.GetTopWindow(0)
    
    while hwnd:
        if win32gui.GetWindowText(hwnd) == title:
            return hwnd
        
        hwnd = win32gui.GetWindow(hwnd, win32con.GW_HWNDNEXT)
    
    return 0


def showMessageBox(msg: str, title="Warning", msgbox_type=1, int icon=win32con.MB_ICONERROR) -> int:
    """
    Description:
        Display an error message window with the specified message.
    ---
    Parameters:
        `msg -> str`:
            The error message to display.
        
        `Title -> str`:
            The title of the message box.
        
        `msgbox_type -> int`:
            The type of the message box (affects the number of buttons). Possible types are:
                
                `1`: One button: `Ok`.
                
                `2`: Two buttons: `Yes`, `No`.
                
                `3`: Two buttons: `OK`, `CANCEL`.
                
                `4`: Two buttons: `Retry`, `Cancel`.
                
                `5`: Three buttons: `Yes`, `No`, `Cancel`.
        
        `icon -> int`:
            The icon of the message box.
    ---
    Returns:
        `int`: The value associated with the click button. Below is a list of possible return values:
            Constant   | Value | Description
            -----------|-------|------------
            IDOK       |   1   | The OK button was selected.
            IDCANCEL   |   2   | The Cancel button was selected.
            DABORT     |   3   | The Abort button was selected.
            IDRETRY    |   4   | The Retry button was selected.
            IDIGNORE   |   5   | The Ignore button was selected.
            IDYES      |   6   | The Yes button was selected.
            IDNO       |   7   | The No button was selected.
            IDTRYAGAIN |   10  | The Try Again button was selected.
            IDCONTINUE |   11  | The Continue button was selected.
    """
    
    cdef int type = {1: win32con.MB_OK, 2: win32con.MB_YESNO, 3: win32con.MB_OKCANCEL,
                     4: win32con.MB_RETRYCANCEL, 5: win32con.MB_YESNOCANCEL}.get(msgbox_type)
    
    # https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-messageboxw
    return ctypes.windll.user32.MessageBoxW(None, msg, title, type | icon | win32con.MB_TOPMOST | win32con.MB_SETFOREGROUND)


def alwaysOnTop() -> None:
    """Toggles the alwaysOnTop feature for the active window."""
    
    cdef int hwnd = win32gui.GetForegroundWindow()
    
    # Check if the window is already topmost.
    cdef int is_topmost = (win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE) & win32con.WS_EX_TOPMOST) >> 3
    
    # Toggle the always on top property of the active window.
    win32gui.SetWindowPos(hwnd, (win32con.HWND_TOPMOST, win32con.HWND_NOTOPMOST)[is_topmost], 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
    
    if is_topmost:
        winsound.PlaySound(r"SFX\no-trespassing-368.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
    else:
        winsound.PlaySound(r"SFX\pedantic-490.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)


# Shake window - Doesn't work if the window is fullscreen
def shakeActiveWindow(cycles=5) -> None:
    """Simulates shake effect on the active window for the specified number of times."""
    
    # Get the handle of the window
    cdef int hwnd = win32gui.GetForegroundWindow()
    
    # Get the original position of the window
    cdef int x, y, width, height
    x, y, width, height = win32gui.GetWindowRect(hwnd)
    
    # Shake the window for a few seconds
    for i in range(cycles):
        # Move the window to a new position. You could also use `ctypes.windll.user32.SetWindowPos`,
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x + i, y + i, width, height, win32con.SWP_NOACTIVATE | win32con.SWP_NOSIZE)
        sleep(0.1)
        
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x - i, y - i, width, height, win32con.SWP_NOACTIVATE | win32con.SWP_NOSIZE)
        sleep(0.1)
    
    # Restore the original position of the window
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y, width, height, win32con.SWP_NOACTIVATE | win32con.SWP_NOSIZE)


def moveActiveWindow(hwnd=0, delta_x=0, delta_y=0, width=0, height=0) -> None:
    """
    Description:
        Moves the active or specified window by an (x, y) pixels, and change its size by (width, height) if passed.
    ---
    Parameters:
        `hwnd -> int`:
            A handle to a specified window. If not set, the active window will be selected.
        
        `x -> int`:
            The horizontal displacement.
        
        `y -> int`:
            The vertical displacement.
        
        `width -> int`:
            A value that will be added to the width of the specified window.
        
        `height -> int`:
            A value that will be added to the height of the specified window.
    """
    
    cdef int flags = (not (width or height)) * win32con.SWP_NOSIZE or win32con.SWP_NOMOVE # | win32con.SWP_FRAMECHANGED | win32con.SWP_DRAWFRAME
    
    # Get the handle of the window.
    if not hwnd:
        hwnd = win32gui.GetForegroundWindow()
    else:
        # Make sure the window is visible.
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    
    # Get the original position of the window.
    cdef int curr_x, curr_y, curr_width, curr_height
    curr_x, curr_y, curr_width, curr_height = win32gui.GetWindowRect(hwnd)
    
    # Change the position of the window.
    # win32gui.MoveWindow(hwnd, curr_x + x, curr_y + y, curr_width + width, curr_height + height, True)
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, curr_x + delta_x, curr_y + delta_y,
                          curr_width + width, curr_height + height, win32con.SWP_NOACTIVATE | flags)


def changeWindowOpacity(hwnd=0, opcode=1, increment=5) -> int:
    """
    Description:
        Increments (`opcode=any non-zero value`) or decrements (`opcode=0`) the opacity of the specified window by an (`increment`) value.
        
        If the specified window handle is `0` or `None`, the foreground window is selected.
    
    ---
    Return:
        `int`: The new opacity value. If -1 returned, then the operation has failed.
    """
    
    if not hwnd:
        hwnd = win32gui.GetForegroundWindow()
    
    # Get the extended window style of the specified window.
    # The specific extended style that controls whether a window is layered or not is WS_EX_LAYERED.
    # To change the opacity of a window, it is necessary to set the WS_EX_LAYERED style so that the window becomes a layered window.
    cdef int exstyle = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
    cdef int alpha
    
    # Check if the specified window is a layered window or not.
    if exstyle & win32con.WS_EX_LAYERED == win32con.WS_EX_LAYERED:
        # The `GetLayeredWindowAttributes` method can only retrieve the opacity value of a layered window. Passing a non-layered window raises an exception.
        try:
            alpha = win32gui.GetLayeredWindowAttributes(hwnd)[1] + (-increment if not opcode else increment)
        
        except pywintypes.error as e:
            print("Warning! This window does not support changing the opacity")
            return -1
    
    # The specified window is not a layered window. This means that it is opaque.
    else:
        alpha = 255 + (-increment if not opcode else increment)
    
    # Clipping the alpha value to a valid range from 0 to 255.
    alpha = max(min(alpha, 255), 25)
    
    # If the opacity is 1.0 or higher, then unset the WS_EX_LAYERED style from the window to make it opaque.
    if alpha == 255:
        exstyle &= ~win32con.WS_EX_LAYERED
    
    else:
        # Setting the WS_EX_LAYERED style on the window to make it transparent.
        exstyle |= win32con.WS_EX_LAYERED
    
    # Modifying the extended window style of the specified window.
    win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, exstyle)
    
    # Setting the window's alpha value to the new value. Note that the `SetLayeredWindowAttributes` method Can only be used on a layered window.
    if exstyle & win32con.WS_EX_LAYERED == win32con.WS_EX_LAYERED:
        win32gui.SetLayeredWindowAttributes(hwnd, 0, alpha, win32con.LWA_ALPHA)
    
    return alpha
