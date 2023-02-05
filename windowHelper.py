"""This module provides functions for dealing with windows."""

import win32gui, win32con, winsound
from time import sleep

def GetHandleByClassName(className: str, check_all=False) -> list[int] | int | None:
    """
    Description:
        - Returns the handle to the window with the specified class name if one exists.
        - If multiple windows with the same class name exist:
            - If `check_all` is `True`, returns the handles of all the found windows.
            - If `check_all` is `False`, the handle of the top window in terms of z-index will be returned.
        - Unlike `win32ui.FindWindow`, this function does not raise an exception if no window is found.
    """
    
    hwnd = win32gui.GetTopWindow(0)
    output = []
    while hwnd:
        if win32gui.GetClassName(hwnd) == className:
            if check_all:
                output.append(hwnd)
            else:
                return hwnd
        
        hwnd = win32gui.GetWindow(hwnd, win32con.GW_HWNDNEXT)
    
    return output or None

def GetHandleByTitle(title: str) -> int | None:
    """Returns the handle to the window with the specified title if one exists."""
    
    hwnd = win32gui.GetTopWindow(0)
    while hwnd:
        if win32gui.GetWindowText(hwnd) == title:
            return hwnd
        
        hwnd = win32gui.GetWindow(hwnd, win32con.GW_HWNDNEXT)
    
    return None

def AlwaysOnTop():
    """Toggles `AlwaysOnTop` on or off for the active window."""
    
    hwnd = win32gui.GetForegroundWindow()
    
    # Check if the window is already topmost.
    is_topmost = (win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE) & win32con.WS_EX_TOPMOST) >> 3
    
    # Toggle the always on top property of the active window.
    win32gui.SetWindowPos(hwnd, (win32con.HWND_TOPMOST, win32con.HWND_NOTOPMOST)[is_topmost], 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
    
    if is_topmost:
        winsound.PlaySound(r"SFX\no-trespassing-368.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
    else:
        winsound.PlaySound(r"SFX\pedantic-490.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)

# Shake window - Doesn't work if the window is fullscreen
def ShakeActiveWindow(cycles=5):
    """Simulates shake effect on the active window for the specified number of times."""
    
    # Get the handle of the window
    hwnd = win32gui.GetForegroundWindow()
    
    # Get the original position of the window
    x, y, width, height = win32gui.GetWindowRect(hwnd)
    
    # Shake the window for a few seconds
    for i in range(cycles):
        # Move the window to a new position. The `win32gui.SetWindowPos`. You could also use `ctypes.windll.user32.SetWindowPos`,
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x + i, y + i, width, height, win32con.SWP_NOACTIVATE | win32con.SWP_NOSIZE)
        sleep(0.1)
        
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x - i, y - i, width, height, win32con.SWP_NOACTIVATE | win32con.SWP_NOSIZE)
        sleep(0.1)
    
    # Restore the original position of the window
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y, width, height, win32con.SWP_NOACTIVATE | win32con.SWP_NOSIZE)

def MoveActiveWindow(hwnd=None, delta_x=0, delta_y=0, width=0, height=0):
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
    
    flags = (not width and not height) * win32con.SWP_NOSIZE or win32con.SWP_NOMOVE # | win32con.SWP_FRAMECHANGED | win32con.SWP_DRAWFRAME
    
    # Get the handle of the window.
    if not hwnd:
        hwnd = win32gui.GetForegroundWindow()
    else:
        # Make sure the window is visible.
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    
    # Get the original position of the window.
    curr_x, curr_y, curr_width, curr_height = win32gui.GetWindowRect(hwnd)
    
    # Change the position of the window.
    # win32gui.MoveWindow(hwnd, curr_x + x, curr_y + y, curr_width + width, curr_height + height, True)
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, curr_x + delta_x, curr_y + delta_y,
                          curr_width + width, curr_height + height, win32con.SWP_NOACTIVATE | flags)
