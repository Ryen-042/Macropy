"""This module provides functions for dealing with windows."""

import win32con


def FindHandleByClassName(className: str, check_all=False) -> list[int]:
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
    ...


def GetHandleByTitle(title: str) -> int:
    """searches for a window with the specified title and returns its handle if found. Otherwise, returns `0`."""
    ...


def ShowMessageBox(msg: str, title="Warning", msgbox_type=1, icon=win32con.MB_ICONERROR) -> int:
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
        `int`: The value associated with the click button.
    """
    ...


def AlwaysOnTop() -> None:
    """Toggles `AlwaysOnTop` on or off for the active window."""
    ...


# Shake window - Doesn't work if the window is fullscreen
def ShakeActiveWindow(cycles=5) -> None:
    """Simulates shake effect on the active window for the specified number of times."""
    ...


def MoveActiveWindow(hwnd=0, delta_x=0, delta_y=0, width=0, height=0) -> None:
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
    ...


def ChangeWindowOpacity(hwnd=0, opcode=1, increment=5) -> int:
    """
    Description:
        Increments (`opcode=any non-zero value`) or decrements (`opcode=0`) the opacity of the specified window by an (`increment`) value.
        
        If the specified window handle is `0` or `None`, the foreground window is selected.
    
    ---
    Return:
        `int`: The new opacity value. If -1 returned, then the operation has failed.
    """
    ...
