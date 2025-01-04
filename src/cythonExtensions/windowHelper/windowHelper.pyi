"""This module provides functions for dealing with windows."""

import win32con


def setWindowProperty(hwnd: int, property: str, value: int) -> None:
    """
    Description:
        Sets the value of `property` to `value` for the specified window.
    """
    ...

def getWindowProperty(hwnd: int, property: str) -> str:
    """
    Description:
        Retrieves the value of `property` from the specified window.
    """
    ...

def sendWindowsMessage(hwnd: int, message: int, wParam: int, lParam: int) -> int:
    """
    Description:
        Sends a message to the specified window.
    """
    ...

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
    ...


def getHandleByTitle(title: str) -> int:
    """searches for a window with the specified title and returns its handle if found. Otherwise, returns `0`."""
    ...


def showMessageBox(msg: str, title="Warning", msgbox_type=1, icon=win32con.MB_ICONERROR) -> int:
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
    ...


def alwaysOnTop() -> None:
    """Toggles the alwaysOnTop feature for the active window."""
    ...


# Shake window - Doesn't work if the window is fullscreen
def shakeActiveWindow(cycles=5) -> None:
    """Simulates shake effect on the active window for the specified number of times."""
    ...


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
    ...


def changeWindowOpacity(hwnd=0, opcode=1, increment=5) -> int:
    """
    Description:
        Increments (`opcode=any non-zero value`) or decrements (`opcode=0`) the opacity of the specified window by an (`increment`) value.
        
        If the specified window handle is `0` or `None`, the foreground window is selected.
    
    ---
    Return:
        `int`: The new opacity value. If -1 returned, then the operation has failed.
    """
    ...
