# cython: language_level = 3str

"""This extension module contains functions for controlling the mouse."""

import win32gui, win32api, win32con

def sendMouseClick(x=0, y=0, button=1, op=1) -> None:
    """
    Description:
        Sends a mouse click to the given position.
        If no position is specified, the current cursor position is used.
    ---
    Parameters:
        
        `button -> int`: Specifies which mouse button to click. One of the following values:
            `1`: left button
            `2`: right button
            `3`: middle button (wheel)
        
        `op -> int`: parameter specifies the type of mouse event to send:
            `1`: send both a mouse down and up events.
            
            `2`: send only a mouse down event.
            
            `3`: send only a mouse up event.
    """
    
    cdef int clickUp, clickDown
    
    # Use current cursor position if (x, y) are not specified.
    if not x | y:
        x, y = win32api.GetCursorPos()
    
    # Mouse down event.
    if op in (1, 2):
        clickDown = [win32con.MOUSEEVENTF_LEFTDOWN, win32con.MOUSEEVENTF_RIGHTDOWN, win32con.MOUSEEVENTF_MIDDLEDOWN][button - 1]
        win32api.mouse_event(clickDown, x, y, 0, 0)
    
    # Mouse up event.
    if op in (1, 3):
        clickUp = [win32con.MOUSEEVENTF_LEFTUP, win32con.MOUSEEVENTF_RIGHTUP, win32con.MOUSEEVENTF_MIDDLEUP][button - 1]
        win32api.mouse_event(clickUp, x, y, 0, 0)

# Source: https://stackoverflow.com/questions/34012543/mouse-click-without-moving-cursor?answertab=scoredesc#tab-top
def sendMouseClickToWindow(hwnd: int, x: int, y: int, button=1) -> None:
    """Sends a mouse click (`button` -> `1: left`, `2: right`, `3: middle`) to the given location in the specified window without moving the mouse cursor."""
    
    cdef int l_param = win32api.MAKELONG(x, y)
    
    if button == 1:
        win32gui.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, l_param)
        win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, l_param)
    
    elif button == 2:
        win32gui.PostMessage(hwnd, win32con.WM_RBUTTONDOWN, 0, l_param)
        win32gui.PostMessage(hwnd, win32con.WM_RBUTTONUP, 0, l_param)
    
    elif button == 3:
        win32gui.PostMessage(hwnd, win32con.WM_MBUTTONDOWN, 0, l_param)
        win32gui.PostMessage(hwnd, win32con.WM_MBUTTONUP, 0, l_param)

# API: https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-mouse_event
def moveCursor(dx=0, dy=0) -> None:
    """Moves the mouse cursor some distance `(dx, dy)` away from its current position."""
    
    cdef int x, y
    
    x, y= win32api.GetCursorPos()
    win32api.SetCursorPos((x + dx, y + dy))


def sendMouseScroll(steps=1, direction=1, wheelDelta=40) -> None:
    """
    Description:
        Sends a mouse scroll event with the specified number of steps and direction.
    ---
    Parameters:
        `steps: int = 1`:
            The number of steps to scroll. Can take negative values.
        
        `direction: int = 1`:
            The direction to scroll in. `1` for vertical, `0` for horizontal.
        
        `wheelDelta: int = 40`:
            The amount of scroll per step.
    """
    
    # keyboard.press("ctrl")
    # ControllerHouse.pynput_mouse.scroll(0, dist)
    # keyboard.release("ctrl")
    
    # API doc: https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-mouse_event
    win32api.mouse_event((win32con.MOUSEEVENTF_HWHEEL, win32con.MOUSEEVENTF_WHEEL)[direction], 0, 0, steps * wheelDelta, 0) # win32con.WHEEL_DELTA