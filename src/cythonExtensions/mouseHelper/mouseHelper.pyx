# cython: embedsignature = True
# cython: language_level = 3str

"""This module contains functions for controlling the mouse."""

import win32gui, win32api, win32con

cpdef void sendMouseClick(int x=0, int y=0, int button=1, int op=1):
    """
    Description:
        Sends a mouse click to the given position.
        If no position is specified, the current cursor position is used.
    ---
    Parameters:
        
        `button -> int`: Specifies which mouse button to click.
            `1`: left.
            
            `2`: right.
        
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
    if op in [1, 2]:
        clickUp = win32con.MOUSEEVENTF_LEFTDOWN if button == 1 else win32con.MOUSEEVENTF_RIGHTDOWN
        win32api.mouse_event(clickUp, x, y, 0, 0)
    
    # Mouse up event.
    if op in [1, 3]:
        clickDown = win32con.MOUSEEVENTF_LEFTUP if button == 1 else win32con.MOUSEEVENTF_RIGHTUP
        win32api.mouse_event(clickDown, x, y, 0, 0)

# Source: https://stackoverflow.com/questions/34012543/mouse-click-without-moving-cursor?answertab=scoredesc#tab-top
cpdef void sendMouseClickToWindow(int hwnd, int x, int y, int button=1):
    """Sends a mouse click (`button` -> `1: left`, `2: right`) to the given location in the specified window without moving the mouse cursor."""
    
    cdef int l_param = win32api.MAKELONG(x, y)
    
    if button == 1:
        win32gui.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, l_param)
        win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, l_param)
    
    elif button == 2:
        win32gui.PostMessage(hwnd, win32con.WM_RBUTTONDOWN, 0, l_param)
        win32gui.PostMessage(hwnd, win32con.WM_RBUTTONUP, 0, l_param)

# API: https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-mouse_event
cpdef void moveCursor(int dx=0, int dy=0):
    """Moves the mouse cursor some distance `(dx, dy)` away from its current position."""
    
    cdef int x, y
    
    x, y= win32api.GetCursorPos()
    win32api.SetCursorPos((x + dx, y + dy))


