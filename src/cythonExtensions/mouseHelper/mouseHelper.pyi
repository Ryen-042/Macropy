"""This module contains functions for controlling the mouse."""


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
    ...


def sendMouseClickToWindow(hwnd: int, x: int, y: int, button=1) -> None:
    """Sends a mouse click (`button` -> `1: left`, `2: right`, `3: middle`) to the given location in the specified window without moving the mouse cursor."""
    ...


def moveCursor(dx=0, dy=0) -> None:
    """Moves the mouse cursor some distance `(dx, dy)` away from its current position."""
    ...
