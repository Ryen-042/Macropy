"""This contains the keyboard listeners responsible for handling hotkey press and release events."""

from pyWinhook.HookManager import KeyboardEvent

def HotkeyPressEvent(event: KeyboardEvent) -> bool:
    """
    Description:
        The callback function responsible for handling hotkey press events.
    ---
    Parameters:
        `event`:
            A keyboard event object.
    ---
    Return:
        `return_the_key -> bool`: Whether to return or suppress the pressed key.
    """
    ...

def ExpanderEvent(event: KeyboardEvent) -> bool:
    """
    Description:
        The callback function responsible for handling the word expansion events.
    ---
    Parameters:
        `event`:
            A keyboard event object.
    """
    ...

def KeyPress(event: KeyboardEvent) -> bool:
    """The main callback function for handling the `KeyPress` events. Used mainly for allowing multiple `KeyPress` callbacks to run simultaneously."""
    ...

def KeyRelease(event: KeyboardEvent) -> bool:
    """
    Description:
        The callback function responsible for handling the `KeyRelease` events.
    ---
    Parameters:
        `event`:
            A keyboard event object.
    """
    ...
