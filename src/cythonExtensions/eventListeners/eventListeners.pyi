"""This contains the keyboard listeners responsible for handling hotkey press and release events."""

from cythonExtensions.hookManager.hookManager import KeyboardEvent

def HotkeyPressEvent(event: KeyboardEvent) -> bool:
    """
    Description:
        The callback function responsible for handling hotkey press events.
    ---
    Parameters:
        `event -> KeyboardEvent`:
            A keyboard event object.
    ---
    Return:
        `suppress_key -> bool`: Whether to suppress the pressed key or return it.
    """
    ...

def ExpanderEvent(event: KeyboardEvent) -> bool:
    """
    Description:
        The callback function responsible for handling the word expansion events.
    ---
    Parameters:
        `event -> KeyboardEvent`:
            A keyboard event object.
    ---
    Output:
        `bool`: Always return True.
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
        `event -> KeyboardEvent`:
            A keyboard event object.
    ---
    Output:
        `bool`: Always return False.
    """
    ...
