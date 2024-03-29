"""This module contains the event listeners responsible for handling keyboard keyPress & keyRelease, and mouse events."""
from cythonExtensions.commonUtils.commonUtils import KeyboardEvent, MouseEvent


def reloadHotkeys() -> None:
    """Reloads the defined hotkeys in the `callbacks` module."""
    ...

def keyPress(event: KeyboardEvent) -> bool:
    """
    Description:
        The callback function responsible for handling hotkey press events.
    ---
    Parameters:
        `event -> KeyboardEvent`:
            A keyboard event object.
    ---
    Return:
        `suppressInput -> bool`: Whether to suppress the pressed key or return it.
    """
    ...


def textExpansion(event: KeyboardEvent) -> bool:
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


def buttonPress(event: MouseEvent) -> bool:
    """The callback function responsible for handling the `buttonPress` events."""
    ...
