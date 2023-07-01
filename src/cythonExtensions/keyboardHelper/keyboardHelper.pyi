"""This module provides functions for manipulating keyboard presses and text expansion."""

import win32con
from typing import Callable, Any

# We must check before sending keys using keybd_event: https://stackoverflow.com/questions/21197257/keybd-event-keyeventf-extendedkey-explanation-required
extended_keys = {
    win32con.VK_RMENU,      # "Rmenue"     # -> Right Alt
    win32con.VK_RCONTROL,   # "Rcontrol"
    win32con.VK_RSHIFT,     # "Rshift"
    win32con.VK_APPS,       # "Apps"       # -> Menu
    win32con.VK_VOLUME_UP,  # "Volume_Up"
    win32con.VK_VOLUME_DOWN,# "Volume_Down"
    win32con.VK_SNAPSHOT,   # "Snapshot"
    win32con.VK_INSERT,     # "Insert"
    win32con.VK_DELETE,     # "Delete"
    win32con.VK_HOME,       # "Home"
    win32con.VK_END,        # "End"
    win32con.VK_CANCEL,     # "Break"
    win32con.VK_PRIOR,      # "Prior"
    win32con.VK_NEXT,       # "Next"
    win32con.VK_UP,         # "Up"
    win32con.VK_DOWN,       # "DOWN"
    win32con.VK_LEFT,       # "LEFT"
    win32con.VK_RIGHT       # "RIGHT"
}


def SimulateKeyPress(key_id: int, key_scancode=0, times=1) -> None:
    """
    Description:
        Simulates key press by sending the specified key for the specified number of times.
    ---
    Parameters:
        `key_id -> int`: the id of the key.
        `key_scancode -> int`: the scancode of the key. Optional.
        `times -> int`: the number of times the key should be pressed.
    """
    ...


def SimulateKeyPressSequence(keys_list: tuple[tuple[int, int] | tuple[Any, Callable[[Any], None]]], delay=0.2) -> None:
    """
    Description:
        Simulating a sequence of key presses.
    ---
    Parameters:
        - `keys_list -> tuple[tuple[int, int] | tuple[Any, Callable[[Any], None]]]`:
            - Two numbers representing the keyID and the scancode, or
            - A key and a function that is used to simulate this key.
        
        - `delay -> float`: The delay between key presses.
    """
    ...


def FindAndSendKeyToWindow(target_className: str, key: Any, send_function: Callable[[Any], None] | None) -> int:
    """
    Description:
        Searches for a window with the specified class name, and, if found, sends the specified key using the passed function.
    ---
    Parameters:
        `target_className -> str`:
            The class name of the target window.
        
        `key -> Any`:
            A key that is passed to the `send_function`.
        
        `send_function -> Callable[[Any], Any] | None`:
            A function that simulates the specified key. If not set, `PostMessage` is used.
    """
    ...


def SimulateHotKeyPress(keys_id_dict: dict[int, int]) -> None:
    """
    Description:
        Simulates hotkey press by sending a keyDown event for each of the specified keys and then sending keyUp events.
    ---
    Parameters:
        `keys_id_dict -> dict[int, int]`:
            Holds the `keyID` and `scancode` of the specified keys.
    """
    ...


def GetCaretPosition(text: str, caret="{!}") -> int:
    """Returns the position of the caret in the given text."""
    ...


def SendTextWithCaret(text: str, caret="{!}") -> None:
    """
    Description:
        - Sends (writes) the specified string to the active window.
        - If the string contains one or more carets:
            - The first caret will be deleted, and
            - The keyboard cursor will be placed where the deleted caret was.
    """
    ...


def ExpandText() -> None:
    """Replacing an abbreviated text with its respective substitution specified by the pressed characters `ctrlHouse.pressed_chars`."""
    ...


def OpenLocation() -> None:
    """Opens a file or a directory specified by the pressed characters `ctrlHouse.pressed_chars`."""
    ...


def CrudeOpenWith(tool_number=4, prog_index=0) -> None:
    """
        Description:
            A crude way for opening a program by using the `open` tool from the `Quick Access Toolbar`.
        ---
        Parameters:
            `tool_number -> int`:
                The number associated with the `open` tool in the `Quick Access Toolbar`.
                To find it, press the alt key while selecting a file in the windows explorer.
            
            `prog_index -> int`:
                The index of the program in the `open` tool dropdown menu.
    """
    ...
