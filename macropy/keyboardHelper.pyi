"""This module provides functions for manipulating keyboard presses and text expansion."""

from typing import Callable, Any

extended_keys: set[int]

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

def FindAndSendKeysToWindow(target_className: str, key: Any, send_function: Callable[[Any], None]) -> None:
    """
    Description:
        Searches for a window with the specified class name, and, if found, sends the specified key using the passed function.
    ---
    Parameters:
        `target_className -> str`:
            The class name of the target window.
        
        `key -> Any`:
            A key that is passed to the `send_function`.
        
        `send_function -> Callable[[Any], None]`:
            A function that simulates the specified key.
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

def SendTextWithCaret(text: str, caret="{!}"):
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