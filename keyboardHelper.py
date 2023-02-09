"""This modules provides functions for manipulating keyboard presses and text expansion."""

import win32gui, win32ui, win32api, win32con, winsound, pywintypes
import keyboard
from common import WindowHouse as winHouse, KB_Con as kbcon
from time import sleep
from typing import Callable

# We must check before sending keys using keybd_event: https://stackoverflow.com/questions/21197257/keybd-event-keyeventf-extendedkey-explanation-required
extended_keys = {
                    win32con.VK_RMENU,      # "Rmenue": # -> Right Alt
                    win32con.VK_RCONTROL,   # "Rcontrol":
                    win32con.VK_RSHIFT,     # "Rshift":
                    win32con.VK_APPS,       # "Apps": # -> Menu
                    win32con.VK_VOLUME_UP,  # "Volume_Up":
                    win32con.VK_VOLUME_DOWN,# "Volume_Down":
                    win32con.VK_SNAPSHOT,   # "Snapshot":
                    win32con.VK_INSERT,     # "Insert":
                    win32con.VK_DELETE,     # "Delete":
                    win32con.VK_HOME,       # "Home":
                    win32con.VK_END,        # "End":
                    win32con.VK_CANCEL,     # "Break":
                    win32con.VK_PRIOR,      # "Prior":
                    win32con.VK_NEXT,       # "Next":
                    win32con.VK_UP,         # "Up":
                    win32con.VK_DOWN,       # "DOWN":
                    win32con.VK_LEFT,       # "LEFT":
                    win32con.VK_RIGHT       # "RIGHT":
}

def SimulateKeyPress(key_id, key_scancode=0, times=1):
    """
    Description:
        Simulates key press by sending the specified key for the specified number of times.
    ---
    Parameters:
        `key_id -> int`: the id of the key.
        `key_scancode -> int`: the scancode of the key. Optional.
        `times -> int`: the number of times the key should be pressed.
    """
    
    # Check if the key is one of the extended keys.
    flags = (key_id in extended_keys) * win32con.KEYEVENTF_EXTENDEDKEY
    
    # Simulating keypress.
    for _ in range(times):
        win32api.keybd_event(key_id, key_scancode, flags, 0) # Simulate KeyDown event.
        win32api.keybd_event(key_id, key_scancode, flags | win32con.KEYEVENTF_KEYUP, 0) # Simulate KeyUp event.

def SimulateKeyPressSequence(keys_list: tuple[tuple[int, int] | tuple[int, Callable]], delay=0.2):
    """
    Description:
        Simulating a sequence of key presses.
    ---
    Parameters:
        - `keys_list -> tuple[tuple[int, int] | tuple[int, Callable]]`:
            - Two numbers representing the keyID and the scancode.
            
            - A key value and a function that will be used to simulate send this key.
        
        * `delay -> float`: The delay between key presses.
    """
    
    # Possible alternative: keyboard.send('alt, 4, down, down, down')
    
    for key, scancode in keys_list:
        flags = (key in extended_keys) * win32con.KEYEVENTF_EXTENDEDKEY
        
        # scancode can be either int or Callable.
        if isinstance(scancode, int):
            win32api.keybd_event(key, scancode, flags, 0) # Simulate KeyDown event.
            win32api.keybd_event(key, scancode, flags | win32con.KEYEVENTF_KEYUP, 0) # Simulate KeyUp event.
        else:
            # If `scancode` is not a number, then it is a callable.
            scancode(key)
        sleep(delay)

def FindAndSendKeysToWindow(target_className:str, key, send_function) -> None:
    """
    Description:
        Searches for a window with the specified class name, and, if found, sends the specified key using the passed function.
    ---
    Parameters:
        `target_className -> str`:
            The class name of the target window.
        
        `key -> Any`:
            A key that is passed to the `send_function`.
        
        `send_function -> Callable`:
            A function that simulates the specified key.
    """
    
    # Checking if there was a stored window handle for the specified window class name.
    target_hwnd = winHouse.GetClassNameHandle(target_className)
    
    # Checking if the window associated with `target_hwnd` does exist. If not, try searching for one.
    if not target_hwnd or not win32gui.IsWindowVisible(target_hwnd):
        ## Method(1) for searching for a window handle given a class name.
        # winHouse.SetClassNameHandle(target_className, win32Helper.HandleByClassName(target_className))
        # If `None`, then there is no such window.
        # if not winHouse.GetClassNameHandle[target_className]: # win32gui.IsWindowVisible()
        #     return
        
        ## Method(2) for searching for a window handle given a class name.
        try:
            target_window = win32ui.FindWindow(target_className, None)
        except win32ui.error: # Window not found.
            winHouse.SetClassNameHandle(target_className, None)
            return
        
        winHouse.SetClassNameHandle(target_className, target_window.GetSafeHwnd())
    
    target_hwnd = winHouse.GetClassNameHandle(target_className)
    
    ##  Method(1) for setting focus to a specific window. Doesn't work if the window is visible, only if it was minimized.
    # target_window.ShowWindow(1)
    # target_window.SetForegroundWindow()
    
    ## Method (No.2) for setting focus to a specific window. Works if the window is minimized or visible: https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindow
    win32gui.ShowWindow(target_hwnd, win32con.SW_RESTORE)
    
    # Sometimes it seems to fail but then work after another call.
    try:
        win32gui.SetForegroundWindow(target_hwnd)
    except pywintypes.error:
        sleep(0.5)
        win32gui.ShowWindow(target_hwnd, win32con.SW_RESTORE)
        
        try:
            win32gui.SetForegroundWindow(target_hwnd)
        except pywintypes.error:
            print(f"Exception occurred while trying set {win32gui.GetWindowText(target_hwnd)} as the forground process.")
            return
    
    send_function(key)

def SimulateHotKeyPress(keys_id_dict: dict[int, int]):
    """
    Description:
        Simulates hotkey press by sending a keyDown event for each specified key and then sending keyUp events.
    ---
    Parameters:
        `keys_id_dict -> dict(int, int)`:
            Holds the keyID and scancode of the specified keys.
    """
    
    for key_id, key_scancode in keys_id_dict.items():
        flags = (key_id in extended_keys) * win32con.KEYEVENTF_EXTENDEDKEY
        win32api.keybd_event(key_id, key_scancode, flags, 0) # Simulate KeyDown event.
    
    for key_id, key_scancode in keys_id_dict.items():
        flags = (key_id in extended_keys) * win32con.KEYEVENTF_EXTENDEDKEY
        win32api.keybd_event(key_id, key_scancode, flags | win32con.KEYEVENTF_KEYUP, 0) # Simulate KeyUp event.

def GetCaretPosition(text: str, caret="{!}"):
    """Returns the position of the caret in the given text."""
    
    caret_pos = text.find(caret)
    return (caret_pos != -1 and caret_pos) or len(text)

def SendTextWithCaret(text, caret="{!}"):
    """
    Description:
        - Sends (writes) the specified string to the active window. 
        - If the string contains one or more carets:
            - The first caret will be deleted, and
            - The keyboard cursor will be placed where the deleted caret was.
    """
    
    caret_pos = GetCaretPosition(text, caret)
    
    text = text[:caret_pos] + text[caret_pos+len(caret):]
    keyboard.write(text)
    
    if caret_pos == len(text):
        return
    
    # This doesn't work properly if the caret is not at the beggining (`HOME`) before inserting the text.
    # mid_pos = len(text) // 2
    # if caret_pos <= mid_pos:
    #     SimulateKeyPress(win32con.VK_HOME,  kbcon.SC_HOME)
    #     SimulateKeyPress(win32con.VK_RIGHT, kbcon.SC_RIGHT, caret_pos)
    # else:
    SimulateKeyPress(win32con.VK_LEFT, kbcon.SC_LEFT, len(text) - caret_pos)

def ExpandText(text, abbr_len):
    """Substitutes an abbreviated text with another one. The abbreviated text is represented by its length."""
    
    # Send ('`' => "Oem_3") then delete it before expansion to silence any suggestions like in the browser.
    SimulateKeyPress(kbcon.VK_BACKTICK, kbcon.SC_BACKTICK)
    
    # Delete the abbreviation and the '`' character.
    SimulateKeyPress(win32con.VK_BACK, kbcon.SC_BACK, abbr_len+1)
    
    # Substitute the abbreviation with the specified text.
    keyboard.write(text)
    winsound.PlaySound(r"SFX\knob-458.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
