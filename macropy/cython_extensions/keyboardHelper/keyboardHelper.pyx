# cython: embedsignature = True
# cython: language_level = 3str

"""This extension module provides functions for manipulating keyboard presses and text expansion."""

import win32gui, win32ui, win32api, win32con, winsound, pywintypes
import keyboard, os
from common import WindowHouse as winHouse, ControllerHouse as ctrlHouse, KB_Con as kbcon
from time import sleep

# We must check before sending keys using keybd_event: https://stackoverflow.com/questions/21197257/keybd-event-keyeventf-extendedkey-explanation-required
cdef set extended_keys = {
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

cpdef void SimulateKeyPress(int key_id, int key_scancode=0, int times=1):
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
    cdef int flags = (key_id in extended_keys) * win32con.KEYEVENTF_EXTENDEDKEY
    
    # Simulating keypress.
    for _ in range(times):
        win32api.keybd_event(key_id, key_scancode, flags, 0) # Simulate KeyDown event.
        win32api.keybd_event(key_id, key_scancode, flags | win32con.KEYEVENTF_KEYUP, 0) # Simulate KeyUp event.

cpdef void SimulateKeyPressSequence(tuple keys_list, float delay=0.2):
    """
    Description:
        Simulating a sequence of key presses.
    ---
    Parameters:
        - `keys_list -> tuple[tuple[int, int] | tuple[Any, Callable[[Any], None]]]`:
            - Two numbers representing the keyID and the scancode, or
            - A key and a function that is used to send this key.
        
        - `delay -> float`: The delay between key presses.
    """
    
    # Possible alternative: keyboard.send('alt, 4, down, down, down')
    cdef key
    cdef int flags
    
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

cpdef void FindAndSendKeysToWindow(str target_className, key, send_function):
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
    
    # Checking if there was a stored window handle for the specified window class name.
    cdef int target_hwnd = winHouse.GetHandleByClassName(target_className)
    
    # Checking if the window associated with `target_hwnd` does exist. If not, try searching for one.
    if not target_hwnd or not win32gui.IsWindowVisible(target_hwnd):
        ## Method(1) for searching for a window handle given a class name.
        # winHouse.SetHandleByClassName(target_className, win32Helper.HandleByClassName(target_className))
        # If `None`, then there is no such window.
        # if not winHouse.GetHandleByClassName[target_className]: # win32gui.IsWindowVisible()
        #     return
        
        ## Method(2) for searching for a window handle given a class name.
        try:
            target_window = win32ui.FindWindow(target_className, None)
        except win32ui.error: # Window not found.
            winHouse.SetHandleByClassName(target_className, None)
            return
        
        winHouse.SetHandleByClassName(target_className, target_window.GetSafeHwnd())
    
    target_hwnd = winHouse.GetHandleByClassName(target_className)
    
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
    
    ## SetFocus(target_hwnd)
    ## BringWindowToTop(target_hwnd)
    ## SetForegroundWindow(target_hwnd)
    ## ShowWindow(198160, 1)
    
    send_function(key)

cpdef void SimulateHotKeyPress(dict keys_id_dict):
    """
    Description:
        Simulates hotkey press by sending a keyDown event for each of the specified keys and then sending keyUp events.
    ---
    Parameters:
        `keys_id_dict -> dict(int, int)`:
            Holds the `keyID` and `scancode` of the specified keys.
    """
    cdef int key_id, key_scancode, flags
    
    for key_id, key_scancode in keys_id_dict.items():
        flags = (key_id in extended_keys) * win32con.KEYEVENTF_EXTENDEDKEY
        win32api.keybd_event(key_id, key_scancode, flags, 0) # Simulate KeyDown event.
    
    for key_id, key_scancode in keys_id_dict.items():
        flags = (key_id in extended_keys) * win32con.KEYEVENTF_EXTENDEDKEY
        win32api.keybd_event(key_id, key_scancode, flags | win32con.KEYEVENTF_KEYUP, 0) # Simulate KeyUp event.

cpdef int GetCaretPosition(str text, str caret="{!}"):
    """Returns the position of the caret in the given text."""
    # caret_pos = text[::-1].find("}!{")
    # return (caret_pos != -1 and len(text) - caret_pos) or -1
    
    cdef int caret_pos = text.find(caret)
    return (caret_pos != -1 and caret_pos) or len(text)

cpdef void SendTextWithCaret(str text, str caret="{!}"):
    """
    Description:
        - Sends (writes) the specified string to the active window.
        - If the string contains one or more carets:
            - The first caret will be deleted, and
            - The keyboard cursor will be placed where the deleted caret was.
    """
    
    cdef int caret_pos = GetCaretPosition(text, caret)
    
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

cpdef void ExpandText():
    """Replacing an abbreviated text with its respective substitution specified by the pressed characters `ctrlHouse.pressed_chars`."""
    
    # Sending ('`' => "Oem_3") then delete it before expansion to silence any suggestions like in the browser address bar.
    SimulateKeyPress(kbcon.VK_BACKTICK, kbcon.SC_BACKTICK)
    
    # Deleting the abbreviation and the '`' character.
    SimulateKeyPress(win32con.VK_BACK, kbcon.SC_BACK, len(ctrlHouse.pressed_chars)+1)
    
    # Substituting the abbreviation with its respective text.
    keyboard.write(ctrlHouse.abbreviations.get(ctrlHouse.pressed_chars))
    
    # Resetting the stored pressed keys.
    ctrlHouse.pressed_chars = ""
    
    winsound.PlaySound(r"SFX\knob-458.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)

cpdef void OpenLocation():
    """Opens a file or a directory specified by the pressed characters `ctrlHouse.pressed_chars`."""
    
    # Opening the file/folder.
    os.startfile(ctrlHouse.locations.get(ctrlHouse.pressed_chars))
    
    # Resetting the stored pressed keys.
    ctrlHouse.pressed_chars = ""
    
    winsound.PlaySound(r"SFX\knob-458.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)

cpdef void CrudeOpenWith(int tool_number=4, int prog_index=0):
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
    
    # In my case, the `open` tool is the forth item in the Quick Access Toolbar.
    SimulateKeyPressSequence(((f"alt+{tool_number}", ctrlHouse.keyboard_send), *((win32con.VK_DOWN, kbcon.SC_DOWN), ) * prog_index, (win32con.VK_RETURN, kbcon.SC_RETURN)))
    
    # Resetting the stored pressed keys.
    ctrlHouse.pressed_chars = ""
    
    winsound.PlaySound(r"SFX\knob-458.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
