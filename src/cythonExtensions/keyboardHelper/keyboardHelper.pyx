# cython: language_level = 3str

"""This extension module provides functions for manipulating keyboard presses and text expansion."""


import win32gui, win32ui, win32api, win32con, winsound, pywintypes
import keyboard, os
from time import sleep

from cythonExtensions.commonUtils.commonUtils import KB_Con as kbcon, WindowHouse as winHouse, ControllerHouse as ctrlHouse
from cythonExtensions.guiHelper.inputWindow import SimpleWindow

# We must check before sending keys using keybd_event: https://stackoverflow.com/questions/21197257/keybd-event-keyeventf-extendedkey-explanation-required
cdef set extended_keys = {
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

cpdef void simulateKeyPress(int key_id, int key_scancode=0, int times=1):
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

cpdef void simulateKeyPressSequence(tuple keys_list, float delay=0.2):
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

def findAndSendKeyToWindow(target_className: str, key, send_function=None) -> int:
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
    ---
    Returns:
        `int`: 1 if the window was found and the key was sent, 0 otherwise.
    """
    
    # Checking if there was a stored window handle for the specified window class name.
    cdef int hwnd = winHouse.getHandleByClassName(target_className)
    
    # Checking if the window associated with `hwnd` does exist. If not, try searching for one.
    if not hwnd or not win32gui.IsWindowVisible(hwnd):
        ## Method(1) for searching for a window handle given a class name.
        # winHouse.setHandleByClassName(target_className, win32Helper.HandleByClassName(target_className))
        # If `None`, then there is no such window.
        # if not winHouse.getHandleByClassName[target_className]: # win32gui.IsWindowVisible()
        #     return
        
        ## Method(2) for searching for a window handle given a class name.
        try:
            target_window = win32ui.FindWindow(target_className, None)
        except win32ui.error: # Window not found.
            print(f"Window with class name '{target_className}' not found.")
            
            winHouse.setHandleByClassName(target_className, 0)
            
            return 0
        
        winHouse.setHandleByClassName(target_className, target_window.GetSafeHwnd())
    
    hwnd = winHouse.getHandleByClassName(target_className)
    
    ##  Method(1) for setting focus to a specific window. Doesn't work if the window is visible, only if it was minimized.
    # target_window.ShowWindow(1)
    # target_window.SetForegroundWindow()
    
    ## SetFocus(hwnd)
    ## BringWindowToTop(hwnd)
    ## SetForegroundWindow(hwnd)
    ## ShowWindow(hwnd, 1)
    
    ## Method (No.2) for setting focus to a specific window. Works if the window is minimized or visible: https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindow
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    
    if send_function:
        
        # Sometimes SetForegroundWindow seems to fail but then work after another call.
        try:
            win32gui.SetForegroundWindow(hwnd)
            
        except pywintypes.error:
            sleep(0.5)
            
            try:
                win32gui.SetForegroundWindow(hwnd)
            
            except pywintypes.error:
                print(f"Exception occurred while trying set {win32gui.GetWindowText(hwnd)} as the forground process.")
                return 0
    
        send_function(key)
    
    else:
        win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, key, 0)
    
    return 1

def simulateHotKeyPress(keys_id_dict: dict[int, int]) -> None:
    """
    Description:
        Simulates hotkey press by sending a keyDown event for each of the specified keys and then sending keyUp events.
    ---
    Parameters:
        `keys_id_dict -> dict[int, int]`:
            Holds the `keyID` and `scancode` of the specified keys.
    """
    cdef int key_id, key_scancode, flags
    
    for key_id, key_scancode in keys_id_dict.items():
        flags = (key_id in extended_keys) * win32con.KEYEVENTF_EXTENDEDKEY
        win32api.keybd_event(key_id, key_scancode, flags, 0) # Simulate KeyDown event.
    
    for key_id, key_scancode in keys_id_dict.items():
        flags = (key_id in extended_keys) * win32con.KEYEVENTF_EXTENDEDKEY
        win32api.keybd_event(key_id, key_scancode, flags | win32con.KEYEVENTF_KEYUP, 0) # Simulate KeyUp event.

def simulateBurstClicks():
    cdef int flags
    cdef float delay
    
    window = SimpleWindow("Key & Delay Input", itemsHeight=30)
    # window.createDynamicInputWindow(["Key", "Delay (ms)"])
    window.createDynamicInputWindow(
        input_labels = ["Delay (ms)"],
        placeholders = ["100"],
        enable_key_capture = True
    )
    
    # If the user didn't provide any input, for example, by closing the window.
    if not window.userInputs or not window.capturedKeyVK:
        return
    
    flags = (window.capturedKeyVK in extended_keys) * win32con.KEYEVENTF_EXTENDEDKEY
    
    delay = int(window.userInputs[0]) / 1000
    
    while ctrlHouse.burstClicksActive:
        win32api.keybd_event(window.capturedKeyVK, 0, flags, 0) # Simulate KeyDown event.
        win32api.keybd_event(window.capturedKeyVK, 0, flags | win32con.KEYEVENTF_KEYUP, 0) # Simulate KeyUp event.
        sleep(delay)

def resetModifierKeys() -> None:
    """Reset the modifer keys bey sending keyUp events."""
    
    modifiers = [
        (win32con.VK_LCONTROL, 29),
        (win32con.VK_LSHIFT, 42),
        (win32con.VK_LMENU, 56),
        (win32con.VK_LWIN, 91),
        (255, 0),  # FN
        (kbcon.VK_BACKTICK, 41)
    ]
    
    for key_id, key_scancode in modifiers:
        flags = (key_id in extended_keys) * win32con.KEYEVENTF_EXTENDEDKEY
        win32api.keybd_event(key_id, key_scancode, win32con.KEYEVENTF_KEYUP | flags, 0)

cdef int getCaretPosition(text, caret="{!}"):
    """Returns the position of the caret in the given text."""
    # caret_pos = text[::-1].find("}!{")
    # return (caret_pos != -1 and len(text) - caret_pos) or -1
    
    cdef int caret_pos = text.find(caret)
    return (caret_pos != -1 and caret_pos) or len(text)

def sendTextWithCaret(text: str, caret="{!}") -> None:
    """
    Description:
        - Sends (writes) the specified string to the active window.
        - If the string contains one or more carets:
            - The first caret will be removed, and
            - The keyboard cursor will be placed where the removed caret was.
    """
    
    cdef int caret_pos = getCaretPosition(text, caret)
    
    text = text[:caret_pos] + text[caret_pos+len(caret):]
    keyboard.write(text)
    
    if caret_pos == len(text):
        return
    
    # This doesn't work properly if the caret is not at the beggining (`HOME`) before inserting the text.
    # mid_pos = len(text) // 2
    # if caret_pos <= mid_pos:
    #     simulateKeyPress(win32con.VK_HOME,  kbcon.SC_HOME)
    #     simulateKeyPress(win32con.VK_RIGHT, kbcon.SC_RIGHT, caret_pos)
    # else:
    simulateKeyPress(win32con.VK_LEFT, kbcon.SC_LEFT, len(text) - caret_pos)

def expandText() -> None:
    """Replacing an abbreviated text with its respective substitution text."""
    
    # Sending ('`' => "Oem_3") then delete it before expansion to silence any suggestions like in the browser address bar.
    simulateKeyPress(kbcon.VK_BACKTICK, kbcon.SC_BACKTICK)
    
    # Deleting the abbreviation and the '`' character.
    simulateKeyPress(win32con.VK_BACK, kbcon.SC_BACK, len(ctrlHouse.pressed_chars) + 1)
    
    # Substituting the abbreviation with its respective text.
    keyboard.write(ctrlHouse.abbreviations.get(ctrlHouse.pressed_chars, ctrlHouse.non_prefixed_abbreviations.get(ctrlHouse.pressed_chars)))
    
    # Resetting the stored pressed keys.
    ctrlHouse.pressed_chars = ""
    
    winsound.PlaySound(r"SFX\knob-458.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)

def undoTextExpansion() -> None:
    """Undoes text expansion by replacing it with its abbreviation."""
    
    text = ctrlHouse.abbreviations.get(ctrlHouse.pressed_chars, ctrlHouse.non_prefixed_abbreviations.get(ctrlHouse.pressed_chars))
    
    # Deleting the expansion.
    simulateKeyPress(win32con.VK_BACK, kbcon.SC_BACK, len(text))
    
    # Replacing the expansion with the abbreviation.
    keyboard.write(ctrlHouse.pressed_chars)
    
    winsound.PlaySound(r"SFX\undo.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)

def openLocation() -> None:
    """Opens a file or a directory specified by the pressed characters `ctrlHouse.pressed_chars`."""
    
    # Opening the file/folder.
    os.startfile(ctrlHouse.locations.get(ctrlHouse.pressed_chars))
    
    # Resetting the stored pressed keys.
    ctrlHouse.pressed_chars = ""
    
    winsound.PlaySound(r"C:\Windows\Media\Windows Navigation Start.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)

def crudeOpenWith(tool_number=4, prog_index=0) -> None:
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
    simulateKeyPressSequence(((f"alt+{tool_number}", keyboard.send), *((win32con.VK_DOWN, kbcon.SC_DOWN), ) * prog_index, (win32con.VK_RETURN, kbcon.SC_RETURN)))
    
    # Resetting the stored pressed keys.
    ctrlHouse.pressed_chars = ""
    
    winsound.PlaySound(r"SFX\knob-458.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
