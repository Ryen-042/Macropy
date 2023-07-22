# cython: embedsignature = True
# cython: language_level = 3str

"""This extension module contains the event listeners responsible for handling keyboard keyPress & keyRelease, and mouse events."""

from cythonExtensions.commonUtils.commonUtils cimport KeyboardEvent, MouseEvent, KB_Con as kbcon
from cythonExtensions.commonUtils.commonUtils import ControllerHouse as ctrlHouse, WindowHouse as winHouse, PThread, Management as mgmt, printModifiers, printLockKeys, printButtons

import win32gui, win32con, threading

from cythonExtensions.eventHandlers.callbacks.callbacks import kbEventHandlers, kbEventHandlersWithLockKeys
from cythonExtensions.keyboardHelper import keyboardHelper as kbHelper


def keyPress(KeyboardEvent event) -> bool:
    """
    Description:
        The callback function responsible for handling hotkey press events.
    ---
    Parameters:
        `event -> KeyboardEvent`:
            A keyboard event object.
    ---
    Return:
        `suppressInput -> bool`: Whether to suppress the last pressed key or return it.
    """
    
    cdef bint suppressInput = mgmt.suppressKbInputs # and not (ctrlHouse.modifiers & ctrlHouse.FN) and
    
    eventHandler = kbEventHandlers.get((ctrlHouse.modifiers, event.KeyID)) or \
       kbEventHandlersWithLockKeys.get((ctrlHouse.locks, ctrlHouse.modifiers, event.KeyID))
    
    if eventHandler:
        PThread(target=eventHandler[0], args=eventHandler[1]).start()
        suppressInput = True
    
    elif ((ctrlHouse.modifiers & ctrlHouse.CTRL) and event.KeyID == kbcon.VK_W) or ((ctrlHouse.modifiers & ctrlHouse.ALT) and event.KeyID == win32con.VK_F4):
        try:
            fg_hwnd = win32gui.GetForegroundWindow()
            if win32gui.GetClassName(fg_hwnd) in ("CabinetWClass", "WorkerW"):
                PThread(target=winHouse.rememberActiveProcessTitle, args=(fg_hwnd,)).start()
        
        except Exception as e:
            print(f"Error: {e}\nA problem occurred while retrieving the className of the foreground window.\n")
    
    
    PThread.kbMsgQueue.put(suppressInput)
    
    return suppressInput


### Word listening operations ###
def textExpansion(KeyboardEvent event) -> bool:
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
    
    #+ Printing some relevant information about the pressed key and hardware metrics.
    if not mgmt.silent:
        print(f"Thread Count: {threading.active_count()} |", event, f"| Counter={mgmt.counter}")
        printModifiers()
        printLockKeys()
        print("")
        
        # sysHelper.displayCPU_Usage(), print("\n")
        
        # This counter used for debugging purposes only.
        mgmt.counter = (mgmt.counter + 1) % 10000
    
    #+ If the first stored character is not one of the defined prefixes, don't check anything else; this is not a valid word for expansion.
    if not ctrlHouse.pressed_chars.startswith((":", "!")):
        if event.Ascii == kbcon.AS_COLON:
            ctrlHouse.pressed_chars = ":"
        
        elif event.Ascii == kbcon.AS_EXCLAM:
            ctrlHouse.pressed_chars = "!"
        
        # If all keys are suppressed, ":" and "!" cannot be pressed (because shift is suppressed, along with the ";" and "1" keys). Thus, we need to check if shift is pressed.
        # Any new prefixes that require shift to be pressed should be added here.
        elif (ctrlHouse.modifiers & ctrlHouse.SHIFT):
            if event.KeyID == kbcon.VK_SEMICOLON:
                ctrlHouse.pressed_chars = ":"
            
            elif event.KeyID == kbcon.VK_1:
                ctrlHouse.pressed_chars = "!"
        
        else:
            return True
        
        print(ctrlHouse.pressed_chars, end="\r")
        return True
    
    #- Most of the function and modifier keys (ctrl, shift, end, home, etc...) have zero Ascii values.
    #- Also, there are some other keys with non-zero Ascii values. These keys are [Space, Enter, Tab, ESC].
    #? If one of the above is true, then return the key without checking anything else. Note that the `Enter` key is returned as `\r`.
    if event.Ascii == 0 or event.KeyID in (win32con.VK_SPACE, win32con.VK_RETURN, win32con.VK_TAB, win32con.VK_ESCAPE):
        # Clear the displayed pressed character keys.
        print(" " * len(ctrlHouse.pressed_chars), end="\r")
        
        # To allow for `ctrl` + character keys, you need to check here for them individually. The same `event.Key` value is returned but `event.Ascii` is 0.
        ctrlHouse.pressed_chars = "" # Resetting the stored pressed keys.
        
        return True
    
    ## Filtering all symbols.
    # elif len(event.Key) > 1:
        # print("A symbol was pressed: {}".format(chr(event.Ascii))
    
    #+ Pop characters if `backspace` is pressed.
    if event.KeyID == win32con.VK_BACK and len(ctrlHouse.pressed_chars):
        ctrlHouse.pressed_chars = ctrlHouse.pressed_chars[:-1] # Remove the last character
        print(ctrlHouse.pressed_chars + " ", end="\r")
        return True
    
    #+ If the key is not filtered above, then it is a valid character key.
    ctrlHouse.pressed_chars += event.Key.lower() # chr(event.Ascii).lower()
    print(ctrlHouse.pressed_chars, end="\r")
    
    #+ Check if the pressed characters match any of the defined abbreviations, and if so, replace them accordingly.
    if ctrlHouse.pressed_chars in ctrlHouse.abbreviations:
        print(" " * len(ctrlHouse.pressed_chars), end="\r")
        kbHelper.expandText()
    
    ### Executing some operations based on the pressed characters. ###
    #+ Opening a file or a directory.
    elif ctrlHouse.pressed_chars in ctrlHouse.locations:
        print(" " * len(ctrlHouse.pressed_chars), end="\r")
        kbHelper.openLocation()
    
    #+ This is a crude way of opening a file using a specific program (open with).
    elif ctrlHouse.pressed_chars == ":\\\\":
        # PThread(target=kbHelper.simulateKeyPressSequence, args=(((win32con.VK_LMENU, 56), (52, 5), (40, 80), (40, 80), (13, 28)))).start()
        print(" " * len(ctrlHouse.pressed_chars), end="\r")
        kbHelper.crudeOpenWith(4, 2)
    
    return True


def buttonPress(MouseEvent event) -> bool:
    """The callback function responsible for handling the `buttonPress` events."""
    
    if not mgmt.silent and not event.Delta:
        printButtons()
        # print(event)
    
    cdef bint suppressInput = False
    
    if event.Delta and (ctrlHouse.modifiers & ctrlHouse.CTRL_SHIFT) == ctrlHouse.CTRL_SHIFT:
        PThread(target=kbHelper.simulateKeyPress, args=((win32con.VK_VOLUME_DOWN, win32con.VK_VOLUME_UP)[event.Delta > 0],)).start()
        
        suppressInput = True
    
    PThread.msMsgQueue.put(suppressInput)
    
    return suppressInput
