# cython: language_level = 3str

"""This extension module contains the event listeners responsible for handling keyboard keyPress & keyRelease, and mouse events."""

cimport cython
from cythonExtensions.commonUtils.commonUtils cimport KeyboardEvent, MouseEvent

import win32gui, win32con, importlib, winsound, os, subprocess

from cythonExtensions.commonUtils.commonUtils import  KB_Con as kbcon, ControllerHouse as ctrlHouse, MouseHouse as msHouse, PThread, Management as mgmt
from cythonExtensions.eventHandlers import callbacks as cbs
from cythonExtensions.keyboardHelper import keyboardHelper as kbHelper
from cythonExtensions.windowHelper import windowHelper as winHelper
from cythonExtensions.explorerHelper import explorerHelper as expHelper

cdef void reloadHotkeys():
    """Reloads the defined hotkeys in the `callbacks` module."""
    
    importlib.reload(cbs)
    
    winsound.PlaySound(r"SFX\completed-voice-ringtone.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)

cdef bint executeExplorerRelatedCallbacks(eventHandler):
    suppressKeyPress = False
    
    # In some rare cases, after getting foreground window handle, if the window is destroyed before calling the
    # GetClassName function, then it raises an exception.
    try:
        classNames = ("CabinetWClass", "WorkerW", "Progman") if eventHandler[2] else ("CabinetWClass",)
        if win32gui.GetClassName(win32gui.GetForegroundWindow()) in classNames:
            PThread(target=eventHandler[0], args=eventHandler[1]).start()
            
            suppressKeyPress = eventHandler[3]
    
    except Exception as e:
        print(f"Error: {e}\nA problem occurred while retrieving the className of the foreground window.\n")
    
    return suppressKeyPress

cdef bint openImageViewer():
    if win32gui.GetClassName(win32gui.GetForegroundWindow()) not in ("CabinetWClass", "WorkerW", "Progman"):
        return False
    
    selectedImage = expHelper.getSelectedItemsFromActiveExplorer(None, ("jpg", "png", "jpeg", "ico", "bmp", "gif", "webp"))
    if not selectedImage:
        return False
    
    hwnd = win32gui.FindWindow("ImageViewerClass", None) # winHelper.findHandleByClassName("ImageViewerClass", False)
    
    if hwnd:
        # If the window title (which is the image path) is the same as the selected image, then close the window.
        if win32gui.GetWindowText(hwnd) == selectedImage[0]:
            win32gui.SendMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            return False
        
        WM_SETIMAGE = win32con.WM_USER + 1
        PThread(target=subprocess.call, args=(("python", "-c", f"from cythonExtensions.guiHelper.imageViewer import BuildString; BuildString({hwnd}, {WM_SETIMAGE}).sendString(r\"{selectedImage[0]}\")"),)).start()
        
        return False
    
    else:
        # If the image viewer is not already running, then start it.
        PThread(target=subprocess.call, args=(("python", "-c", f"from cythonExtensions.guiHelper.imageViewer import ImageViewer; ImageViewer(r\"{selectedImage[0]}\", r\"{selectedImage[0]}\")"),)).start()
        
        return False

def keyPress(KeyboardEvent event) -> bool:
    """
    Description:
        The callback function that maps the pressed keys to their corresponding functions and executes them.
    ---
    Parameters:
        `event -> KeyboardEvent`:
            A keyboard event object.
    ---
    Return:
        `suppressKeyPress -> bool`: Whether to suppress the last pressed key or return it.
    """
    
    # This is used to prevent the `backtick` key from being sent when trying to trigger an internal hotkey.
    # We also need to check if a modifier is pressed to prevent blocking any external hotkeys from other applications that use the `backtick` key.
    if event.KeyID == kbcon.VK_BACKTICK:
        mgmt.isBacktickTheOnlyModiferPressed = ctrlHouse.modifiers == ctrlHouse.BACKTICK
    # elif ctrlHouse.modifiers & ctrlHouse.BACKTICK: # A key is pressed while the backtick is pressed, so don't send backtick when it is released.
    else: # A key is pressed while the backtick is pressed, so don't send backtick when it is released.
        mgmt.isBacktickTheOnlyModiferPressed = False
    
    # # Setting the shared variable to indicate that the mouse volume control hotkey (Ctrl + Shift) is pressed.
    # if ctrlHouse.modifiers & ctrlHouse.CTRL_SHIFT:
    #     mgmt.mouseVolumeControlSVar.value = True
    
    cdef bint suppressKeyPress = mgmt.suppressKbInputs or mgmt.isBacktickTheOnlyModiferPressed # and not (ctrlHouse.modifiers & ctrlHouse.FN) and
    
    eventHandler = cbs.kbEventHandlers.get((ctrlHouse.modifiers, event.KeyID)) or \
        (ctrlHouse.SCROLL and cbs.kbEventHandlersWithSCROLL_On.get((ctrlHouse.modifiers, event.KeyID))) or \
        cbs.kbEventHandlersWithExplorerFocus.get((ctrlHouse.modifiers, event.KeyID))
    
    if eventHandler:
        if len(eventHandler) == 2:
            PThread(target=eventHandler[0], args=eventHandler[1]).start()
            suppressKeyPress = True
        
        elif len(eventHandler) == 4:
            executeExplorerRelatedCallbacks(eventHandler)
    
    #+ Opening the selected image file from the active explorer window: 'Space'
    elif not ctrlHouse.modifiers and event.KeyID == win32con.VK_SPACE:
        suppressKeyPress |= openImageViewer()
    
    #+ Reload the hotkeys: Ctrl + Alt + Win + 'R'*
    elif (ctrlHouse.modifiers & ctrlHouse.CTRL_ALT_WIN) == ctrlHouse.CTRL_ALT_WIN and event.KeyID == kbcon.VK_R:
        PThread(target=reloadHotkeys).start()
        
        suppressKeyPress = True
    
    elif ctrlHouse.burstClicksActive and event.KeyID == win32con.VK_ESCAPE:
        ctrlHouse.burstClicksActive = False
        
        suppressKeyPress = True
    
    PThread.kbMsgQueue.put(suppressKeyPress)
    return suppressKeyPress

def keyRelease(KeyboardEvent event) -> bool:
    """
    Description:
        A callback function that is called when a key is released.
    ---
    Parameters:
        `event -> KeyboardEvent`:
            A keyboard event object.
    ---
    Return:
        Always returns False.
    """
    
    if not ctrlHouse.SCROLL and event.KeyID == kbcon.VK_BACKTICK and mgmt.isBacktickTheOnlyModiferPressed:
        mgmt.isBacktickTheOnlyModiferPressed = False
        PThread(target=kbHelper.simulateKeyPress, args=((kbcon.VK_BACKTICK,))).start()
    
    # if not ctrlHouse.modifiers & ctrlHouse.CTRL_SHIFT:
    #     mgmt.mouseVolumeControlSVar.value = False

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
    
    if len(ctrlHouse.pressed_chars) > ctrlHouse.max_alias_length:
        print(" " * len(ctrlHouse.pressed_chars), end="\r")
        
        ctrlHouse.pressed_chars = ""
        
        return True
    
    #+ Printing some relevant information about the pressed key and hardware metrics.
    if not mgmt.silent:
        print(event, f"| Counter={mgmt.counter}") #, f"Thread Count: {threading.active_count()} |", )
        
        print(f"CTRL={    bool(ctrlHouse.modifiers & ctrlHouse.CTRL)    }", end=", ")
        print(f"SHIFT={   bool(ctrlHouse.modifiers & ctrlHouse.SHIFT)   }", end=", ")
        print(f"ALT={     bool(ctrlHouse.modifiers & ctrlHouse.ALT)     }", end=", ")
        print(f"WIN={     bool(ctrlHouse.modifiers & ctrlHouse.WIN)     }", end=", ")
        print(f"FN={      bool(ctrlHouse.modifiers & ctrlHouse.FN)      }", end=", ")
        print(f"BACKTICK={bool(ctrlHouse.modifiers & ctrlHouse.BACKTICK)}", end=" | ")
        
        print(f"NUMLOCK={ctrlHouse.NUMLOCK}")
        print(f"CAPSLOCK={ctrlHouse.CAPITAL}", end=", ")
        print(f"SCROLL={ctrlHouse.SCROLL}", end=", ")
        print("")
        
        # sysHelper.displayCPU_Usage(), print("\n")
        
        # This counter used for debugging purposes only.
        mgmt.counter = (mgmt.counter + 1) % 10000
    
    #+ If the first stored character is not one of the defined prefixes, don't check anything else; this is not a valid word for expansion.
    # The outer conditional supports using multiple prefixes after the first one. The inner ones checks if one of the defined prefixes was pressed or not to stop further processing.
    # if not ctrlHouse.pressed_chars.startswith((":", "!")):
    #     if event.Ascii == kbcon.AS_COLON:
    #         ctrlHouse.pressed_chars = ":"
        
    #     elif event.Ascii == kbcon.AS_EXCLAM:
    #         ctrlHouse.pressed_chars = "!"
        
    #     # If all keys are suppressed, the prefixes ":" and "!" cannot be pressed because shift is suppressed, along with the ";" and "1" keys, preventing the prefixes
    #     # that require shift from being pressed. Thus, we need to check if shift is pressed. Any new prefixes that require shift to be pressed should be added here.
    #     elif (ctrlHouse.modifiers & ctrlHouse.SHIFT):
    #         if event.KeyID == kbcon.VK_SEMICOLON:
    #             ctrlHouse.pressed_chars = ":"
            
    #         elif event.KeyID == kbcon.VK_1:
    #             ctrlHouse.pressed_chars = "!"
        
    #     else:
    #         return True
        
    #     print(ctrlHouse.pressed_chars, end="\r")
    #     return True
    
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
    if event.KeyID == win32con.VK_BACK:
        if len(ctrlHouse.pressed_chars):
            ctrlHouse.pressed_chars = ctrlHouse.pressed_chars[:-1] # Remove the last character
            print(ctrlHouse.pressed_chars + " ", end="\r")
        
        elif ctrlHouse.pressed_chars_backup:
            ctrlHouse.pressed_chars = ctrlHouse.pressed_chars_backup
            ctrlHouse.pressed_chars_backup = ""
            
            kbHelper.undoTextExpansion()
            
            print(ctrlHouse.pressed_chars, end="\r")
        
        return True
    
    #+ If the key is not filtered above, then it is a valid character key.
    ctrlHouse.pressed_chars += event.Key.lower() # chr(event.Ascii).lower()
    print(ctrlHouse.pressed_chars, end="\r")
    
    #+ Check if the pressed characters match any of the defined abbreviations, and replace them accordingly.
    if ctrlHouse.pressed_chars in ctrlHouse.abbreviations or ctrlHouse.pressed_chars in ctrlHouse.non_prefixed_abbreviations:
        print(" " * len(ctrlHouse.pressed_chars), end="\r")
        
        ctrlHouse.pressed_chars_backup = ctrlHouse.pressed_chars
        kbHelper.expandText()
        
        return True
    
    ### Executing some operations based on the pressed characters. ###
    #+ Opening a file or a directory.
    if ctrlHouse.pressed_chars in ctrlHouse.locations:
        print(" " * len(ctrlHouse.pressed_chars), end="\r")
        
        kbHelper.openLocation()
    
    elif ctrlHouse.pressed_chars == ">cls":
        print("    ", end="\r")
        
        ctrlHouse.pressed_chars = ""
        
        os.system("cls")
        
    
    elif ctrlHouse.pressed_chars == "!bst":
        ctrlHouse.burstClicksActive = True
        kbHelper.simulateBurstClicks()
    
    #+ This is a crude way of opening a file using a specific program (open with).
    # elif ctrlHouse.pressed_chars == ":\\\\":
    #     # PThread(target=kbHelper.simulateKeyPressSequence, args=(((win32con.VK_LMENU, 56), (52, 5), (40, 80), (40, 80), (13, 28)))).start()
    #     print(" " * len(ctrlHouse.pressed_chars), end="\r")
    #     kbHelper.crudeOpenWith(4, 2)
    
    ctrlHouse.pressed_chars_backup = ""
    
    return True

@cython.wraparound(False)
def buttonPress(MouseEvent event) -> bool:
    """The callback function responsible for handling the button press and wheel movement events."""
    
    #TODO: Handle when a mouse button is pressed in multiprocess mode.
    ctrlHouse.pressed_chars = ""
    ctrlHouse.pressed_chars_backup = ""
    
    #TODO: Handle when the terminal logging option is changed in multiprocess mode.
    if not mgmt.silent and not event.Delta:
        print(
            f"LB={(msHouse.buttons & msHouse.LButton)  == msHouse.LButton},",
            f"RB={(msHouse.buttons & msHouse.RButton)  == msHouse.RButton},",
            f"MB={(msHouse.buttons & msHouse.MButton)  == msHouse.MButton},",
            f"X1={(msHouse.buttons & msHouse.X1Button) == msHouse.X1Button},",
            f"X2={(msHouse.buttons & msHouse.X2Button) == msHouse.X2Button}",
            f"| pos=({msHouse.x}, {msHouse.y}) | WDelta={msHouse.delta} {'' if not msHouse.delta else 'H' if msHouse.horizontal else 'V'}\n"
        )
        
        # print(event)
    
    cdef bint suppressInput = False
    
    # if event.Delta and (ctrlHouse.modifiers & ctrlHouse.CTRL_SHIFT) == ctrlHouse.CTRL_SHIFT:
    if event.Delta and mgmt.mouseVolumeControlSVar:
        PThread(target=kbHelper.simulateKeyPress, args=((win32con.VK_VOLUME_DOWN, win32con.VK_VOLUME_UP)[event.Delta > 0],)).start()
        
        suppressInput = True
    
    PThread.msMsgQueue.put(suppressInput)
    
    return suppressInput
