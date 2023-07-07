# cython: embedsignature = True
# cython: language_level = 3str

"""This extension module contains the event listeners responsible for handling keyboard keyPress & keyRelease, and mouse events."""

import win32gui, win32con, winsound
import os, subprocess
from cythonExtensions.explorerHelper import explorerHelper as expHelper
from cythonExtensions.systemHelper import systemHelper   as sysHelper
from cythonExtensions.windowHelper import windowHelper   as winHelper
from cythonExtensions.keyboardHelper import keyboardHelper as kbHelper
from cythonExtensions.mouseHelper import mouseHelper    as msHelper
from cythonExtensions.commonUtils.commonUtils import ControllerHouse as ctrlHouse, WindowHouse as winHouse, PThread, Management as mgmt, KB_Con as kbcon
from cythonExtensions.commonUtils import commonUtils
from cythonExtensions.commonUtils.commonUtils cimport KeyboardEvent, MouseEvent
from cythonExtensions.eventListeners import taskConfigs as taskCfg
import threading

cpdef bint KeyPress(KeyboardEvent event):
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
    
    cdef bint suppressInput = mgmt.supressKbInputs # and not (ctrlHouse.modifiers & ctrlHouse.FN) and
    
    if not ctrlHouse.modifiers:
        #++ Scrolling by simulating mouse scroll events. ('W'*, 'S'*, 'A'*, 'D'*)
        if (ctrlHouse.locks & ctrlHouse.SCROLL) and event.KeyID in (kbcon.VK_W, kbcon.VK_S, kbcon.VK_A, kbcon.VK_D):
            dist = 1
            
            PThread(target=commonUtils.SendMouseScroll, args=[*{kbcon.VK_W: (dist, 1), kbcon.VK_S: (-dist, 1), kbcon.VK_A: (-dist, 0), kbcon.VK_D: (dist, 0)}.get(event.KeyID), taskCfg.WHEEL_SCROLL_DISTANCE]).start()
            
            suppressInput = True
        
        #+ Moving the mouse cursor: [";" | "'" | "/" | "."] + {Alt | Shift}
        elif (ctrlHouse.locks & ctrlHouse.SCROLL) and event.KeyID in (kbcon.VK_SEMICOLON, kbcon.VK_SINGLE_QUOTES, kbcon.VK_SLASH, kbcon.VK_PERIOD):
            dist = taskCfg.MOUSE_MOVEMENT_DISTANCE_BASE
            
            PThread(target=msHelper.moveCursor, args=[*{kbcon.VK_SEMICOLON: (0, -dist), kbcon.VK_SINGLE_QUOTES: (dist, 0), kbcon.VK_SLASH: (0, dist), kbcon.VK_PERIOD: (-dist, 0)}.get(event.KeyID)]).start()
            
            suppressInput = True
        
        #+ Zooming in by simulating mouse scroll events: ScrollLock + ('E'*, 'Q'*)
            # PThread(target=kbHelper.SimulateHotKeyPress, args=[({win32con.VK_CONTROL:29, kbcon.VK_MINUS: kbcon.AS_MINUS}, {win32con.VK_CONTROL:29, kbcon.VK_EQUALS: kbcon.AS_EQUALS})[event.KeyID == kbcon.VK_E]]).start()
        
        #+ Sending mouse click: ScrollLock + ('E'*, 'Q'*)
        elif (ctrlHouse.locks & ctrlHouse.SCROLL) and event.KeyID in (kbcon.VK_E, kbcon.VK_Q):
            
            # Release the previously held mouse button.
            if ctrlHouse.heldMouseBtn:
                PThread(target=msHelper.sendMouseClick, args=[0, 0, ctrlHouse.heldMouseBtn, 3]).start()
            
            # If the same button is requested to be pressed, then its release (above) is enough.
            if (1, 2)[event.KeyID == kbcon.VK_E] != ctrlHouse.heldMouseBtn:
                PThread(target=msHelper.sendMouseClick, args=[0, 0, (1, 2)[event.KeyID == kbcon.VK_E]]).start()
            
            ctrlHouse.heldMouseBtn = 0
            
            suppressInput = True
        
        PThread.kbMsgQueue.put(suppressInput)
        
        return suppressInput
    
    ### Script/System management operations ###
    #+ Restart the script with elevated privileges.
    if (ctrlHouse.modifiers & ctrlHouse.CTRL_ALT_FN_WIN) in (ctrlHouse.CTRL_ALT_WIN, ctrlHouse.CTRL_FN_WIN, ctrlHouse.CTRL_ALT_FN_WIN) and event.KeyID == win32con.VK_ESCAPE:
        PThread(target=sysHelper.StartWithElevatedPrivileges, args=[True, True, win32con.SW_FORCEMINIMIZE]).start()
        suppressInput = True
    
    #+ Exitting the script: [FN | Win] + Esc
    elif (ctrlHouse.modifiers & ctrlHouse.FN_WIN) and event.KeyID == win32con.VK_ESCAPE:
        PThread(target=sysHelper.TerminateScript, kwargs={"graceful": True}).start()
        
        suppressInput = True
    
    #+ Displaying a notification window indicating that the script is still running: [FN | Win] + '?'*
    elif (ctrlHouse.modifiers & ctrlHouse.FN_WIN) and event.KeyID == kbcon.VK_SLASH:
        PThread(target=sysHelper.SendScriptWorkingNotification, args=[False]).start()
        
        suppressInput = True
    
    #+ Putting the device into sleep mode: Ctrl + [FN | Alt] + Win + 'S'*
    elif (ctrlHouse.modifiers & ctrlHouse.CTRL_ALT_FN_WIN) in (ctrlHouse.CTRL_ALT_WIN, ctrlHouse.CTRL_FN_WIN, ctrlHouse.CTRL_ALT_FN_WIN) and event.KeyID == kbcon.VK_S:
        PThread(target=sysHelper.GoToSleep).start()
        
        suppressInput = True
    
    #+ Shutting down the system: Ctrl + [FN | Alt] + Win + 'Q'*
    elif (ctrlHouse.modifiers & ctrlHouse.CTRL_ALT_FN_WIN) in (ctrlHouse.CTRL_ALT_WIN, ctrlHouse.CTRL_FN_WIN, ctrlHouse.CTRL_ALT_FN_WIN) and event.KeyID == kbcon.VK_Q:
        PThread(target=sysHelper.Shutdown, args=[True]).start()
        
        suppressInput = True
    
    #+ Incresing/Decreasing the system volume: Ctrl + Shift + (['+'* | Add], ['-'* | Subtract])
    elif (ctrlHouse.modifiers & ctrlHouse.CTRL_SHIFT) == ctrlHouse.CTRL_SHIFT and event.KeyID in (kbcon.VK_EQUALS, kbcon.VK_MINUS, win32con.VK_ADD, win32con.VK_SUBTRACT):
        PThread(target=kbHelper.SimulateKeyPress, args=[(win32con.VK_VOLUME_DOWN, win32con.VK_VOLUME_UP)[event.KeyID in (kbcon.VK_EQUALS, win32con.VK_ADD)]]).start()
        
        suppressInput = True
    
    #+  Incresing/Decreasing brightness: '`' + [F2 | F3]
    elif (ctrlHouse.modifiers & ctrlHouse.BACKTICK) and event.KeyID in (win32con.VK_F2, win32con.VK_F3):
        PThread(target=sysHelper.ChangeBrightness, args=[event.KeyID == win32con.VK_F3]).start()
        
        suppressInput = True
    
    #+ Clearing console output: # FN + Ctrl + 'C'*
    elif (ctrlHouse.modifiers & ctrlHouse.CTRL_FN) == ctrlHouse.CTRL_FN and event.KeyID == kbcon.VK_C:
        os.system("cls")
        
        suppressInput = True
    
    #+ Toggle the terminal output: ALT + FN + 'S'*
    elif (ctrlHouse.modifiers & ctrlHouse.ALT_FN) == ctrlHouse.ALT_FN and event.KeyID == kbcon.VK_S:
        mgmt.silent ^= 1
        
        if mgmt.silent:
            winsound.PlaySound(r"SFX\no-trespassing-368.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
        else:
            winsound.PlaySound(r"SFX\pedantic-490.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
        
        suppressInput = True
    
    #+ Disable the keyboard keys: ALT + FN + 'D'*
    # Note that the hotkeys and text expainsion operations will still work.
    elif (ctrlHouse.modifiers & ctrlHouse.ALT_FN) == ctrlHouse.ALT_FN and event.KeyID == kbcon.VK_D:
        mgmt.supressKbInputs ^= 1
        
        if mgmt.supressKbInputs:
            winsound.PlaySound(r"SFX\no-trespassing-368.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
        else:
            winsound.PlaySound(r"SFX\pedantic-490.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
        
        suppressInput = True
    
    #+ Suspend the process of the active window: '`' + PAUSE
    elif (ctrlHouse.modifiers & ctrlHouse.BACKTICK) and event.KeyID == win32con.VK_PAUSE:
        PThread(target=sysHelper.suspendProcess).start()
        winsound.PlaySound(r"SFX\no-trespassing-368.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
        suppressInput = True
    
    #+ Resume the suspended process of the active window: ALT + PAUSE
    elif (ctrlHouse.modifiers & ctrlHouse.ALT) and event.KeyID == win32con.VK_PAUSE:
        PThread(target=sysHelper.resumeProcess).start()
        winsound.PlaySound(r"SFX\pedantic-490.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
        suppressInput = True
    
    ### Window management Operations ###
    #+  Incresing/Decreasing the opacity of the active window: '`' + (['+'* | Add], ['-'* | Subtract])
    elif (ctrlHouse.modifiers & ctrlHouse.BACKTICK) and event.KeyID in (kbcon.VK_EQUALS, kbcon.VK_MINUS, win32con.VK_ADD, win32con.VK_SUBTRACT):
        PThread(target=winHelper.ChangeWindowOpacity, args=[0, event.KeyID in (kbcon.VK_EQUALS, win32con.VK_ADD)]).start()
        suppressInput = True
    
    #+ Toggling the `AlwaysOnTop` state for the focused window: # FN + Ctrl + 'A'*
    elif (ctrlHouse.modifiers & ctrlHouse.CTRL_FN) == ctrlHouse.CTRL_FN and event.KeyID == kbcon.VK_A:
        PThread(target=winHelper.AlwaysOnTop).start()
        
        suppressInput = True
    
    #+ Moving the active window: '`' + [↑ | → | ↓ | ←] + {Alt | Shift | Win}
    elif (ctrlHouse.modifiers & ctrlHouse.BACKTICK) and event.KeyID in (win32con.VK_UP, win32con.VK_RIGHT, win32con.VK_DOWN, win32con.VK_LEFT):
        
        if ctrlHouse.modifiers & ctrlHouse.ALT: # If `Alt` is pressed, increase the movement distance.
            dist = taskCfg.WINDOW_MOVEMENT_DISTANCE_LARGE
        elif ctrlHouse.modifiers & ctrlHouse.SHIFT: # If `Shift` is pressed, decrease the movemen distance.
            dist = taskCfg.WINDOW_MOVEMENT_DISTANCE_SMALL
        else:
            dist = taskCfg.WINDOW_MOVEMENT_DISTANCE_MEDIUM
        
        if ctrlHouse.modifiers & ctrlHouse.WIN: # If `Win` is pressed, changing window size instead of moving it.
            PThread(target=winHelper.MoveActiveWindow, kwargs={win32con.VK_UP: {'height': dist}, win32con.VK_RIGHT: {'width': dist}, win32con.VK_DOWN: {'height': -dist}, win32con.VK_LEFT: {'width': -dist}}.get(event.KeyID)).start()
        else:
            PThread(target=winHelper.MoveActiveWindow, kwargs={win32con.VK_UP: {'delta_y': -dist}, win32con.VK_RIGHT: {'delta_x': dist}, win32con.VK_DOWN: {'delta_y': dist}, win32con.VK_LEFT: {'delta_x': -dist}}.get(event.KeyID)).start()
        
        suppressInput = True
    
    #+ Creating a new file: Ctrl + Shift + 'M'*
    elif (ctrlHouse.modifiers & ctrlHouse.CTRL_SHIFT) == ctrlHouse.CTRL_SHIFT and event.KeyID == kbcon.VK_M:
        fg_hwnd = win32gui.GetForegroundWindow()
        
        if fg_hwnd and win32gui.GetClassName(fg_hwnd) in ("CabinetWClass", "WorkerW"):
            PThread(target=expHelper.CreateFile).start()
            
            suppressInput = True
    
    #+ Copying the full path to the selected files in the active explorer/desktop window: Shift + F2
    elif (ctrlHouse.modifiers & ctrlHouse.SHIFT) and event.KeyID == win32con.VK_F2:
        fg_hwnd = win32gui.GetForegroundWindow()
        
        if fg_hwnd and win32gui.GetClassName(fg_hwnd) in ("CabinetWClass", "WorkerW"):
            PThread(target=expHelper.CopySelectedFileNames).start()
            
            suppressInput = True
    
    #+ Merging the selected images from the active explorer window into a PDF file: Ctrl + Shift + 'P'*
    elif (ctrlHouse.modifiers & ctrlHouse.CTRL_SHIFT) == ctrlHouse.CTRL_SHIFT and event.KeyID == kbcon.VK_P:
        fg_hwnd = win32gui.GetForegroundWindow()
        
        if fg_hwnd and win32gui.GetClassName(fg_hwnd) == "CabinetWClass":
            PThread(target=expHelper.ImagesToPDF).start()
            
            suppressInput = True
    
    #+ Converting the selected powerpoint files from the active explorer window into PDF files: '`' + 'P'*
    elif (ctrlHouse.modifiers & ctrlHouse.BACKTICK) and event.KeyID == kbcon.VK_P:
        fg_hwnd = win32gui.GetForegroundWindow()
        
        if fg_hwnd and win32gui.GetClassName(fg_hwnd) == "CabinetWClass":
            PThread(target=expHelper.OfficeFileToPDF, kwargs={"office_application": "Powerpoint"}).start()
        
        suppressInput = True
    
    #+ Converting the selected word files from the active explorer window into PDF files: '`' + 'O'*
    elif (ctrlHouse.modifiers & ctrlHouse.BACKTICK) and event.KeyID == kbcon.VK_O:
        fg_hwnd = win32gui.GetForegroundWindow()
        
        if fg_hwnd and win32gui.GetClassName(fg_hwnd) == "CabinetWClass":
            PThread(target=expHelper.OfficeFileToPDF, kwargs={"office_application": "Word"}).start()
        
        suppressInput = True
    
    #+ Storing the address of the closed windows explorer:  [Ctrl + 'W'* | Alt + F4]
    elif ((ctrlHouse.modifiers & ctrlHouse.CTRL) and event.KeyID == kbcon.VK_W) or ((ctrlHouse.modifiers & ctrlHouse.ALT) and event.KeyID == win32con.VK_F4):
        fg_hwnd = win32gui.GetForegroundWindow()
        
        if fg_hwnd and win32gui.GetClassName(fg_hwnd) == "CabinetWClass":
            PThread(target=winHouse.RememberActiveProcessTitle, args=[fg_hwnd]).start()
    
    #+ Opening closed windows explorer: Ctrl + [FN | Win] + 'T'*
    elif (ctrlHouse.modifiers & ctrlHouse.CTRL_FN_WIN) in (ctrlHouse.CTRL_WIN, ctrlHouse.CTRL_FN, ctrlHouse.CTRL_FN_WIN) and event.KeyID == kbcon.VK_T:
        if winHouse.closedExplorers:
            PThread(target=os.startfile, args=[winHouse.closedExplorers.pop()]).start()
            suppressInput = True
    
    ### Starting other scripts ###
    #+ Starting the image processing script: '`' + '\' + {Shift}
    elif (ctrlHouse.modifiers & ctrlHouse.BACKTICK) and event.KeyID == kbcon.VK_BACKSLASH:
        if ctrlHouse.modifiers & ctrlHouse.SHIFT: # Invert without showing the Image window.
            PThread(target=subprocess.call, args=[("python", "-c", "from cythonExtensions.imageInverter import imageInverter as img_inv; img_inv.BeginImageProcessing(False, False)")]).start()
        
        else: # Show the Image window.
            if not winHelper.GetHandleByTitle("Image Window"):
                winsound.PlaySound(r"C:\Windows\Media\Windows Proximity Notification.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
                
                PThread(target=subprocess.call, args=[("python", "-c", "from cythonExtensions.imageInverter import imageInverter as img_inv; img_inv.BeginImageProcessing(True, False)")]).start()
        
        suppressInput = True
    
    ### Keyboard/Mouse control operations ###
    #+ Scrolling by simulating mouse scroll events. ScrollLock + Alt + ('W'*, 'S'*, 'A'*, 'D'*)
    # Don't use shift as it causes the scrolling to be horizontal.
    elif (ctrlHouse.locks & ctrlHouse.SCROLL) and (ctrlHouse.modifiers & ctrlHouse.ALT) and event.KeyID in (kbcon.VK_W, kbcon.VK_S, kbcon.VK_A, kbcon.VK_D):
        # if ctrlHouse.modifiers & ctrlHouse.ALT:
        dist = taskCfg.SCROLL_DISTANCE_MULTIPLIER
        # else:
        #     dist = 1
        
        PThread(target=commonUtils.SendMouseScroll, args=[*{kbcon.VK_W: (dist, 1), kbcon.VK_S: (-dist, 1), kbcon.VK_A: (-dist, 0), kbcon.VK_D: (dist, 0)}.get(event.KeyID), taskCfg.WHEEL_SCROLL_DISTANCE]).start()
        
        suppressInput = True
    
    #+ Sending mouse holding click: ScrollLock + BACKTICK + ('E'*, 'Q'*)
    elif (ctrlHouse.locks & ctrlHouse.SCROLL) and (ctrlHouse.modifiers & ctrlHouse.BACKTICK) and event.KeyID in (kbcon.VK_E, kbcon.VK_Q):
        # Release the previously held mouse button.
        if ctrlHouse.heldMouseBtn:
            PThread(target=msHelper.sendMouseClick, args=[0, 0, ctrlHouse.heldMouseBtn, 3]).start()
        
        # If the same button is requested to be held again, then its release (above) is enough.
        if (1, 2)[event.KeyID == kbcon.VK_E] == ctrlHouse.heldMouseBtn:
            ctrlHouse.heldMouseBtn = 0
        
        else:
            ctrlHouse.heldMouseBtn = (1, 2)[event.KeyID == kbcon.VK_E]
            PThread(target=msHelper.sendMouseClick, args=[0, 0, ctrlHouse.heldMouseBtn, 2]).start()
        
        suppressInput = True
    
    #+ Zooming in by simulating mouse scroll events: ScrollLock + Ctrl + ('E'*, 'Q'*)
    elif (ctrlHouse.locks & ctrlHouse.SCROLL) and (ctrlHouse.modifiers & ctrlHouse.CTRL) and event.KeyID in (kbcon.VK_E, kbcon.VK_Q):
        dist = taskCfg.ZOOMING_DISTANCE
        
        PThread(target=commonUtils.SendMouseScroll, args=[(-dist, dist)[event.KeyID == kbcon.VK_E], 1, taskCfg.ZOOMING_DISTANCE_MULTIPLIER]).start()
        
        suppressInput = True
    
    #+ Moving the mouse cursor: [Alt | Shift] + [";" | "'" | "/" | "."]
    elif (ctrlHouse.locks & ctrlHouse.SCROLL) and event.KeyID in (kbcon.VK_SEMICOLON, kbcon.VK_SINGLE_QUOTES, kbcon.VK_SLASH, kbcon.VK_PERIOD):
        if ctrlHouse.modifiers & ctrlHouse.ALT: # If `Alt` is pressed, increase the movement distance.
            dist = taskCfg.MOUSE_MOVEMENT_DISTANCE_LARGE
        elif ctrlHouse.modifiers & ctrlHouse.SHIFT: # If `Shift` is pressed, decrease the movement distance.
            dist = taskCfg.MOUSE_MOVEMENT_DISTANCE_SMALL
        else:
            dist = taskCfg.MOUSE_MOVEMENT_DISTANCE_MEDIUM
        
        PThread(target=msHelper.moveCursor, args=[*{kbcon.VK_SEMICOLON: (0, -dist), kbcon.VK_SINGLE_QUOTES: (dist, 0), kbcon.VK_SLASH: (0, dist), kbcon.VK_PERIOD: (-dist, 0)}.get(event.KeyID)]).start()
        
        suppressInput = True
    
    #+ Sending arrow keys to the top MPC-HC window: '`' + ('W'*, 'S'*)
    #+ Playing/Pausing the top MPC-HC window: '`' + Space:
    elif (ctrlHouse.modifiers & ctrlHouse.BACKTICK) and event.KeyID in (kbcon.VK_W, kbcon.VK_S, win32con.VK_SPACE):
        # PThread(target=kbHelper.FindAndSendKeyToWindow, args=["MediaPlayerClassicW", {"W": "right", "S": "left", "Space": "space"}.get(event.Key.capitalize()), keyboard.send]).start()
        PThread(target=kbHelper.FindAndSendKeyToWindow, args=["MediaPlayerClassicW", {"W": win32con.VK_RIGHT, "S": win32con.VK_LEFT, "Space": win32con.VK_SPACE}.get(event.Key.capitalize())]).start()
        
        suppressInput = True
    
    #+ Toggling ScrollLock (useful when the keyboard doesn't have the ScrLck key): [FN | Win] + CapsLock
    elif event.KeyID == win32con.VK_CAPITAL and (ctrlHouse.modifiers & ctrlHouse.FN_WIN):
        winsound.PlaySound(r"SFX\pedantic-490.wav" if not ctrlHouse.locks & ctrlHouse.SCROLL else r"SFX\no-trespassing-368.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
        PThread(target=kbHelper.SimulateKeyPress, args=[win32con.VK_SCROLL, 0x46]).start()
        
        suppressInput = True
    
    PThread.kbMsgQueue.put(suppressInput)
    
    return suppressInput


### Word listening operations ###
cpdef bint textExpansion(KeyboardEvent event):
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
        commonUtils.PrintModifiers()
        commonUtils.PrintLockKeys()
        print("")
        
        # sysHelper.DisplayCPUsage(), print("\n")
        
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
        kbHelper.ExpandText()
    
    ### Executing some operations based on the pressed characters. ###
    #+ Opening a file or a directory.
    elif ctrlHouse.pressed_chars in ctrlHouse.locations:
        print(" " * len(ctrlHouse.pressed_chars), end="\r")
        kbHelper.OpenLocation()
    
    #+ This is a crude way of opening a file using a specific program (open with).
    elif ctrlHouse.pressed_chars == ":\\\\":
        # PThread(target=kbHelper.SimulateKeyPressSequence, args=[((win32con.VK_LMENU, 56), (52, 5), (40, 80), (40, 80), (13, 28))]).start()
        print(" " * len(ctrlHouse.pressed_chars), end="\r")
        kbHelper.CrudeOpenWith(4, 2)
    
    return True


cpdef bint ButtonPress(MouseEvent event):
    """The callback function responsible for handling the `ButtonPress` events."""
    
    if not mgmt.silent and not event.Delta:
        commonUtils.PrintButtons()
        # print(event)
    
    cdef bint suppressInput = False
    
    if event.Delta and (ctrlHouse.modifiers & ctrlHouse.CTRL_SHIFT) == ctrlHouse.CTRL_SHIFT:
        PThread(target=kbHelper.SimulateKeyPress, args=[(win32con.VK_VOLUME_DOWN, win32con.VK_VOLUME_UP)[event.Delta > 0]]).start()
        
        suppressInput = True
    
    PThread.msMsgQueue.put(suppressInput)
    
    return suppressInput
