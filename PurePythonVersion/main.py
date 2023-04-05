"""This module contains the wiring for the entire script."""

#! First and foremost, make sure that this is the only running instance of the script, otherwise, terminate this one.
# Source: https://www.oreilly.com/library/view/python-cookbook/0596001673/ch06s09.html
from win32event import CreateMutex
from win32api import GetLastError
from winerror import ERROR_ALREADY_EXISTS
import os, winsound

#+ Changing the working directory to where this script is.
os.chdir(os.path.dirname(__file__))

#+ Making sure this is the only running instance of the script by creating a mutex.
#? If another process tried to create a new mutex with the same name, the handle to the existing
 # mutex object is returned, and an error is reported, which can be examined using `GetLastError()`.
handle = CreateMutex(None, 1, 'Macropy')
if GetLastError() == ERROR_ALREADY_EXISTS:
    print("Warning! This script is already running!")
    print("This instance is being terminated...")
    winsound.PlaySound(r"SFX\denied.wav", winsound.SND_FILENAME)
    os._exit(1)

import win32gui, win32con, pythoncom, pyWinhook, subprocess
import explorerHelper   as expHelper
import systemHelper     as sysHelper
import windowHelper     as winHelper
import keyboardHelper   as kbHelper
import threading, sys
from time import sleep, time
from common import ControllerHouse as ctrlHouse, WindowHouse as winHouse, PThread, Management as mgmt, KB_Con as kbcon
from GUI_Helper import TTS_House as ttsHouse
from pynput.keyboard import Key as kbKey

#! If you want to run code continuously in the background: https://stackoverflow.com/questions/9705982/pythonw-exe-or-python-exe
    # pythonw YOUR-FILE.pyw
#! Now the process will run continuously in the background. To stop the process, you can run the command:
    # TASKKILL /F /IM pythonw.exe


def HotkeyPressEvent(event) -> bool:
    """
    Description:
        The callback function responsible for handling hotkey press events.
    ---
    Parameters:
        `event`:
            A keyboard event object.
    ---
    Return:
        `return_the_key -> bool`: Whether to return or suppress the pressed key.
    """
    
    #! Specify whether to return the pressed key or suppress it.
    # Always suppress the pressed key if `FN` is also pressed. `not mgmt.suppress_all_keys` adds support for suppressing all keyboard presses.
    # `or SHIFT` is necessary for enabling text expansion operations while suppressing all keyboard presses.
    return_the_key = not mgmt.suppress_all_keys or (ctrlHouse.modifiers & ctrlHouse.SHIFT != 0) # not (ctrlHouse.modifiers & ctrlHouse.FN) and 
    
    ### Script/System management operations ###
    #+ Exitting the script: [FN | Win] + Esc
    if (ctrlHouse.modifiers & ctrlHouse.FN_WIN) and event.KeyID == win32con.VK_ESCAPE:
        PThread(target=sysHelper.TerminateScript, kwargs={"graceful": True}).start()
        
        return_the_key = False
    
    #+ Displaying a notification window indicating that the script is still running: [FN | Win] + '?'*
    elif (ctrlHouse.modifiers & ctrlHouse.FN_WIN) and event.KeyID == kbcon.VK_SLASH:
        PThread(target=sysHelper.SendScriptWorkingNotification).start()
        
        return_the_key = False
    
    #+ Putting the device into sleep mode: Ctrl + [FN | Alt] + Win + 'S'*
    elif (ctrlHouse.modifiers & ctrlHouse.CTRL_ALT_FN_WIN) in (ctrlHouse.CTRL_ALT_WIN, ctrlHouse.CTRL_FN_WIN, ctrlHouse.CTRL_ALT_FN_WIN) and event.KeyID == kbcon.VK_S:
        PThread(target=sysHelper.GoToSleep).start()
        
        return_the_key = False
    
    #+ Shutting down the system: Ctrl + [FN | Alt] + Win + 'Q'*
    elif (ctrlHouse.modifiers & ctrlHouse.CTRL_ALT_FN_WIN) in (ctrlHouse.CTRL_ALT_WIN, ctrlHouse.CTRL_FN_WIN, ctrlHouse.CTRL_ALT_FN_WIN) and event.KeyID == kbcon.VK_Q:
        PThread(target=sysHelper.Shutdown, args=[True]).start()
        
        return_the_key = False
    
    #+ Zooming in by simulating mouse scroll events: ScrLk_ON + Ctrl + ('E'*, 'Q'*)
    elif (ctrlHouse.locks & ctrlHouse.SCROLL) and (ctrlHouse.modifiers & ctrlHouse.CTRL) and event.KeyID in(kbcon.VK_E, kbcon.VK_Q):
        dist = 1
        PThread(target=ctrlHouse.pynput_mouse.scroll, args=[0, (-dist, dist)[event.KeyID == kbcon.VK_E]]).start() # Up
        
        return_the_key = False
    
    #+ Incresing/Decreasing the system volume: Ctrl + Shift + ['+'* | '-'* | Add | Subtract]
    elif (ctrlHouse.modifiers & ctrlHouse.CTRL_SHIFT) == ctrlHouse.CTRL_SHIFT and event.KeyID in (kbcon.VK_EQUALS, kbcon.VK_MINUS, win32con.VK_ADD, win32con.VK_SUBTRACT):
        # PThread(target=keyboard.send, args=[{win32con.VK_F2: "volume down", win32con.VK_F3: "volume up"}.get(event.KeyID)]).start()
        PThread(target=kbHelper.SimulateKeyPress, args=[(win32con.VK_VOLUME_DOWN, win32con.VK_VOLUME_UP)[event.KeyID in (kbcon.VK_EQUALS, win32con.VK_ADD)]]).start()
        
        return_the_key = False
    
    #+  Incresing/Decreasing brightness: '`' + [F2 | F3]
    elif (ctrlHouse.modifiers & ctrlHouse.BACKTICK) and event.KeyID in (win32con.VK_F2, win32con.VK_F3):
        PThread(target=sysHelper.ChangeBrightness, args=[event.KeyID == win32con.VK_F3]).start()
        
        return_the_key = False
    
    #+ Clearing console output: # FN + Ctrl + 'C'*
    elif (ctrlHouse.modifiers & ctrlHouse.CTRL_FN) == ctrlHouse.CTRL_FN and event.KeyID == kbcon.VK_C:
        PThread(target=os.system, args=['cls']).start()
        
        return_the_key = False
    
    #+ Toggle the terminal output: ALT + FN + 'S'*
    elif (ctrlHouse.modifiers & ctrlHouse.ALT_FN) == ctrlHouse.ALT_FN and event.KeyID == kbcon.VK_S:
        mgmt.silent ^= 1
        if mgmt.silent:
            winsound.PlaySound(r"SFX\no-trespassing-368.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
        else:
            winsound.PlaySound(r"SFX\pedantic-490.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
        
        return_the_key = False
    
    #+ Disable the keyboard keys: ALT + FN + 'D'*
    # Note that the hotkeys and text expainsion operations will still work.
    elif (ctrlHouse.modifiers & ctrlHouse.ALT_FN) == ctrlHouse.ALT_FN and event.KeyID == kbcon.VK_D:
        mgmt.suppress_all_keys ^= 1
        if mgmt.suppress_all_keys:
            winsound.PlaySound(r"SFX\no-trespassing-368.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
        else:
            winsound.PlaySound(r"SFX\pedantic-490.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
        
        return_the_key = False
    
    
    ### Window management Operations ###
    #+  Incresing/Decreasing the opacity of the active window: '`' + ['+'* | '-'*, Add, Subtract]
    elif (ctrlHouse.modifiers & ctrlHouse.BACKTICK) and event.KeyID in (kbcon.VK_EQUALS, kbcon.VK_MINUS, win32con.VK_ADD, win32con.VK_SUBTRACT):
        PThread(target=winHelper.ChangeWindowOpacity, args=[0, event.KeyID in (kbcon.VK_EQUALS, win32con.VK_ADD)]).start()
        
        return_the_key = False
    
    #+ Toggling the `AlwaysOnTop` state for the focused window: # FN + Ctrl + 'A'*
    elif (ctrlHouse.modifiers & ctrlHouse.CTRL_FN) == ctrlHouse.CTRL_FN and event.KeyID == kbcon.VK_A:
        PThread(target=winHelper.AlwaysOnTop).start()
        
        return_the_key = False
    
    #+ Moving the active window: '`' + [↑ | → | ↓ | ←] + {Alt | Shift | Win}
    elif (ctrlHouse.modifiers & ctrlHouse.BACKTICK) and event.KeyID in (win32con.VK_UP, win32con.VK_RIGHT, win32con.VK_DOWN, win32con.VK_LEFT):
        dist = 7
        if ctrlHouse.modifiers & ctrlHouse.ALT: # If `Alt` is pressed, increase the movement distance.
            dist = 15
        elif ctrlHouse.modifiers & ctrlHouse.SHIFT: # If `Shift` is pressed, decrease the movemen distance.
            dist = 2
        
        if ctrlHouse.modifiers & ctrlHouse.WIN: # If `Win` is pressed, changing window size instead of moving it.
            PThread(target=winHelper.MoveActiveWindow, kwargs={win32con.VK_UP: {'height': dist}, win32con.VK_RIGHT: {'width': dist}, win32con.VK_DOWN: {'height': -dist}, win32con.VK_LEFT: {'width': -dist}}.get(event.KeyID)).start()
        else:
            PThread(target=winHelper.MoveActiveWindow, kwargs={win32con.VK_UP: {'delta_y': -dist}, win32con.VK_RIGHT: {'delta_x': dist}, win32con.VK_DOWN: {'delta_y': dist}, win32con.VK_LEFT: {'delta_x': -dist}}.get(event.KeyID)).start()
        
        return_the_key = False
    
    #+ Creating a new file: Ctrl + Shift + 'M'*
    elif (ctrlHouse.modifiers & ctrlHouse.CTRL_SHIFT) == ctrlHouse.CTRL_SHIFT and event.KeyID == kbcon.VK_M:
        if win32gui.GetClassName(win32gui.GetForegroundWindow()) in ("CabinetWClass", "WorkerW"):
            PThread(target=expHelper.CtrlShift_M).start()
            return_the_key = False
    
    #+ Copying the full path to the selected files in the active explorer/desktop window: Shift + F2
    elif (ctrlHouse.modifiers & ctrlHouse.SHIFT) and event.KeyID == win32con.VK_F2:
        if win32gui.GetClassName(win32gui.GetForegroundWindow()) in ("CabinetWClass", "WorkerW"):
            PThread(target=expHelper.CopySelectedFileNames).start()
            
            return_the_key = False
    
    #+ Merging the selected images from the active explorer window into a PDF file: Ctrl + Shift + 'P'*
    elif (ctrlHouse.modifiers & ctrlHouse.CTRL_SHIFT) == ctrlHouse.CTRL_SHIFT and event.KeyID == kbcon.VK_P:
        if win32gui.GetClassName(win32gui.GetForegroundWindow()) == "CabinetWClass":
            PThread(target=expHelper.CtrlShift_P).start()
            return_the_key = False

    #+ Converting the selected powerpoint files from the active explorer window into PDF files: '`' + 'P'*
    elif (ctrlHouse.modifiers & ctrlHouse.BACKTICK) and event.KeyID == kbcon.VK_P:
        if win32gui.GetClassName(win32gui.GetForegroundWindow()) == "CabinetWClass":
            PThread(target=expHelper.OfficeFileToPDF, args=["Powerpoint"]).start()
        
        return_the_key = False
    
    #+ Converting the selected word files from the active explorer window into PDF files: '`' + 'O'*
    elif (ctrlHouse.modifiers & ctrlHouse.BACKTICK) and event.KeyID == kbcon.VK_O:
        if win32gui.GetClassName(win32gui.GetForegroundWindow()) == "CabinetWClass":
            PThread(target=expHelper.OfficeFileToPDF, args=["Word"]).start()
        
        return_the_key = False
    
    #+ Storing the address of the closed windows explorer:  [Ctrl + 'W'* | Alt + F4]
    elif ((ctrlHouse.modifiers & ctrlHouse.CTRL) and event.KeyID == kbcon.VK_W) or ((ctrlHouse.modifiers & ctrlHouse.ALT) and event.KeyID == win32con.VK_F4):
        fg_hwnd = win32gui.GetForegroundWindow()
        
        if win32gui.GetClassName(fg_hwnd) == "CabinetWClass":
            PThread(target=winHouse.RememberActiveProcessTitle, args=[fg_hwnd]).start()
    
    #+ Opening closed windows explorer: Ctrl + FN + 'T'*
    elif (ctrlHouse.modifiers & ctrlHouse.CTRL_FN) == ctrlHouse.CTRL_FN and event.KeyID == kbcon.VK_T:
        if winHouse.closedExplorers:
            PThread(target=os.startfile, args=[winHouse.closedExplorers.pop()]).start()
            return_the_key = False
    
    ### Starting other scripts ###
    #+ Starting the image processing script: '`' + '\' + {Shift}
    elif (ctrlHouse.modifiers & ctrlHouse.BACKTICK) and event.KeyID == kbcon.VK_BACKSLASH:
        if ctrlHouse.modifiers & ctrlHouse.SHIFT: # Invert without showing the Image window.
                PThread(target=lambda: not winHelper.GetHandleByTitle("Image Window") and (subprocess.call(("python", os.path.join(os.path.dirname(__file__), "image_inverter.py"), "1")), winsound.PlaySound(r"SFX\coins-497.wav", winsound.SND_FILENAME|winsound.SND_ASYNC))).start()
        else: # Show the Image window.
            PThread(target=lambda: not winHelper.GetHandleByTitle("Image Window") and (winsound.PlaySound(r"C:\Windows\Media\Windows Proximity Notification.wav", winsound.SND_FILENAME|winsound.SND_ASYNC), subprocess.call(("python", os.path.join(os.path.dirname(__file__), "image_inverter.py"))))).start()
        
        return_the_key = False
    
    #+ Starting the web browser script.
    # elif (ctrlHouse.modifiers & ctrlHouse.BACKTICK) and event.KeyID == kbcon.VK_M: # '`' + 'M'*
    #     PThread(target=lambda: (winsound.PlaySound(r"C:\Windows\Media\Windows Proximity Notification.wav", winsound.SND_FILENAME|winsound.SND_ASYNC), subprocess.call(("python", os.path.join(os.path.split(__file__)[0], "webrowser.py"))))).start()
    
    #+ Pausing the running TTS: 'FN' + Shift + 'R'*
    elif (ctrlHouse.modifiers & ctrlHouse.SHIFT_FN) == ctrlHouse.SHIFT_FN and event.KeyID == kbcon.VK_R:
        winsound.PlaySound(r"C:\Windows\Media\Windows Proximity Notification.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
        if ttsHouse.status == 1:    # If status = Running, then pause.
            ttsHouse.status = 2
            ttsHouse.op_called = False
        elif ttsHouse.status == 2:  # If status = Paused, then continue playing.
            with ttsHouse.condition:
                ttsHouse.condition.notify()
        
        return_the_key = False
    
    #+ Stopping the TTS reader: Alt + FN + 'R'*
    elif (ctrlHouse.modifiers & ctrlHouse.ALT_FN) == ctrlHouse.ALT_FN and event.KeyID == kbcon.VK_R:
        ttsHouse.status = 3
        ttsHouse.op_called = False
        
        return_the_key = False
    
    #+ Reading selected text aloud: FN + 'R'*
    elif (ctrlHouse.modifiers & ctrlHouse.FN) and event.KeyID == kbcon.VK_R:
        PThread(target=ttsHouse.ScheduleSpeak).start()
        
        return_the_key = False
    
    
    ### Keyboard/Mouse control operations ###
    #+ Sending arrow keys to the top MPC-HC window: '`' + ['W'* | 'S'*]
    elif (ctrlHouse.modifiers & ctrlHouse.BACKTICK) and event.KeyID in (kbcon.VK_W, kbcon.VK_S):
        PThread(target=kbHelper.FindAndSendKeysToWindow, args=["MediaPlayerClassicW", {kbcon.VK_W: kbKey.right, kbcon.VK_S: kbKey.left}.get(event.KeyID), ctrlHouse.pynput_kb.press]).start()
        # PThread(target=kbHelper.FindAndSendKeysToWindow, args=["MediaPlayerClassicW", {"W": "right", "S": "left"}.get(event.Key), ctrlhouse.keyboard_send]).start()
        
        return_the_key = False
    
    #+ Playing/Pausing the top MPC-HC window: '`' + Space:
    elif (ctrlHouse.modifiers & ctrlHouse.BACKTICK) and event.KeyID == win32con.VK_SPACE: # '`' + Space:
        PThread(target=kbHelper.FindAndSendKeysToWindow, args=["MediaPlayerClassicW", kbKey.space, ctrlHouse.pynput_kb.press]).start()
        
        return_the_key = False
    
    #+ Toggling ScrLck (useful when the keyboard doesn't have the ScrLck key): FN + Tab
    elif event.KeyID == win32con.VK_TAB and (ctrlHouse.modifiers & ctrlHouse.FN):
        PThread(target=lambda: (kbHelper.SimulateKeyPress(win32con.VK_SCROLL, 0x46),
                                winsound.PlaySound(r"SFX\pedantic-490.wav" if ctrlHouse.locks & ctrlHouse.SCROLL else r"SFX\no-trespassing-368.wav", winsound.SND_FILENAME|winsound.SND_ASYNC))).start()
        
        return_the_key = False
    
    PThread.msgQueue.put(return_the_key)
    return return_the_key

### Word listening operations ###
def ExpanderEvent(event):
    """
    Description:
        The callback function responsible for handling the word expansion events.
    ---
    Parameters:
        `event`:
            A keyboard event object.
    """
    
    #+ If the first stored character is not one of the defined prefixes, don't check anything else; this is not a valid word for expansion.
    if not ctrlHouse.pressed_chars.startswith((":", "!")):
        if event.Ascii == kbcon.AS_COLON:
            ctrlHouse.pressed_chars = ":"
            # print(f"\n{ctrlHouse.pressed_chars}\n\033[F\033[A", end="")
            print(ctrlHouse.pressed_chars, end="\r")
        
        elif event.Ascii == kbcon.AS_EXCLAM:
            ctrlHouse.pressed_chars = "!"
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
    ctrlHouse.pressed_chars += chr(event.Ascii).lower()
    print(ctrlHouse.pressed_chars, end="\r")
    
    #+ Check if the pressed characters match any of the defined abbreviations, and if so, replace them accordingly.
    if ctrlHouse.pressed_chars in ctrlHouse.abbreviations:
        print(" " * len(ctrlHouse.pressed_chars), end="\r")
        PThread(target=kbHelper.ExpandText).start()
    
    ### Executing some operations based on the pressed characters. ###
    #+ Opening a file or a directory.
    elif ctrlHouse.pressed_chars in ctrlHouse.locations:
        print(" " * len(ctrlHouse.pressed_chars), end="\r")
        PThread(target=kbHelper.OpenLocation).start()
    
    #+ This is a crude way of opening a file using a specific program (open with).
    elif ctrlHouse.pressed_chars == ":\\\\":
        # PThread(target=kbHelper.SimulateKeySequencePresses, args=[[(win32con.VK_LMENU, 56), (52, 5), (40, 80), (40, 80), (13, 28)]]).start()
        print(" " * len(ctrlHouse.pressed_chars), end="\r")
        PThread(target=kbHelper.CrudeOpenWith, args=[(("alt+4", ctrlHouse.keyboard_send), (win32con.VK_DOWN, kbcon.SC_DOWN), (win32con.VK_DOWN, kbcon.SC_DOWN), (win32con.VK_RETURN, kbcon.SC_RETURN))]).start()
    
    return True


def KeyPress(event):
    """The main callback function for handling the `KeyPress` events. Used mainly for allowing multiple `KeyPress` callbacks to run simultaneously."""
    
    #! Updating states of lock keys.
    ctrlHouse.UpdateLocks(event)
    
    #! Distinguish between real user input and keyboard input generated by programs/scripts.
    if event.Injected == 16:
        return True # Don't suppress the key.
    
    #! Updating states of modifier keys.
    ctrlHouse.UpdateModifiers_KeyDown(event)
    
    # print(f"KeyPress event. Thread count: {threading.active_count()}")
    
    # Hotkey events require pressing at least one modifier key.
    if ctrlHouse.modifiers:
        #+ For the hotkey events.
        PThread(target=HotkeyPressEvent, args=[event]).start()
    
    #++ Scrolling by simulating mouse scroll events. ['W'*, 'S'*, 'A'*, 'D'*]
    elif (ctrlHouse.locks & ctrlHouse.SCROLL) and event.KeyID in (kbcon.VK_W, kbcon.VK_S, kbcon.VK_A, kbcon.VK_D):
        dist = 3
        PThread(target=ctrlHouse.pynput_mouse.scroll, args=[*{kbcon.VK_W: (0,  dist), kbcon.VK_S: (0, -dist), kbcon.VK_A: (-dist, 0), kbcon.VK_D: (dist, 0)}.get(event.KeyID)]).start() # Up
        
        PThread.msgQueue.put(False)
    
    else:
        # print("\033[F")
        
        # if mgmt.prev_event:
        #     print("\033[F\033[K\033[F\033[K", end="")
        
        # debugHouse.prev_event = 0
        
        # sysHelper.DisplayCPUsage()
        
        PThread.msgQueue.put(True and not mgmt.suppress_all_keys)
    
    #+ Printing some relevant information about the pressed key and hardware metrics.
    if not mgmt.silent:
        PThread(target=lambda event: (print(f"Thread Count: {threading.active_count()} |", f"Asc={event.Ascii}, Key={event.Key}, ID={event.KeyID}, SC={event.ScanCode} | Inj={event.Injected} | Ext={event.Extended} | Trans={event.IsTransition()} | Msg='{event.MessageName}' ({event.Message}) | Alt={event.Alt} | Counter={mgmt.counter}"),
                                    ctrlHouse.PrintModifiers(), ctrlHouse.PrintLockKeys(), print("")), args=[event]).start() # sysHelper.DisplayCPUsage(), print("\n")
        
        # A counter used for debugging purposes only. Counter reset after 10k counts.
        mgmt.counter = (mgmt.counter + 1) % 10000
    
    #+ For the text expansion events.
    PThread(target=ExpanderEvent, args=[event]).start()
    
    #+ Getting the return value from the NormalKeyPress thread to determine whether to return the pressed key or suppress it.
    return PThread.msgQueue.get()


def KeyRelease(event):
    """
    Description:
        The callback function responsible for handling the `KeyRelease` events.
    ---
    Parameters:
        `event`:
            A keyboard event object.
    """
    
    # Updating states of modifier keys.
    ctrlHouse.UpdateModifiers_KeyUp(event)
    
    return True


def main():
    """The main entry for the entire script. Configures and starts the keyboard listeners."""
    
    winsound.PlaySound(r"SFX\achievement-message-tone.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
    print(f"Initializing script...")
    
    #+ Initializing the uncaught exception logger.
    sys.excepthook = mgmt.LogUncaughtExceptions
    
    #+ Initializing the COM library for the main thread. Sets the current thread to be a COM apartment thread: https://stackoverflow.com/questions/21141217/how-to-launch-win32-applications-in-separate-threads-in-python
    #? As long as the Automation object is used in the same thread in which it was created, the COM library is already initialized and you do not need to call the CoInitialize function so this line is redundant.
    pythoncom.CoInitialize()
    
    #+ Initializing Text-to-Speech class container.
    # ttsHouse = TTS_House()
    
    #+ For a valid window size reporting.
    sysHelper.EnableDPI_Awareness()
    
    #+ Scheduling a checker to notify if a process with elevated privileges is active while the script does not have elevated privileges. 
    #? This is necessary because no keyboard events are reported when this scenario happens.
    # if not sysHelper.IsProcessElevated(-1):
        # print("Starting the elevated processes checker...")
        # PThread(target=sysHelper.ScheduleElevatedProcessChecker).start()
    
    #+ Initializing a keyboard hook manager.
    kbHook = pyWinhook.HookManager()
    
    # gui = DebuggingOutputGUI()
    # PThread(target=gui.run).start()
    
    
    #+ Initializing all keyboard listeners.
    print("Initializing keyboard listeners...")
    kbHook.KeyDown = KeyPress
    kbHook.KeyUp = KeyRelease
    kbHook.HookKeyboard()
    
    #+ Starting the program main loop.
    print("Activating keyboard listeners...\n")
    
    # To receive os events, there are two possible methods: 1. Using `PumpMessages()` => Blocking call.
    # pythoncom.PumpMessages()
    
    # 2. Using `PumpWaitingMessages()` => Blocks until an event arrives then returns.
    while not mgmt.terminate_script:
        pythoncom.PumpWaitingMessages()
    
    ##! Reaching this point means that the script is being terminated.
    
    # Wait for a certain number of seconds before forcefully stopping the running threads.
    countdown_start = time()
    
    # Get a list of all running threads
    alive_threads = threading.enumerate()
    
    # Count the number of threads in the list
    still_alive = len(alive_threads)
    
    for thread in alive_threads:
        if thread != threading.main_thread():
            print(f"{still_alive} thread{'s are' if still_alive > 1 else ' is'} still active.", end="\r")
            while True:
                if time() - countdown_start >= 10:
                    print(f"\n{still_alive} threads are still running after 10s wait. The threads will be forcefully terminated...")
                    still_alive = 1
                    break
                
                if thread.is_alive():
                    sleep(0.2)
                else:
                    print(f"{still_alive} threads are still active.", end="\r")
                    break
        
        still_alive -= 1
        
        if not still_alive:
            break
    
    #! Un-initializing the COM library if the main loop is terminated.
    print("Un-initializing COM library in the main thread...")
    pythoncom.CoUninitialize()
    
    #! Terminate the script.
    # sysHelper.TerminateScript()

if __name__ == '__main__':
    # Running the script with profiling enabled or running it without profiling.
    if len(sys.argv) > 1 and sys.argv[1] in ("-p", "--profile", "--prof"):
        print("PROFILING ENABLED.")
        import cProfile, pstats
        
        mgmt.silent = True
        
        with cProfile.Profile() as profile:
            main()
        
        from datetime import datetime as dt
        
        profiling_results = pstats.Stats(profile)
        profiling_results.sort_stats(pstats.SortKey.TIME)
        profiling_results.print_stats()
        
        dump_loc = os.path.join(os.path.dirname(__file__), "dumpfiles")
        os.makedirs(dump_loc, exist_ok=True)
        print(f"Dumping profile to:", end=" ")
        dump_loc = os.path.join(dump_loc, f"{dt.now().strftime('%Y-%m-%d (%I%p-%M-%S)')}.prof")
        print(dump_loc)
        profiling_results.dump_stats(dump_loc)
    else:
        main()
