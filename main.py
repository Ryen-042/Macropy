"""This module contains the wiring for the entire script."""

#! First and foremost, make sure that this is the only running instance of the script, otherwise, terminate this one.
# Source: https://www.oreilly.com/library/view/python-cookbook/0596001673/ch06s09.html
from win32event import CreateMutex
from win32api import GetLastError
from winerror import ERROR_ALREADY_EXISTS
import os, winsound

#+ Change the working directory to where this script is.
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
from common import ControllerHouse as ctrlHouse, PThread, DebuggingHouse, KB_Con as kbcon
# from GUI_Helper import TTS_House
from pynput.keyboard import Key as kbKey
from win11toast import toast # , notify, update_progress

#! If you want to run code continuously in the background: https://stackoverflow.com/questions/9705982/pythonw-exe-or-python-exe
    # pythonw YOUR-FILE.pyw
#! Now the process will run continuously in the background. To stop the process, you can run the command:
    # TASKKILL /F /IM pythonw.exe


def NormalKeyPress(event):
    """
    Description:
        The callback function responsible for handling normal key press events.
    ---
    Parameters:
        `event`:
            A keyboard event object.
    """
    
    # A counter used for debugging purposes only. Counter reset after 10k counts.
    DebuggingHouse.counter = (DebuggingHouse.counter + 1) % 10000
    
    #! Used for determining whether to return the pressed key or suppress it.
    suppress_key = True
    
    #! Updating states of modifier keys.
    ctrlHouse.UpdateModifierKeys(event, "|")
    
    ### Script/System management operations ###
    #+ Exitting the script.
    if ctrlHouse.modifiers.FN and event.KeyID == win32con.VK_ESCAPE: # FN + 'Esc'
        PThread(target=sysHelper.TerminateScript).start()
        PThread.outputQueue.put(False)
        return False
    
    #+ Displaying a notification window indicating that the script is still running.
    elif ctrlHouse.modifiers.FN and event.KeyID == kbcon.VK_SLASH:  # FN + (['/', '?'] => 'Oem_2')
            buttons = ( {'activationType': 'protocol', 'arguments': "0:", 'content': 'Exit Script', "hint-buttonStyle": "Critical"},
                        {'activationType': 'protocol', 'arguments': os.path.split(__file__)[0], 'content': 'Open Script Folder'},
                        {'activationType': 'protocol', 'arguments': "1:", 'content': 'Nice Work', "hint-buttonStyle": "Success"})
            
            # notify() doesn't work properly here. Use toast() inside a thread instead.
            PThread(target = lambda: toast('Script is Running.', 'The script is running in the background.', buttons=buttons,
                                            on_click = lambda args: sysHelper.TerminateScript(args["arguments"][0] == "0"),
                                            icon = {"src": r"E:\UnKnown\Programming\Python\Year 22\Automation\Macropy\pyCom\keyboard.png",
                                                    'placement': 'appLogoOverride'},
                                            image = {'src': r"E:\UnKnown\Programming\Python\Year 22\Automation\Macropy\pyCom\keyboard (0.5).png",
                                                     'placement': 'hero'},
                                            audio={'silent': 'true'})).start()
    
    #+ Putting the device into sleep mode.
    elif ctrlHouse.modifiers.WIN and ctrlHouse.modifiers.FN and ctrlHouse.modifiers.CTRL and event.KeyID == kbcon.VK_S: # Win + FN + Ctrl + ['s', 'S']
        PThread(target=sysHelper.GoToSleep).start()
    
    #+ Shutdown the system.
    elif ctrlHouse.modifiers.WIN and ctrlHouse.modifiers.FN and ctrlHouse.modifiers.CTRL and event.KeyID == kbcon.VK_Q: # Win + FN + Ctrl + ['q', 'Q']
        PThread(target=sysHelper.Shutdown).start()
    
    #+ Incresing/Decreasing the system volume.
    elif ctrlHouse.modifiers.CTRL and ctrlHouse.modifiers.SHIFT and event.KeyID in (kbcon.VK_EQUALS, kbcon.VK_MINUS): # Ctrl + Shift + (['=', '+'] Or  ['-', '_'])
        PThread(target=kbHelper.SimulateKeyPress, args=[{kbcon.VK_EQUALS: win32con.VK_VOLUME_UP, kbcon.VK_MINUS: win32con.VK_VOLUME_DOWN}.get(event.KeyID)]).start()
    
    #+  Incresing/Decreasing brightness.
    elif ctrlHouse.modifiers.BACKTICK and event.KeyID in (win32con.VK_F2, win32con.VK_F3):  # '`' + ('F2' Or  'F3')
        PThread(target=sysHelper.ChangeBrightness, args=[event.KeyID == win32con.VK_F3]).start()
    
    #+  Incresing/Decreasing the active window opacity.
    elif ctrlHouse.modifiers.BACKTICK and event.KeyID in (kbcon.VK_EQUALS, kbcon.VK_MINUS): # '`' + (['=', '+'] Or  ['-', '_'])
        PThread(target=sysHelper.ChangeWindowOpacity, args=[0, event.KeyID == kbcon.VK_EQUALS]).start()
    
    #+ Toggling the `AlwaysOnTop` state for the focused window.
    elif ctrlHouse.modifiers.FN and ctrlHouse.modifiers.CTRL and event.KeyID == kbcon.VK_A: # FN + Ctrl + ['a', 'A']
        PThread(target=winHelper.AlwaysOnTop).start()
    
    #+ Moving the active window [Up, Right, Down, Left].
    elif ctrlHouse.modifiers.BACKTICK and event.KeyID in (win32con.VK_UP, win32con.VK_RIGHT, win32con.VK_DOWN, win32con.VK_LEFT): # '`' + (↑ Or → Or ↓ Or ←)
        dist = 7
        if ctrlHouse.modifiers.ALT: # If `Alt` is pressed, increase the movement distance.
            dist = 15
        elif ctrlHouse.modifiers.SHIFT: # If `Shift` is pressed, decrease the movemen distance.
            dist = 2
        
        if ctrlHouse.modifiers.FN: # Changing window size
            PThread(target=winHelper.MoveActiveWindow, kwargs={win32con.VK_UP: {'height': dist}, win32con.VK_RIGHT: {'width': dist}, win32con.VK_DOWN: {'height': -dist}, win32con.VK_LEFT: {'width': -dist}}.get(event.KeyID)).start()
        else:
            PThread(target=winHelper.MoveActiveWindow, kwargs={win32con.VK_UP: {'delta_y': -dist}, win32con.VK_RIGHT: {'delta_x': dist}, win32con.VK_DOWN: {'delta_y': dist}, win32con.VK_LEFT: {'delta_x': -dist}}.get(event.KeyID)).start()
    
    #+ Clearing console output.
    elif ctrlHouse.modifiers.FN and ctrlHouse.modifiers.CTRL and event.KeyID == kbcon.VK_C: # FN + Ctrl + ['c', 'C']
        PThread(target=os.system, args=['cls']).start()
    
    #+ Toggle the terminal output.
    elif ctrlHouse.modifiers.FN and ctrlHouse.modifiers.ALT and event.KeyID == kbcon.VK_S: # FN + ALT + ['s', 'S']
        DebuggingHouse.silent ^= 1
        if DebuggingHouse.silent:
            winsound.PlaySound(r"SFX\no-trespassing-368.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
        else:
            winsound.PlaySound(r"SFX\pedantic-490.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
    
    ### Automation operations ###
    #+ Creating a new file.
    elif ctrlHouse.modifiers.CTRL and ctrlHouse.modifiers.SHIFT and event.KeyID == kbcon.VK_M:    # Ctrl + Shift + ['m', 'M']
        if win32gui.GetClassName(win32gui.GetForegroundWindow()) in ("CabinetWClass", "WorkerW"):
            PThread(target=expHelper.CtrlShift_M).start()
        else:
            suppress_key = False
    
    #+ Copying the full path to the selected file in the active explorer/desktop window.
    elif ctrlHouse.modifiers.SHIFT and event.KeyID == win32con.VK_F2: # Shift + 'F2'
        if win32gui.GetClassName(win32gui.GetForegroundWindow()) in ("CabinetWClass", "WorkerW"):
            PThread(target=expHelper.CopySelectedFileNames).start()
        else:
            suppress_key = False
    
    #+ Merging the selected images from the active explorer window into a PDF file.
    elif ctrlHouse.modifiers.CTRL and ctrlHouse.modifiers.SHIFT and event.KeyID == kbcon.VK_P: # Ctrl + Shift + ['p', 'P']
        if win32gui.GetClassName(win32gui.GetForegroundWindow()) == "CabinetWClass":
            PThread(target=expHelper.CtrlShift_P).start()
        else:
            suppress_key = False
    
    #+ Converting the selected powerpoint files from the active explorer window into PDF files.
    elif ctrlHouse.modifiers.BACKTICK and event.KeyID == kbcon.VK_P: # '`' + ['p', 'P']
        if win32gui.GetClassName(win32gui.GetForegroundWindow()) == "CabinetWClass":
            PThread(target=expHelper.OfficeFileToPDF, args=["Powerpoint"]).start()
        else:
            suppress_key = False
    
    #+ Converting the selected word files from the active explorer window into PDF files.
    elif ctrlHouse.modifiers.BACKTICK and event.KeyID == kbcon.VK_O: # '`' + ['o', 'O']
        if win32gui.GetClassName(win32gui.GetForegroundWindow()) == "CabinetWClass":
            PThread(target=expHelper.OfficeFileToPDF, args=["Word"]).start()
        else:
            suppress_key = False
    
    #+ Pausing the running TTS.
    # elif ctrlHouse.modifiers.FN and ctrlHouse.modifiers.SHIFT and event.KeyID == kbcon.VK_R: # 'FN' + Shift + ['r', 'R']
    #     winsound.PlaySound(r"C:\Windows\Media\Windows Proximity Notification.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
    #     if ttsHouse.status == 1:
    #         ttsHouse.status = 2
    #         ttsHouse.op_called = False
    #     elif ttsHouse.status == 2:
    #         with ttsHouse.condition:
    #             ttsHouse.condition.notify()
    
    # elif ctrlHouse.modifiers.FN and ctrlHouse.modifiers.ALT and event.KeyID == kbcon.VK_R: # 'FN' + Alt + ['r', 'R']
    #     ttsHouse.status = 3
    #     ttsHouse.op_called = False
    
    #+ Reading selected text aloud.
    # elif ctrlHouse.modifiers.FN and event.KeyID == kbcon.VK_R: # 'FN' + ['r', 'R']
    #     PThread(target=ttsHouse.ScheduleSpeak).start()
    
    #+ Starting the image processing script.
    elif ctrlHouse.modifiers.BACKTICK and event.KeyID == kbcon.VK_BACKSLASH: # '`' + ['\' ("Oem_5")]
        PThread(target=lambda: not winHelper.GetHandleByTitle("Image Window") and (winsound.PlaySound(r"C:\Windows\Media\Windows Proximity Notification.wav", winsound.SND_FILENAME|winsound.SND_ASYNC), subprocess.call(("python", os.path.join(os.path.split(__file__)[0], "image_inverter.py"))))).start()
    
    #+ Starting the web browser script.
    elif ctrlHouse.modifiers.BACKTICK and event.KeyID == kbcon.VK_M: # '`' + ['m', 'M']
        PThread(target=lambda: (winsound.PlaySound(r"C:\Windows\Media\Windows Proximity Notification.wav", winsound.SND_FILENAME|winsound.SND_ASYNC), subprocess.call(("python", os.path.join(os.path.split(__file__)[0], "webrowser.py"))))).start()
    
    ### Keyboard/Mouse control operations ###
    #+ Sending arrow keys to the top MPC-HC window.
    elif ctrlHouse.modifiers.BACKTICK and event.KeyID in (kbcon.VK_W, kbcon.VK_S): # '`' + (['w', 'W'] Or ['s', 'S']):
        PThread(target=kbHelper.FindAndSendKeysToWindow, args=["MediaPlayerClassicW", {kbcon.VK_W: kbKey.right, kbcon.VK_S: kbKey.left}.get(event.KeyID), ctrlHouse.pynput_kb.press]).start()
    
    #+ Playing/Pausing the top MPC-HC window.
    elif ctrlHouse.modifiers.BACKTICK and event.KeyID == win32con.VK_SPACE: # '`' + Space:
        PThread(target=kbHelper.FindAndSendKeysToWindow, args=["MediaPlayerClassicW", kbKey.space, ctrlHouse.pynput_kb.press]).start()
    
    #+ Toggling ScrLck.
    elif ctrlHouse.modifiers.FN and event.KeyID == win32con.VK_TAB: # FN + 'Tab`
        PThread(target=kbHelper.SimulateKeyPress, args=[win32con.VK_SCROLL, 0x46]).start()
    
    #+ Scrolling by simulating mouse scroll events. ["W", "S", "A", "D"]
    elif ctrlHouse.locks.SCROLL and event.KeyID in [kbcon.VK_W, kbcon.VK_S, kbcon.VK_A, kbcon.VK_D]: # ScrLk_ON + (['w', 'W'] Or ['s', 'S'] Or ['a', 'A'] Or ['d', 'D']):
        dist = 1
        if ctrlHouse.modifiers.ALT: # If `Alt` is pressed, increase the schrolling distance - `Shift` here doesn't work :(.
            dist = 3
        PThread(target=ctrlHouse.pynput_mouse.scroll, args=[*{kbcon.VK_W: (0,  dist), kbcon.VK_S: (0, -dist), kbcon.VK_A: (-dist, 0), kbcon.VK_D: (dist, 0)}.get(event.KeyID)]).start() # Up
    
    #! If no condition was matched, then just return the typed key.
    else:
        suppress_key = ctrlHouse.modifiers.FN
    
    PThread.outputQueue.put(not suppress_key)
    return not suppress_key

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
    ctrlHouse.UpdateModifierKeys(event, "&")
    
    return True

### Word listening operations ###
def ExpanderKeyPress(event):
    """
    Description:
        The callback function responsible for handling the word expansion events.
    ---
    Parameters:
        `event`:
            A keyboard event object.
    """
    
    #+ If the first stored character is not one of the defined prefixes, don't check anything else; this is not a valid word for expansion.
    if not ctrlHouse.pressed_chars.startswith(":"):
        if event.Ascii == kbcon.AS_COLON:
            ctrlHouse.pressed_chars = ":"
            print(ctrlHouse.pressed_chars)
        return True
    
    #- Most of the function and modifier keys like ctrl, shift, end, home, ect... have Ascii values equal to zero.
    #- The key Ascii is a non-zero value, but the key is not a normal character. These keys are [Space, Enter, Tab, ESC].
    #? If one of the above is true, then return the key without checking anything else. Note that `Enter` key is returned as `\r`.
    if event.Ascii == 0 or event.KeyID in [win32con.VK_SPACE, win32con.VK_RETURN, win32con.VK_TAB, win32con.VK_ESCAPE]:
        print(f"Ascii: {event.Ascii} | Key: {event.Key}")
        
        # To allow for `ctrl` + character keys, you need to check here for them individually. The same `event.Key` value is returned but `event.Ascii` is 0.
        ctrlHouse.pressed_chars = "" # Reset stored pressed keys.
        return True
    
    #+ Pop characters if `backspace` is pressed.
    if event.KeyID == win32con.VK_BACK and len(ctrlHouse.pressed_chars):
        ctrlHouse.pressed_chars = ctrlHouse.pressed_chars[:-1] # Remove the last character
        print(ctrlHouse.pressed_chars)
        return True
    
    #+ If key is not filtered above, then it is a valid character key.
    ctrlHouse.pressed_chars += chr(event.Ascii)
    print(ctrlHouse.pressed_chars)
    
    #+ Check if the pressed characters match any of the defined abbreviations, and if so, replace them accordingly.
    if ctrlHouse.pressed_chars in ctrlHouse.abbreviations:
        PThread(target=kbHelper.ExpandText, args=[ctrlHouse.abbreviations.get(ctrlHouse.pressed_chars), len(ctrlHouse.pressed_chars)]).start()
    
    ### Executing some operations based on the pressed characters. ###
    #+ This is a crude way of opening a file using a specific program (open with).
    elif ctrlHouse.pressed_chars == ":\\\\":
        PThread(target=kbHelper.SimulateKeyPressSequence, args=[(("alt+4", ctrlHouse.keyboard_send), (win32con.VK_DOWN, kbcon.SC_DOWN), (win32con.VK_DOWN, kbcon.SC_DOWN), (win32con.VK_RETURN, kbcon.SC_RETURN))]).start()
    
    return True


def KeyPress(event):
    """The main callback function for handling the `KeyPress` events. Used mainly for allowing multiple `KeyPress` callbacks to run simultaneously."""
    
    #! Updating states of lock keys.
    ctrlHouse.UpdateLockKeys(event)
    
    #! Distinguish between real user input and keyboard input generated by programmes/scripts.
    if event.Injected == 16:
        return True
    
    # print(f"KeyPress event. Thread count: {threading.active_count()}")
    
    #+ Running multiple keydown events simultaneously.
    PThread(target=NormalKeyPress, args=[event]).start()
    PThread(target=ExpanderKeyPress, args=[event]).start()
    
    #+ Printing some relevant information about the pressed key and hardware metrics.
    PThread(target=lambda event: not DebuggingHouse.silent and (print(f"Thread Count: {threading.active_count()} |", event.Ascii,  event.Key, event.KeyID, event.ScanCode, "|", event.Injected, event.IsInjected(), "|", event.Extended, event.IsExtended(), "|", event.IsTransition(), event.Transition, "|", event.Message, f"'{event.MessageName}'", "|", event.Alt, event.IsAlt(), DebuggingHouse.counter), print(ctrlHouse.modifiers, ctrlHouse.locks, end="\n\n\n"),
                                  sysHelper.DisplayCPUsage(), print("\n")), args=[event]).start()
    
    #+ Getting the return value from the NormalKeyPress thread to determine whether to retrun the pressed key or suppress it.
    return PThread.outputQueue.get()


def main():
    """The main entry for the entire script. Configures and starts the keyboard listeners."""
    
    winsound.PlaySound(r"SFX\achievement-message-tone.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
    print(f"Initializing script...")
    
    #+ Initializing the uncaught exception logger.
    sys.excepthook = DebuggingHouse.LogUncaughtExceptions
    
    #+ Initializing the COM library for the main thread. Sets the current thread to be a COM apartment thread: https://stackoverflow.com/questions/21141217/how-to-launch-win32-applications-in-separate-threads-in-python
    #? As long as the Automation object is used in the same thread in which it was created, the COM library is already initialized and you do not need to call the CoInitialize function so this line is redundant.
    pythoncom.CoInitialize()
    
    #+ Initializing necessary data containers.
    # ttsHouse = TTS_House()
    
    #+ For a valid window size reporting.
    sysHelper.EnableDPI_Awareness()
    
    #+ Scheduling a checker to notify if a process with elevated privileges is active while the script does not have elevated privileges. 
    #? This is necessary because no keyboard events are reported when this scenario happens.
    if not sysHelper.IsProcessElevated(-1):
        print("Starting the elevated processes checker...")
        PThread(target=sysHelper.ScheduleElevatedProcessChecker).start()
    
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
    print("Activating keyboard listeners...")
    
    # To receive os events, there is two possible methods: 1. Using `PumpMessages()` => Blocking call.
    # pythoncom.PumpMessages()
    
    # 2. Using `PumpWaitingMessages()` => Blocks until an event arrive then returns.
    while not DebuggingHouse.terminate_script:
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
            while True:
                if time() - countdown_start >= 10:
                    print(f"{still_alive} threads are still running after 10s wait. The threads will be forcefully terminated...")
                    still_alive = 1
                    break
                
                if thread.is_alive():
                    sleep(0.2)
                else:
                    break
        still_alive -= 1
        
        if not still_alive:
            break
    
    #! Un-initializing the COM library if the main loop is terminated.
    print("Un-initializing COM library...")
    pythoncom.CoUninitialize()
    
    #! Terminate the script.
    sysHelper.TerminateScript()

if __name__ == '__main__':
    # Running the script with profiling enabled or running it without profiling.
    if len(sys.argv) > 1 and sys.argv[1] in ("-p", "--profile", "--prof"):
        import cProfile, pstats
        
        DebuggingHouse.silent = True
        
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
