"""This extension module contains the event handlers for the keyboard and mouse events."""

from cythonExtensions.commonUtils.commonUtils import KB_Con as kbcon, ControllerHouse as ctrlHouse, WindowHouse as winHouse, PThread, Management as mgmt

import win32con, os, subprocess, winsound
from collections import defaultdict
from typing import Callable, Tuple

from cythonExtensions.explorerHelper import explorerHelper as expHelper
from cythonExtensions.systemHelper   import systemHelper   as sysHelper
from cythonExtensions.windowHelper   import windowHelper   as winHelper
from cythonExtensions.keyboardHelper import keyboardHelper as kbHelper
from cythonExtensions.mouseHelper    import mouseHelper    as msHelper
import scriptConfigs as configs


# Task-specific configuration values for the callbacks
WHEEL_SCROLL_DISTANCE = 80
SCROLL_DISTANCE_MULTIPLIER = 3

ZOOMING_DISTANCE = 20
ZOOMING_DISTANCE_MULTIPLIER = 1

MOUSE_MOVEMENT_DISTANCE_SMALL  = 5
MOUSE_MOVEMENT_DISTANCE_BASE   = 20
MOUSE_MOVEMENT_DISTANCE_MEDIUM = 40
MOUSE_MOVEMENT_DISTANCE_LARGE  = 80

WINDOW_MOVEMENT_DISTANCE_SMALL  = 2
WINDOW_MOVEMENT_DISTANCE_MEDIUM = 10
WINDOW_MOVEMENT_DISTANCE_LARGE  = WINDOW_MOVEMENT_DISTANCE_MEDIUM * 2


def openClosedExplorer() -> bool:
    if winHouse.closedExplorers:
        PThread(target=os.startfile, args=(winHouse.closedExplorers.pop(),)).start()
        
        return True
    
    return False


def callImageUtilsScript(withGUI=True) -> bool:
    if withGUI and not winHelper.getHandleByTitle("Image Window"):
        winsound.PlaySound(r"C:\Windows\Media\Windows Proximity Notification.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
        
        PThread(target=subprocess.call, args=(("python", "-c", "from cythonExtensions.imageUtils.imageEditor import ImageEditor; ImageEditor(save_near_module=False).runEditor()"),)).start()
    
    elif not withGUI:
        PThread(target=subprocess.call, args=(("python", "-c", "from cythonExtensions.imageUtils.imageHelper import invertClipboardImage; invertClipboardImage()"),)).start()
    
    return True


def callSendMouseClick(button: int) -> bool:
    # Release the previously held mouse button.
    if ctrlHouse.heldMouseBtn:
        PThread(target=msHelper.sendMouseClick, args=(0, 0, ctrlHouse.heldMouseBtn, 3)).start()
    
    # If the same button is requested to be pressed, then its release (above) is enough, otherwise, send it.
    if (1, 2, 3)[button] != ctrlHouse.heldMouseBtn:
        PThread(target=msHelper.sendMouseClick, args=(0, 0, (1, 2, 3)[button])).start()
    
    ctrlHouse.heldMouseBtn = 0
    
    return True


def callSendMouseHoldingClick(button: int) -> bool:
    # Release the previously held mouse button.
    if ctrlHouse.heldMouseBtn:
        PThread(target=msHelper.sendMouseClick, args=(0, 0, ctrlHouse.heldMouseBtn, 3)).start()
    
    # If the same button is requested to be held again, then its release (above) is enough.
    if (1, 2, 3)[button] == ctrlHouse.heldMouseBtn:
        ctrlHouse.heldMouseBtn = 0
    
    else:
        ctrlHouse.heldMouseBtn = (1, 2, 3)[button]
        PThread(target=msHelper.sendMouseClick, args=(0, 0, ctrlHouse.heldMouseBtn, 2)).start()
    
    return True


PThread.throttle(0.3)
def callUndoRedo(undo=True) -> bool:
    keys = {win32con.VK_CONTROL: 29}
    
    if undo:
        keys[kbcon.VK_Z] = kbcon.SC_Z
    else:
        keys[kbcon.VK_Y] = kbcon.SC_Y
    
    kbHelper.simulateHotKeyPress(keys)
    
    return True


def groupEntries(kbEventHandlers: dict[Tuple[int, int], Tuple[Callable, Tuple]]) -> dict[str, Tuple[Tuple[str, Tuple]]]:
    grouped_entries = defaultdict(list)
    
    con_mappings = {getattr(win32con, name): name for name in dir(win32con) if name.startswith("VK_")}
    ctrlHouse_mappings = {getattr(ctrlHouse, name): name for name in dir(ctrlHouse) if not name.startswith("__") and type(getattr(ctrlHouse, name)) is int and name.isupper() and "MASK" not in name and name not in ["NUMLOCK", "CAPITAL", "SCROLL"]}
    kbcon_mappings = kbcon._value2member_map_
    
    for key, value in kbEventHandlers.items():
        if key == (ctrlHouse.BACKTICK, kbcon.VK_H):
            continue
        
        function_name = value[0].__name__
        grouped_entries[function_name].append(([ctrlHouse_mappings.get(k, None) or con_mappings.get(k, None) or (kbcon_mappings[k]._name_ if k in kbcon_mappings else None) or k for k in key], str(value[1:] if len(value) > 2 else value[1])))
    
    return dict(grouped_entries)


def printWrappedDict(grouped_entries, levels=2, indent='    '):
    grouped_entries_text = '{\n'
    
    if levels == 2:
        for key, value in grouped_entries.items():
            grouped_entries_text += f"{indent}\"{key}\": [\n"
            for v in value:
                grouped_entries_text += f"{2 * indent}{v},\n"
            
            grouped_entries_text = grouped_entries_text.rstrip(',\n') + '\n' + indent + '],\n'
    
    else:
        for key, value in grouped_entries.items():
            grouped_entries_text += f"{indent}\"{key}\": {value},\n"
    
    grouped_entries_text = grouped_entries_text.rstrip(',\n') + '\n}\n'
    print(grouped_entries_text)


def printHotkeysAndTextExpansionTriggers():
    os.system("cls")
    print("Hotkeys:")
    printWrappedDict(groupEntries(kbEventHandlers))
    printWrappedDict(groupEntries(kbEventHandlersWithSCROLL_On))
    printWrappedDict(groupEntries(kbEventHandlersWithExplorerFocus))
    printWrappedDict(configs.LOCATIONS, 1)
    printWrappedDict(configs.ABBREVIATIONS, 1)


# Event handlers, which are called when the specified trigger is detected.
# The key of the dictionary is a tuple, where the first element is the modifier keys mask and the second element is the normal keys mask to be pressed.
# The value of the dictionary is a tuple, where the first element is the function to be called and the second element is the arguments to be passed to the function.
kbEventHandlers: dict[tuple[int, int], tuple[Callable, tuple]] = {
    ### Script/System management operations ###
    #+ Restart the script with elevated privileges: Ctrl + [Fn | Alt] + Win + Esc
    (ctrlHouse.CTRL_ALT_WIN, win32con.VK_ESCAPE): (sysHelper.startWithElevatedPrivileges, (True, True, win32con.SW_FORCEMINIMIZE)),
    (ctrlHouse.CTRL_WIN_FN, win32con.VK_ESCAPE):  (sysHelper.startWithElevatedPrivileges, (True, True, win32con.SW_FORCEMINIMIZE)),
    
    #+ Terminating the script: [Fn | Win] + Esc
    (ctrlHouse.FN, win32con.VK_ESCAPE):  (sysHelper.terminateScript, (True,)),
    (ctrlHouse.WIN, win32con.VK_ESCAPE): (sysHelper.terminateScript, (True,)),
    
    #+ Displaying a notification window indicating that the script is still running: [Fn | Win] + '?'*
    (ctrlHouse.FN, kbcon.VK_SLASH):  (sysHelper.sendScriptWorkingNotification, (False,)),
    (ctrlHouse.WIN, kbcon.VK_SLASH): (sysHelper.sendScriptWorkingNotification, (False,)),
    
    #+ Putting the device into sleep mode: Ctrl + [Fn | Alt] + Win + 'S'*
    (ctrlHouse.CTRL_ALT_WIN, kbcon.VK_S): (sysHelper.goToSleep, ()),
    (ctrlHouse.CTRL_WIN_FN, kbcon.VK_S):  (sysHelper.goToSleep, ()),
    
    #+ Shutting down the system: Ctrl + [Fn | Alt] + Win + 'Q'*
    (ctrlHouse.CTRL_ALT_WIN, kbcon.VK_Q): (sysHelper.shutdown, (True,)),
    (ctrlHouse.CTRL_WIN_FN, kbcon.VK_Q):  (sysHelper.shutdown, (True,)),
    
    #+ Toggling the monitor mode: Ctrl + Alt + Win + (1, 2, 3)
    (ctrlHouse.CTRL_ALT_WIN, kbcon.VK_1): (sysHelper.toggleMonitorMode, (1,)),
    (ctrlHouse.CTRL_ALT_WIN, kbcon.VK_2): (sysHelper.toggleMonitorMode, (2,)),
    (ctrlHouse.CTRL_ALT_WIN, kbcon.VK_3): (sysHelper.toggleMonitorMode, (3,)),
    
    #+ Incresing/Decreasing the system volume: Ctrl + Shift + (['+'* | Add], ['-'* | Subtract])
    (ctrlHouse.CTRL_SHIFT, kbcon.VK_EQUALS): (kbHelper.simulateKeyPress, (win32con.VK_VOLUME_UP,)),
    (ctrlHouse.CTRL_SHIFT, kbcon.VK_MINUS):  (kbHelper.simulateKeyPress, (win32con.VK_VOLUME_DOWN,)),
    (ctrlHouse.CTRL_SHIFT, win32con.VK_ADD): (kbHelper.simulateKeyPress, (win32con.VK_VOLUME_UP,)),
    (ctrlHouse.CTRL_SHIFT, win32con.VK_SUBTRACT): (kbHelper.simulateKeyPress, (win32con.VK_VOLUME_DOWN,)),
    
    #+ Incresing/Decreasing brightness: '`' + (F2, F3)
    (ctrlHouse.BACKTICK, win32con.VK_F2): (sysHelper.changeBrightness, (False,)),
    (ctrlHouse.BACKTICK, win32con.VK_F3): (sysHelper.changeBrightness, (True,)),
    
    #+ Clearing console output: # Fn + Ctrl + 'C'*
    (ctrlHouse.CTRL_FN, kbcon.VK_C): (os.system, ("cls",)),
    
    #+ Toggling the terminal output: ALT + Fn + 'S'*
    (ctrlHouse.ALT_FN, kbcon.VK_S): (mgmt.toggleSilentMode, ()),
    
    #+ Disabling the keyboard keys: ALT + Fn + 'D'*
    (ctrlHouse.ALT_FN, kbcon.VK_D): (mgmt.toggleKeyboardInputs, ()),
    
    # #+ Suspending the process of the active window: '`' + [PAUSE | HOME]
    # (ctrlHouse.BACKTICK, win32con.VK_PAUSE): (sysHelper.suspendProcess, ()),
    # (ctrlHouse.BACKTICK, win32con.VK_HOME):  (sysHelper.suspendProcess, ()),
    
    # #+ Resuming the suspended process of the active window: ALT + [PAUSE | HOME]
    # (ctrlHouse.ALT, win32con.VK_PAUSE): (sysHelper.resumeProcess, ()),
    # (ctrlHouse.ALT, win32con.VK_HOME):  (sysHelper.resumeProcess, ()),
    
    ### Window management Operations ###
    #+ Incresing/Decreasing the opacity of the active window: '`' + (['+'* | Add], ['-'* | Subtract])
    (ctrlHouse.BACKTICK, kbcon.VK_EQUALS): (winHelper.changeWindowOpacity, (0, True)),
    (ctrlHouse.BACKTICK, win32con.VK_ADD): (winHelper.changeWindowOpacity, (0, True)),
    (ctrlHouse.BACKTICK, kbcon.VK_MINUS):  (winHelper.changeWindowOpacity, (0, False)),
    (ctrlHouse.BACKTICK, win32con.VK_SUBTRACT): (winHelper.changeWindowOpacity, (0, False)),
    
    #+ Toggling the `alwaysOnTop` state for the focused window: # Fn + Ctrl + 'A'*
    (ctrlHouse.CTRL_FN, kbcon.VK_A): (winHelper.alwaysOnTop, ()),
    (ctrlHouse.CTRL_WIN, kbcon.VK_A): (winHelper.alwaysOnTop, ()),
    
    #+ Moving the active window: '`' + (↑, →, ↓, ←) + {Alt | Shift}
    (ctrlHouse.SHIFT_BACKTICK, win32con.VK_UP):    (winHelper.moveActiveWindow, (0, 0, -WINDOW_MOVEMENT_DISTANCE_SMALL, 0, 0)),
    (ctrlHouse.SHIFT_BACKTICK, win32con.VK_DOWN):  (winHelper.moveActiveWindow, (0, 0, WINDOW_MOVEMENT_DISTANCE_SMALL, 0, 0)),
    (ctrlHouse.SHIFT_BACKTICK, win32con.VK_RIGHT): (winHelper.moveActiveWindow, (0, WINDOW_MOVEMENT_DISTANCE_SMALL, 0, 0, 0)),
    (ctrlHouse.SHIFT_BACKTICK, win32con.VK_LEFT):  (winHelper.moveActiveWindow, (0, -WINDOW_MOVEMENT_DISTANCE_SMALL, 0, 0, 0)),
    
    (ctrlHouse.ALT_BACKTICK, win32con.VK_UP):    (winHelper.moveActiveWindow, (0, 0, -WINDOW_MOVEMENT_DISTANCE_LARGE, 0, 0)),
    (ctrlHouse.ALT_BACKTICK, win32con.VK_DOWN):  (winHelper.moveActiveWindow, (0, 0, WINDOW_MOVEMENT_DISTANCE_LARGE, 0, 0)),
    (ctrlHouse.ALT_BACKTICK, win32con.VK_RIGHT): (winHelper.moveActiveWindow, (0, WINDOW_MOVEMENT_DISTANCE_LARGE, 0, 0, 0)),
    (ctrlHouse.ALT_BACKTICK, win32con.VK_LEFT):  (winHelper.moveActiveWindow, (0, -WINDOW_MOVEMENT_DISTANCE_LARGE, 0, 0, 0)),
    
    (ctrlHouse.BACKTICK, win32con.VK_UP):    (winHelper.moveActiveWindow, (0, 0, -WINDOW_MOVEMENT_DISTANCE_MEDIUM, 0, 0)),
    (ctrlHouse.BACKTICK, win32con.VK_DOWN):  (winHelper.moveActiveWindow, (0, 0, WINDOW_MOVEMENT_DISTANCE_MEDIUM, 0, 0)),
    (ctrlHouse.BACKTICK, win32con.VK_RIGHT): (winHelper.moveActiveWindow, (0, WINDOW_MOVEMENT_DISTANCE_MEDIUM, 0, 0, 0)),
    (ctrlHouse.BACKTICK, win32con.VK_LEFT):  (winHelper.moveActiveWindow, (0, -WINDOW_MOVEMENT_DISTANCE_MEDIUM, 0, 0, 0)),
    
    #+ Changing the size of the active window: '`' + (↑, →, ↓, ←) + Win
    (ctrlHouse.WIN_BACKTICK, win32con.VK_UP):   (winHelper.moveActiveWindow, (0, 0, 0, 0, -WINDOW_MOVEMENT_DISTANCE_SMALL)),
    (ctrlHouse.WIN_BACKTICK, win32con.VK_DOWN): (winHelper.moveActiveWindow, (0, 0, 0, 0, WINDOW_MOVEMENT_DISTANCE_SMALL)),
    (ctrlHouse.WIN_BACKTICK, win32con.VK_RIGHT):(winHelper.moveActiveWindow, (0, 0, 0, WINDOW_MOVEMENT_DISTANCE_SMALL, 0)),
    (ctrlHouse.WIN_BACKTICK, win32con.VK_LEFT): (winHelper.moveActiveWindow, (0, 0, 0, -WINDOW_MOVEMENT_DISTANCE_SMALL, 0)),
    
    #+ Opening closed windows explorer: Ctrl + [Fn | Win] + 'T'*
    (ctrlHouse.CTRL_WIN, kbcon.VK_T): (openClosedExplorer, ()),
    (ctrlHouse.CTRL_FN, kbcon.VK_T): (openClosedExplorer, ()),
    
    ### Starting other scripts ###
    #+ Starting one of the image processing scripts: '`' + '\' + {Shift | Alt}
    (ctrlHouse.BACKTICK, kbcon.VK_BACKSLASH):       (callImageUtilsScript, (True,)),
    (ctrlHouse.SHIFT_BACKTICK, kbcon.VK_BACKSLASH): (callImageUtilsScript, (False,)),
    
    ### Keyboard/Mouse control operations ###
    #+ Sending arrow keys to the top MPC-HC window: '`' + ('W'*, 'S'*)
    (ctrlHouse.BACKTICK, kbcon.VK_W): (kbHelper.findAndSendKeyToWindow, ("MediaPlayerClassicW", win32con.VK_RIGHT)),
    (ctrlHouse.BACKTICK, kbcon.VK_S): (kbHelper.findAndSendKeyToWindow, ("MediaPlayerClassicW", win32con.VK_LEFT)),
    
    #+ Playing/Pausing the top MPC-HC window: '`' + Space:
    (ctrlHouse.BACKTICK, win32con.VK_SPACE): (kbHelper.findAndSendKeyToWindow, ("MediaPlayerClassicW", win32con.VK_SPACE)),
    
    #+ Toggling ScrollLock (useful when the keyboard doesn't have the ScrLck key): [Fn | Win] + CapsLock
    (ctrlHouse.FN, win32con.VK_CAPITAL):  (lambda: (winsound.PlaySound(r"SFX\pedantic-490.wav" if not ctrlHouse.SCROLL else r"SFX\no-trespassing-368.wav", winsound.SND_FILENAME|winsound.SND_ASYNC), kbHelper.simulateKeyPress(win32con.VK_SCROLL, 0x46)), ()),
    (ctrlHouse.WIN, win32con.VK_CAPITAL): (lambda: (winsound.PlaySound(r"SFX\pedantic-490.wav" if not ctrlHouse.SCROLL else r"SFX\no-trespassing-368.wav", winsound.SND_FILENAME|winsound.SND_ASYNC), kbHelper.simulateKeyPress(win32con.VK_SCROLL, 0x46)), ()),
}
"""Dictionary of keyboard event handlers that ignore the state of the lock keys."""


#+ Test: '`' + 'H'*
kbEventHandlers[(ctrlHouse.BACKTICK, kbcon.VK_H)] = (printHotkeysAndTextExpansionTriggers, ())


kbEventHandlersWithSCROLL_On: dict[tuple[int, int], tuple[Callable, tuple]] = {
    ### Keyboard/Mouse control operations ###
    #+ Scrolling by simulating mouse scroll events. ScrollLock + ('W'*, 'S'*, 'A'*, 'D'*) + {Alt}
    (0, kbcon.VK_W): (msHelper.sendMouseScroll, (1,  1, WHEEL_SCROLL_DISTANCE)),
    (0, kbcon.VK_S): (msHelper.sendMouseScroll, (-1, 1, WHEEL_SCROLL_DISTANCE)),
    (0, kbcon.VK_A): (msHelper.sendMouseScroll, (-1, 0, WHEEL_SCROLL_DISTANCE)),
    (0, kbcon.VK_D): (msHelper.sendMouseScroll, (1,  0, WHEEL_SCROLL_DISTANCE)),
    
    # Don't use shift as it causes the scrolling to be horizontal.
    (ctrlHouse.ALT, kbcon.VK_W): (msHelper.sendMouseScroll, (SCROLL_DISTANCE_MULTIPLIER,  1, WHEEL_SCROLL_DISTANCE)),
    (ctrlHouse.ALT, kbcon.VK_S): (msHelper.sendMouseScroll, (-SCROLL_DISTANCE_MULTIPLIER, 1, WHEEL_SCROLL_DISTANCE)),
    (ctrlHouse.ALT, kbcon.VK_A): (msHelper.sendMouseScroll, (-SCROLL_DISTANCE_MULTIPLIER, 0, WHEEL_SCROLL_DISTANCE)),
    (ctrlHouse.ALT, kbcon.VK_D): (msHelper.sendMouseScroll, (SCROLL_DISTANCE_MULTIPLIER,  0, WHEEL_SCROLL_DISTANCE)),
    
    #+ Zooming in by simulating mouse scroll events: ScrollLock + Ctrl + ('E'*, 'Q'*)
    (ctrlHouse.CTRL, kbcon.VK_E): (msHelper.sendMouseScroll, (ZOOMING_DISTANCE,  1, ZOOMING_DISTANCE_MULTIPLIER)),
    (ctrlHouse.CTRL, kbcon.VK_Q): (msHelper.sendMouseScroll, (-ZOOMING_DISTANCE, 1, ZOOMING_DISTANCE_MULTIPLIER)),
    
    #+ Sending mouse clicks: ScrollLock + ('E'*, 'Q'*, 2)
    (0, kbcon.VK_Q): (callSendMouseClick, (0,)),
    (0, kbcon.VK_E): (callSendMouseClick, (1,)),
    (0, kbcon.VK_2): (callSendMouseClick, (2,)),
    
    #+ Sending mouse holding click: ScrollLock + BACKTICK + ('E'*, 'Q'*, 2)
    (ctrlHouse.BACKTICK, kbcon.VK_Q): (callSendMouseHoldingClick, (0,)),
    (ctrlHouse.BACKTICK, kbcon.VK_E): (callSendMouseHoldingClick, (1,)),
    (ctrlHouse.BACKTICK, kbcon.VK_2): (callSendMouseHoldingClick, (2,)),
    
    #+ Moving the mouse cursor: (";", "'", "/", ".") + {Alt | Shift}
    (0, kbcon.VK_SEMICOLON):     (msHelper.moveCursor, (0, -MOUSE_MOVEMENT_DISTANCE_BASE)),
    (0, kbcon.VK_SINGLE_QUOTES): (msHelper.moveCursor, (MOUSE_MOVEMENT_DISTANCE_BASE, 0)),
    (0, kbcon.VK_SLASH):         (msHelper.moveCursor, (0, MOUSE_MOVEMENT_DISTANCE_BASE)),
    (0, kbcon.VK_PERIOD):        (msHelper.moveCursor, (-MOUSE_MOVEMENT_DISTANCE_BASE, 0)),
    
    (ctrlHouse.ALT, kbcon.VK_SEMICOLON):     (msHelper.moveCursor, (0, -MOUSE_MOVEMENT_DISTANCE_LARGE)),
    (ctrlHouse.ALT, kbcon.VK_SINGLE_QUOTES): (msHelper.moveCursor, (MOUSE_MOVEMENT_DISTANCE_LARGE, 0)),
    (ctrlHouse.ALT, kbcon.VK_SLASH):         (msHelper.moveCursor, (0, MOUSE_MOVEMENT_DISTANCE_LARGE)),
    (ctrlHouse.ALT, kbcon.VK_PERIOD):        (msHelper.moveCursor, (-MOUSE_MOVEMENT_DISTANCE_LARGE, 0)),
    
    (ctrlHouse.SHIFT, kbcon.VK_SEMICOLON):     (msHelper.moveCursor, (0, -MOUSE_MOVEMENT_DISTANCE_SMALL)),
    (ctrlHouse.SHIFT, kbcon.VK_SINGLE_QUOTES): (msHelper.moveCursor, (MOUSE_MOVEMENT_DISTANCE_SMALL, 0)),
    (ctrlHouse.SHIFT, kbcon.VK_SLASH):         (msHelper.moveCursor, (0, MOUSE_MOVEMENT_DISTANCE_SMALL)),
    (ctrlHouse.SHIFT, kbcon.VK_PERIOD):        (msHelper.moveCursor, (-MOUSE_MOVEMENT_DISTANCE_SMALL, 0)),
    
    (0, kbcon.VK_SEMICOLON):     (msHelper.moveCursor, (0, -MOUSE_MOVEMENT_DISTANCE_MEDIUM)),
    (0, kbcon.VK_SINGLE_QUOTES): (msHelper.moveCursor, (MOUSE_MOVEMENT_DISTANCE_MEDIUM, 0)),
    (0, kbcon.VK_SLASH):         (msHelper.moveCursor, (0, MOUSE_MOVEMENT_DISTANCE_MEDIUM)),
    (0, kbcon.VK_PERIOD):        (msHelper.moveCursor, (-MOUSE_MOVEMENT_DISTANCE_MEDIUM, 0)),
    
    #+ Simulating the keyboard hotkey `Ctrl + C`: ScrollLock + BACKTICK
    (ctrlHouse.BACKTICK, kbcon.VK_BACKTICK): (kbHelper.simulateHotKeyPress, ({win32con.VK_CONTROL: 29, kbcon.VK_C: kbcon.SC_C},)),
    
    #+ Simulating the keyboard hotkey `Ctrl + Z`: ScrollLock + F1
    (0, win32con.VK_F1): (callUndoRedo, (True,)), # (keyboard.send, ("ctrl+z",)),
    
    #+ Simulating the keyboard hotkey `Ctrl + Shift + Z`: ScrollLock + F2
    (0, win32con.VK_F2): (callUndoRedo, (False,)), # (keyboard.send, ("ctrl+shift+z",)),
}
"""A dictionary of the keyboard event handlers that requires the SCROLL Lock to be on."""

kbEventHandlersWithExplorerFocus: dict[tuple[int, int], tuple[Callable, tuple, bool, bool]] = {
    #+ Creating a new file: Ctrl + Shift + 'M'*
    (ctrlHouse.CTRL_SHIFT, kbcon.VK_M): (expHelper.createNewFile, (), True, True),
    
    #+ Copying the full path to the selected files in the active explorer/desktop window: Shift + F2
    (ctrlHouse.SHIFT, win32con.VK_F2): (expHelper.copySelectedFileNames, (), True, True),
    
    #+ Converting the selected powerpoint/word files from the active explorer window into PDF files: '`' + ('P'*, 'O'*)
    (ctrlHouse.BACKTICK, kbcon.VK_P): (expHelper.officeFileToPDF, (None, "Powerpoint"), True, False),
    (ctrlHouse.BACKTICK, kbcon.VK_O): (expHelper.officeFileToPDF, (None, "Word"), True, False),
    
    #+ Remember the title of the active explorer window (i.e., the path of the current directory): Ctrl + Shift + 'R'*
    (ctrlHouse.CTRL, kbcon.VK_W):    (winHouse.rememberActiveProcessTitle, (), False, False),
    (ctrlHouse.ALT, win32con.VK_F4): (winHouse.rememberActiveProcessTitle, (), False, False),
    
    #+ Merging the selected images from the active explorer window into a PDF file: Ctrl + Shift + 'P'*
    (ctrlHouse.CTRL_SHIFT, kbcon.VK_P): (subprocess.call, (("python", "-c", "from cythonExtensions.imageUtils.imageHelper import imagesToPDF; imagesToPDF()"),), False, True),
    (ctrlHouse.CTRL_SHIFT_ALT, kbcon.VK_P): (subprocess.call, (("python", "-c", "from cythonExtensions.imageUtils.imageHelper import imagesToPDF; imagesToPDF(2)"),), False, True),
    
    #+ Converting the selected image files from the active explorer window into '.ico' files: Ctrl + Alt + Win + 'I'*
    (ctrlHouse.CTRL_ALT_WIN, kbcon.VK_I): (subprocess.call, (("python", "-c", "from cythonExtensions.imageUtils.imageHelper import iconize; iconize()"),)),
    
    #+ Converting the selected '.mp3' files from the active explorer window into '.wav' files: Ctrl + Alt + Win + 'M'*
    (ctrlHouse.CTRL_ALT_WIN, kbcon.VK_M): (expHelper.genericFileConverter, (None, (".mp3", ), lambda f1, f2: subprocess.call(["ffmpeg", "-loglevel", "error", "-hide_banner", "-nostats",'-i', f1, f2]), "", ".wav"), True, True),
}
""""
A dictionary of the keyboard event handlers that requires the explorer window to have focus.

The value tuple of the handlers dictionary must contain four items and the last two must be boolean values specifying:
    - Whether to check (`True`) or pass (`False`) if the desktop window is in focus.
    - Whether to suppress (`True`) or pass (`False`) the pressed keys.
"""
