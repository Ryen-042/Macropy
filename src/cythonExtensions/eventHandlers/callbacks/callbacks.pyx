# cython: embedsignature = True
# cython: language_level = 3str

"""This extension module contains the event handlers for the keyboard and mouse events."""

from cythonExtensions.commonUtils.commonUtils cimport KB_Con as kbcon
from cythonExtensions.commonUtils.commonUtils import ControllerHouse as ctrlHouse, WindowHouse as winHouse, PThread, Management as mgmt, sendMouseScroll

import win32gui, win32con, os, subprocess, winsound

from cythonExtensions.explorerHelper import explorerHelper as expHelper
from cythonExtensions.systemHelper   import systemHelper   as sysHelper
from cythonExtensions.windowHelper   import windowHelper   as winHelper
from cythonExtensions.keyboardHelper import keyboardHelper as kbHelper
from cythonExtensions.mouseHelper    import mouseHelper    as msHelper
from cythonExtensions.eventHandlers import taskConfigs    as taskCfg


def callCreateNewFile() -> bool:
    # In some rare cases, after getting foreground window handle, if the window is destroyed before calling the
    # GetClassName function, then it raises an exception.
    try:
        if win32gui.GetClassName(win32gui.GetForegroundWindow()) in ("CabinetWClass", "WorkerW"):
            PThread(target=expHelper.createNewFile).start()
            
            return True
    
    except Exception as e:
        print(f"Error: {e}\nA problem occurred while retrieving the className of the foreground window.\n")
    
    return False


def callCopyFileNames() -> bool:
    try:
        if win32gui.GetClassName(win32gui.GetForegroundWindow()) in ("CabinetWClass", "WorkerW"):
            PThread(target=expHelper.copySelectedFileNames).start()
            
            return True
    
    except Exception as e:
        print(f"Error: {e}\nA problem occurred while retrieving the className of the foreground window.\n")
    
    return False


def callOfficeToPDF(appName="Word") -> bool:
    try:
        if win32gui.GetClassName(win32gui.GetForegroundWindow()) == "CabinetWClass":
            PThread(target=expHelper.officeFileToPDF, args=(None, appName)).start()
        
        return True
    
    except Exception as e:
        print(f"Error: {e}\nA problem occurred while retrieving the className of the foreground window.\n")
    
    return False


def openClosedExplorer() -> bool:
    if winHouse.closedExplorers:
        PThread(target=os.startfile, args=(winHouse.closedExplorers.pop(),)).start()
        
        return True
    
    return False


def callImageUtilsScript(withGUI=True) -> bool:
    if withGUI and not winHelper.getHandleByTitle("Image Window"):
        winsound.PlaySound(r"C:\Windows\Media\Windows Proximity Notification.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
        
        PThread(target=subprocess.call, args=(("python", "-c", "from cythonExtensions.imageUtils import imageUtils as img_inv; img_inv.runScript(True, False)"),)).start()
    
    elif not withGUI:
        PThread(target=subprocess.call, args=(("python", "-c", "from cythonExtensions.imageUtils import imageUtils as img_inv; img_inv.runScript(False, False)"),)).start()
    
    return True


def callSendMouseClick(rightButton=True) -> bool:
    # Release the previously held mouse button.
    if ctrlHouse.heldMouseBtn:
        PThread(target=msHelper.sendMouseClick, args=(0, 0, ctrlHouse.heldMouseBtn, 3)).start()
    
    # If the same button is requested to be pressed, then its release (above) is enough.
    if (1, 2)[rightButton] != ctrlHouse.heldMouseBtn:
        PThread(target=msHelper.sendMouseClick, args=(0, 0, (1, 2)[rightButton])).start()
    
    ctrlHouse.heldMouseBtn = 0
    
    return True


def callSendMouseHoldingClick(rightButton=True) -> bool:
    # Release the previously held mouse button.
    if ctrlHouse.heldMouseBtn:
        PThread(target=msHelper.sendMouseClick, args=(0, 0, ctrlHouse.heldMouseBtn, 3)).start()
    
    # If the same button is requested to be held again, then its release (above) is enough.
    if (1, 2)[rightButton] == ctrlHouse.heldMouseBtn:
        ctrlHouse.heldMouseBtn = 0
    
    else:
        ctrlHouse.heldMouseBtn = (1, 2)[rightButton]
        PThread(target=msHelper.sendMouseClick, args=(0, 0, ctrlHouse.heldMouseBtn, 2)).start()
    
    return True


kbEventHandlers = {
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
    
    #+ Suspending the process of the active window: '`' + [PAUSE | HOME]
    (ctrlHouse.BACKTICK, win32con.VK_PAUSE): (sysHelper.suspendProcess, ()),
    (ctrlHouse.BACKTICK, win32con.VK_HOME):  (sysHelper.suspendProcess, ()),
    
    #+ Resuming the suspended process of the active window: ALT + [PAUSE | HOME]
    (ctrlHouse.ALT, win32con.VK_PAUSE): (sysHelper.resumeProcess, ()),
    (ctrlHouse.ALT, win32con.VK_HOME):  (sysHelper.resumeProcess, ()),
    
    ### Window management Operations ###
    #+  Incresing/Decreasing the opacity of the active window: '`' + (['+'* | Add], ['-'* | Subtract])
    (ctrlHouse.BACKTICK, kbcon.VK_EQUALS): (winHelper.changeWindowOpacity, (0, True)),
    (ctrlHouse.BACKTICK, win32con.VK_ADD): (winHelper.changeWindowOpacity, (0, True)),
    (ctrlHouse.BACKTICK, kbcon.VK_MINUS):  (winHelper.changeWindowOpacity, (0, False)),
    (ctrlHouse.BACKTICK, win32con.VK_SUBTRACT): (winHelper.changeWindowOpacity, (0, False)),
    
    #+ Toggling the `alwaysOnTop` state for the focused window: # Fn + Ctrl + 'A'*
    (ctrlHouse.CTRL_FN, kbcon.VK_A): (winHelper.alwaysOnTop, ()),
    
    #+ Moving the active window: '`' + (↑, →, ↓, ←) + {Alt | Shift}
    (ctrlHouse.SHIFT_BACKTICK, win32con.VK_UP):    (winHelper.moveActiveWindow, (0, 0, -taskCfg.WINDOW_MOVEMENT_DISTANCE_SMALL, 0, 0)),
    (ctrlHouse.SHIFT_BACKTICK, win32con.VK_DOWN):  (winHelper.moveActiveWindow, (0, 0, taskCfg.WINDOW_MOVEMENT_DISTANCE_SMALL, 0, 0)),
    (ctrlHouse.SHIFT_BACKTICK, win32con.VK_RIGHT): (winHelper.moveActiveWindow, (0, taskCfg.WINDOW_MOVEMENT_DISTANCE_SMALL, 0, 0, 0)),
    (ctrlHouse.SHIFT_BACKTICK, win32con.VK_LEFT):  (winHelper.moveActiveWindow, (0, -taskCfg.WINDOW_MOVEMENT_DISTANCE_SMALL, 0, 0, 0)),
    
    (ctrlHouse.ALT_BACKTICK, win32con.VK_UP):    (winHelper.moveActiveWindow, (0, 0, -taskCfg.WINDOW_MOVEMENT_DISTANCE_LARGE, 0, 0)),
    (ctrlHouse.ALT_BACKTICK, win32con.VK_DOWN):  (winHelper.moveActiveWindow, (0, 0, taskCfg.WINDOW_MOVEMENT_DISTANCE_LARGE, 0, 0)),
    (ctrlHouse.ALT_BACKTICK, win32con.VK_RIGHT): (winHelper.moveActiveWindow, (0, taskCfg.WINDOW_MOVEMENT_DISTANCE_LARGE, 0, 0, 0)),
    (ctrlHouse.ALT_BACKTICK, win32con.VK_LEFT):  (winHelper.moveActiveWindow, (0, -taskCfg.WINDOW_MOVEMENT_DISTANCE_LARGE, 0, 0, 0)),
    
    (ctrlHouse.BACKTICK, win32con.VK_UP):    (winHelper.moveActiveWindow, (0, 0, -taskCfg.WINDOW_MOVEMENT_DISTANCE_MEDIUM, 0, 0)),
    (ctrlHouse.BACKTICK, win32con.VK_DOWN):  (winHelper.moveActiveWindow, (0, 0, taskCfg.WINDOW_MOVEMENT_DISTANCE_MEDIUM, 0, 0)),
    (ctrlHouse.BACKTICK, win32con.VK_RIGHT): (winHelper.moveActiveWindow, (0, taskCfg.WINDOW_MOVEMENT_DISTANCE_MEDIUM, 0, 0, 0)),
    (ctrlHouse.BACKTICK, win32con.VK_LEFT):  (winHelper.moveActiveWindow, (0, -taskCfg.WINDOW_MOVEMENT_DISTANCE_MEDIUM, 0, 0, 0)),
    
    #+ Changing the size of the active window: '`' + (↑, →, ↓, ←) + Win
    (ctrlHouse.WIN_BACKTICK, win32con.VK_UP):   (winHelper.moveActiveWindow, (0, 0, 0, 0, -taskCfg.WINDOW_MOVEMENT_DISTANCE_SMALL)),
    (ctrlHouse.WIN_BACKTICK, win32con.VK_DOWN): (winHelper.moveActiveWindow, (0, 0, 0, 0, taskCfg.WINDOW_MOVEMENT_DISTANCE_SMALL)),
    (ctrlHouse.WIN_BACKTICK, win32con.VK_RIGHT):(winHelper.moveActiveWindow, (0, 0, 0, taskCfg.WINDOW_MOVEMENT_DISTANCE_SMALL, 0)),
    (ctrlHouse.WIN_BACKTICK, win32con.VK_LEFT): (winHelper.moveActiveWindow, (0, 0, 0, -taskCfg.WINDOW_MOVEMENT_DISTANCE_SMALL, 0)),
    
    #+ Creating a new file: Ctrl + Shift + 'M'*
    (ctrlHouse.CTRL_SHIFT, kbcon.VK_M): (callCreateNewFile, ()),
    
    #+ Copying the full path to the selected files in the active explorer/desktop window: Shift + F2
    (ctrlHouse.SHIFT, win32con.VK_F2): (callCopyFileNames, ()),
    
    #+ Merging the selected images from the active explorer window into a PDF file: Ctrl + Shift + 'P'*
    (ctrlHouse.CTRL_SHIFT, kbcon.VK_P): (subprocess.call, (("python", "-c", "from cythonExtensions.imageUtils.imageConverter import imagesToPDF; imagesToPDF()"),)),
    
    #+ Converting the selected powerpoint/word files from the active explorer window into PDF files: '`' + ('P'*, 'O'*)
    (ctrlHouse.BACKTICK, kbcon.VK_P): (callOfficeToPDF, ("Powerpoint",)),
    (ctrlHouse.BACKTICK, kbcon.VK_O): (callOfficeToPDF, ("Word",)),
    
    #+ Opening closed windows explorer: Ctrl + [Fn | Win] + 'T'*
    (ctrlHouse.CTRL_WIN, kbcon.VK_T): (openClosedExplorer, ()),
    (ctrlHouse.CTRL_FN, kbcon.VK_T): (openClosedExplorer, ()),
    
    ### Starting other scripts ###
    #+ Starting one of the image processing scripts: '`' + '\' + {Shift | Alt}
    (ctrlHouse.BACKTICK, kbcon.VK_BACKSLASH):       (callImageUtilsScript, (True,)),
    (ctrlHouse.SHIFT_BACKTICK, kbcon.VK_BACKSLASH): (callImageUtilsScript, (False,)),
    (ctrlHouse.ALT_BACKTICK, kbcon.VK_BACKSLASH):   (subprocess.call, (("python", "-c", "from cythonExtensions.imageUtils.imageConverter import iconize; iconize()"),)),
    
    ### Keyboard/Mouse control operations ###
    #+ Sending arrow keys to the top MPC-HC window: '`' + ('W'*, 'S'*)
    #+ Playing/Pausing the top MPC-HC window: '`' + Space:
    (ctrlHouse.BACKTICK, kbcon.VK_W): (kbHelper.findAndSendKeyToWindow, ("MediaPlayerClassicW", win32con.VK_RIGHT)),
    (ctrlHouse.BACKTICK, kbcon.VK_S): (kbHelper.findAndSendKeyToWindow, ("MediaPlayerClassicW", win32con.VK_LEFT)),
    (ctrlHouse.BACKTICK, win32con.VK_SPACE): (kbHelper.findAndSendKeyToWindow, ("MediaPlayerClassicW", win32con.VK_SPACE)),
    
    #+ Toggling ScrollLock (useful when the keyboard doesn't have the ScrLck key): [Fn | Win] + CapsLock
    (ctrlHouse.FN, win32con.VK_CAPITAL):  (lambda: (winsound.PlaySound(r"SFX\pedantic-490.wav" if not ctrlHouse.locks & ctrlHouse.SCROLL else r"SFX\no-trespassing-368.wav", winsound.SND_FILENAME|winsound.SND_ASYNC), kbHelper.simulateKeyPress(win32con.VK_SCROLL, 0x46)), ()),
    (ctrlHouse.WIN, win32con.VK_CAPITAL): (lambda: (winsound.PlaySound(r"SFX\pedantic-490.wav" if not ctrlHouse.locks & ctrlHouse.SCROLL else r"SFX\no-trespassing-368.wav", winsound.SND_FILENAME|winsound.SND_ASYNC), kbHelper.simulateKeyPress(win32con.VK_SCROLL, 0x46)), ()),
}

kbEventHandlersWithLockKeys = {
    ### Keyboard/Mouse control operations ###
    #+ Scrolling by simulating mouse scroll events. ScrollLock + ('W'*, 'S'*, 'A'*, 'D'*) + {Alt}
    (ctrlHouse.SCROLL, 0, kbcon.VK_W): (sendMouseScroll, (1,  1, taskCfg.WHEEL_SCROLL_DISTANCE)),
    (ctrlHouse.SCROLL, 0, kbcon.VK_S): (sendMouseScroll, (-1, 1, taskCfg.WHEEL_SCROLL_DISTANCE)),
    (ctrlHouse.SCROLL, 0, kbcon.VK_A): (sendMouseScroll, (-1, 0, taskCfg.WHEEL_SCROLL_DISTANCE)),
    (ctrlHouse.SCROLL, 0, kbcon.VK_D): (sendMouseScroll, (1,  0, taskCfg.WHEEL_SCROLL_DISTANCE)),
    
    # Don't use shift as it causes the scrolling to be horizontal.
    (ctrlHouse.SCROLL, ctrlHouse.ALT, kbcon.VK_W): (sendMouseScroll, (taskCfg.SCROLL_DISTANCE_MULTIPLIER,  1, taskCfg.WHEEL_SCROLL_DISTANCE)),
    (ctrlHouse.SCROLL, ctrlHouse.ALT, kbcon.VK_S): (sendMouseScroll, (-taskCfg.SCROLL_DISTANCE_MULTIPLIER, 1, taskCfg.WHEEL_SCROLL_DISTANCE)),
    (ctrlHouse.SCROLL, ctrlHouse.ALT, kbcon.VK_A): (sendMouseScroll, (-taskCfg.SCROLL_DISTANCE_MULTIPLIER, 0, taskCfg.WHEEL_SCROLL_DISTANCE)),
    (ctrlHouse.SCROLL, ctrlHouse.ALT, kbcon.VK_D): (sendMouseScroll, (taskCfg.SCROLL_DISTANCE_MULTIPLIER,  0, taskCfg.WHEEL_SCROLL_DISTANCE)),
    
    #+ Zooming in by simulating mouse scroll events: ScrollLock + Ctrl + ('E'*, 'Q'*)
    (ctrlHouse.SCROLL, ctrlHouse.CTRL, kbcon.VK_E): (sendMouseScroll, (-taskCfg.ZOOMING_DISTANCE, 1, taskCfg.ZOOMING_DISTANCE_MULTIPLIER)),
    (ctrlHouse.SCROLL, ctrlHouse.CTRL, kbcon.VK_Q): (sendMouseScroll, (taskCfg.ZOOMING_DISTANCE,  1, taskCfg.ZOOMING_DISTANCE_MULTIPLIER)),
    
    #+ Sending mouse click: ScrollLock + ('E'*, 'Q'*)
    (ctrlHouse.SCROLL, 0, kbcon.VK_E): (callSendMouseClick, (True,)),
    (ctrlHouse.SCROLL, 0, kbcon.VK_Q): (callSendMouseClick, (False,)),
    
    #+ Sending mouse holding click: ScrollLock + BACKTICK + ('E'*, 'Q'*)
    (ctrlHouse.SCROLL, ctrlHouse.BACKTICK, kbcon.VK_E): (callSendMouseHoldingClick, (True,)),
    (ctrlHouse.SCROLL, ctrlHouse.BACKTICK, kbcon.VK_Q): (callSendMouseHoldingClick, (False,)),
    
    #+ Moving the mouse cursor: (";", "'", "/", ".") + {Alt | Shift}
    (ctrlHouse.SCROLL, 0, kbcon.VK_SEMICOLON):     (msHelper.moveCursor, (0, -taskCfg.MOUSE_MOVEMENT_DISTANCE_BASE)),
    (ctrlHouse.SCROLL, 0, kbcon.VK_SINGLE_QUOTES): (msHelper.moveCursor, (taskCfg.MOUSE_MOVEMENT_DISTANCE_BASE, 0)),
    (ctrlHouse.SCROLL, 0, kbcon.VK_SLASH):         (msHelper.moveCursor, (0, taskCfg.MOUSE_MOVEMENT_DISTANCE_BASE)),
    (ctrlHouse.SCROLL, 0, kbcon.VK_PERIOD):        (msHelper.moveCursor, (-taskCfg.MOUSE_MOVEMENT_DISTANCE_BASE, 0)),
    
    (ctrlHouse.SCROLL, ctrlHouse.ALT, kbcon.VK_SEMICOLON):     (msHelper.moveCursor, (0, -taskCfg.MOUSE_MOVEMENT_DISTANCE_LARGE)),
    (ctrlHouse.SCROLL, ctrlHouse.ALT, kbcon.VK_SINGLE_QUOTES): (msHelper.moveCursor, (taskCfg.MOUSE_MOVEMENT_DISTANCE_LARGE, 0)),
    (ctrlHouse.SCROLL, ctrlHouse.ALT, kbcon.VK_SLASH):         (msHelper.moveCursor, (0, taskCfg.MOUSE_MOVEMENT_DISTANCE_LARGE)),
    (ctrlHouse.SCROLL, ctrlHouse.ALT, kbcon.VK_PERIOD):        (msHelper.moveCursor, (-taskCfg.MOUSE_MOVEMENT_DISTANCE_LARGE, 0)),
    
    (ctrlHouse.SCROLL, ctrlHouse.SHIFT, kbcon.VK_SEMICOLON):     (msHelper.moveCursor, (0, -taskCfg.MOUSE_MOVEMENT_DISTANCE_SMALL)),
    (ctrlHouse.SCROLL, ctrlHouse.SHIFT, kbcon.VK_SINGLE_QUOTES): (msHelper.moveCursor, (taskCfg.MOUSE_MOVEMENT_DISTANCE_SMALL, 0)),
    (ctrlHouse.SCROLL, ctrlHouse.SHIFT, kbcon.VK_SLASH):         (msHelper.moveCursor, (0, taskCfg.MOUSE_MOVEMENT_DISTANCE_SMALL)),
    (ctrlHouse.SCROLL, ctrlHouse.SHIFT, kbcon.VK_PERIOD):        (msHelper.moveCursor, (-taskCfg.MOUSE_MOVEMENT_DISTANCE_SMALL, 0)),
    
    (ctrlHouse.SCROLL, 0, kbcon.VK_SEMICOLON):     (msHelper.moveCursor, (0, -taskCfg.MOUSE_MOVEMENT_DISTANCE_MEDIUM)),
    (ctrlHouse.SCROLL, 0, kbcon.VK_SINGLE_QUOTES): (msHelper.moveCursor, (taskCfg.MOUSE_MOVEMENT_DISTANCE_MEDIUM, 0)),
    (ctrlHouse.SCROLL, 0, kbcon.VK_SLASH):         (msHelper.moveCursor, (0, taskCfg.MOUSE_MOVEMENT_DISTANCE_MEDIUM)),
    (ctrlHouse.SCROLL, 0, kbcon.VK_PERIOD):        (msHelper.moveCursor, (-taskCfg.MOUSE_MOVEMENT_DISTANCE_MEDIUM, 0)),
}