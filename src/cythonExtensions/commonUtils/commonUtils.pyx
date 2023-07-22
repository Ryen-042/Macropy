# cython: embedsignature = True
# cython: language_level = 3str

"""This extension module provides class definitions used in other modules."""

from cythonExtensions.commonUtils cimport commonUtils
from cythonExtensions.commonUtils.commonUtils cimport KB_Con as kb_con

import win32gui, win32api, win32con, win32clipboard, pythoncom
import threading, queue, winsound
from win32com.client import Dispatch
from time import time
from collections import deque
from traceback import format_tb, format_exc
from typing import Callable

import scriptConfigs as configs
from cythonExtensions.windowHelper import windowHelper as winHelper


cdef class BaseEvent:
    """
    Description:
        The base class for all windows hook events.
    ---
    Parameters:
        `EventID -> int`: The event ID (the message code).
        
        `EventName -> str`: The name of the event (message).
        
        `Flags -> int`: The flags associated with the event.
    """
    
    cdef public int EventId, Flags
    cdef public str EventName
    
    def __init__(self, event_id: int, event_name: str, flags: int):
        self.EventId   = event_id
        self.EventName = event_name
        self.Flags     = flags
    
    def __repr__(self) -> str:
        return f"EvtId={self.EventId}, EvtName='{self.EventName}', Flags={self.Flags}"


cdef class KeyboardEvent(commonUtils.BaseEvent):
    """
    Description:
        Holds information about a keyboard event.
    ---
    Parameters:
        `EventID -> int`: The event ID (the message code).
        
        `EventName -> str`: The name of the event (message).
        
        `Flags -> int`: The flags associated with the event.
        
        `Key -> str`: The name of the key.
        
        `KeyID -> int`: The virtual key code.
        
        `Scancode -> int`: The key scancode.
        
        `Ascii -> int`: The ASCII value of the key.
        
        `Injected -> bool`: Whether or not the event was injected.
        
        `Extended -> bool`: Whether or not the event is an extended key event.
        
        `Shift -> bool`: Whether or not the shift key is pressed.
        
        `Alt -> bool`: Whether or not the alt key is pressed.
        
        `Transition -> bool`: Whether or not the key is transitioning from up to down.
    """
    
    cdef public int KeyID, Scancode, Ascii
    cdef public str Key
    cdef public bint Injected, Extended, Shift, Alt, Transition
    
    def __init__(self, event_id: int, event_name: str, vkey_code: int, scancode: int, key_ascii: int, key_name: str,
                 flags: int, injected: bint, extended: bint, shift: bint, alt: bint, transition: bint):
        
        super(commonUtils.KeyboardEvent, self).__init__(event_id, event_name, flags)
        
        self.KeyID = vkey_code
        self.Scancode = scancode
        self.Ascii = key_ascii
        self.Key = key_name
        self.Injected = injected
        self.Extended = extended
        self.Transition = transition
        self.Shift = shift
        self.Alt = alt
    
    def __repr__(self) -> str:
        return f"Key={self.Key}, ID={self.KeyID}, SC={self.Scancode}, Asc={self.Ascii} | Inj={self.Injected}, Ext={self.Extended}, Shift={self.Shift}, Alt={self.Alt}, Trans={self.Transition} | EvtId={self.EventId}, EvtName='{self.EventName}', Flags={self.Flags}"


cdef dict buttonMapping = {16: "LB", 8: "RB", 4: "MB", 2: "X1", 1: "X2"}
"""Maps the mouse button id to the button name."""


cdef class MouseEvent(commonUtils.BaseEvent):
    """
    Description:
        Holds information about a mouse event.
    ---
    Parameters:
        - `EventID -> int`: The event ID (the message code).
        
        - `EventName -> str`: The name of the event (message).
        
        - `Flags -> int`: The flags associated with the event.
        
        - `X -> int`: The X coordinate of the mouse pointer, relative to the top-left corner of the screen.
        
        - `Y -> int`: The Y coordinate of the mouse pointer, relative to the top-left corner of the screen.
        
        - `MouseData -> int`: The mouse movement distance, or the wheel delta, typically expressed in multiples or divisions of `WHEEL_DELTA`.
        
        - `IsMouseAbsolute -> bool`: Specifies whether the coordinates are mapped to the entire desktop or just to the particular window. Can have one of the following values:
            Value   | Meaning
            --------|--------
            `False` | Coordinates are relative to the window.
            `True`  | Coordinates are absolute.
        
        - `IsMouseInWindow -> bool`: Specifies whether the message was sent from the application's message queue (whether the mouse cursor was within the application's window). Can have one of the following values:
            Value | Meaning
            ------|--------
            False | The message was sent from the system message queue.
            True  | The message was sent from the application's message queue.
        
        - `Delta -> int`: The amount of wheel movement (positive for up, negative for down) which is a multiple of `WHEEL_DELTA`. This parameter is only valid for `WM_MOUSEWHEEL` messages. Can have positive or negative values, depending on the direction of the rotation:
            Value | Meaning
            ------|--------
            `+ve` | The wheel was rotated forward, away from the user.
            `-ve` | The wheel was rotated backward, toward the user.
        
        - `IsWheelHorizontal -> bool`: Specifies whether the wheel was moved horizontally. True if the message is `WM_MOUSEHWHEEL`.
        
        - `PressedButton -> int`: Specifies which mouse button was pressed. Can have one of the following values:
            Value | Meaning
            ------|--------
            16    | Left mouse button was pressed.
            8     | Right mouse button was pressed.
            4     | Middle mouse button was pressed.
            2     | X1 mouse button was pressed.
            1     | X2 mouse button was pressed.
    """
    
    cdef public int X, Y, MouseData, Delta, PressedButton # "Time"
    cdef public bint IsMouseAbsolute, IsMouseInWindow, IsWheelHorizontal
    
    def __init__(self, event_id: int, event_name: str, flags: int,
                 x: int, y: int, mouse_data: int, is_mouse_absolute: bool,
                 is_mouse_in_window: bool, wheel_delta: int, is_wheel_horizontal: bool,
                 pressed_button: int):
        super(commonUtils.MouseEvent, self).__init__(event_id, event_name, flags)
        
        self.X = x
        self.Y = y
        self.MouseData = mouse_data
        self.IsMouseAbsolute = is_mouse_absolute
        self.IsMouseInWindow = is_mouse_in_window
        self.Delta = wheel_delta
        self.IsWheelHorizontal = is_wheel_horizontal
        self.PressedButton = pressed_button
    
    def __repr__(self) -> str:
        return f"X={self.X}, Y={self.Y}, MouseData={self.MouseData}, IsMouseAbsolute={self.IsMouseAbsolute}, " \
               f"IsMouseInWindow={self.IsMouseInWindow}, Delta={self.Delta}, " \
               f"IsWheelHorizontal={self.IsWheelHorizontal}{f', PressedButton={buttonMapping[self.PressedButton]}' if self.PressedButton else ''}, " \
               f"EvtId={self.EventId}, EvtName='{self.EventName}', Flags={self.Flags}"


class ControllerHouse:
    """A class that stores information for managing and controlling keyboard."""
    
    __slots__ = ()
    
    modifiers = 0
    """An int packing the states of the keyboard modifier keys (pressed or not).
    Each of the least significant 14 bits (`0b00_0000_0000`) represent a single modifier key:
    Key Name | Value | Representation
    ---------|-------|-----------
    CTRL     | 8192  | 0b10000000000000
    LCTRL    | 4096  | 0b1000000000000
    RCTRL    | 2048  | 0b100000000000
    SHIFT    | 1024  | 0b10000000000
    LSHIFT   | 512   | 0b1000000000
    RSHIFT   | 256   | 0b100000000
    ALT      | 128   | 0b10000000
    LALT     | 64    | 0b1000000
    RALT     | 32    | 0b100000
    WIN      | 16    | 0b10000
    LWIN     | 8     | 0b1000
    RWIN     | 4     | 0b100
    FN       | 2     | 0b10
    BACKTICK | 1     | 0b1
    """
    
    # Masks for extracting individual keys from the `modifiers` packed int.
    CTRL     = 0b10000000000000 # 1 << 13 = 8192
    LCTRL    = 0b1000000000000  # 1 << 12 = 4096
    RCTRL    = 0b100000000000   # 1 << 11 = 2048
    SHIFT    = 0b10000000000    # 1 << 10 = 1024
    LSHIFT   = 0b1000000000     # 1 << 9  = 512
    RSHIFT   = 0b100000000      # 1 << 8  = 256
    ALT      = 0b10000000       # 1 << 7  = 128
    LALT     = 0b1000000        # 1 << 6  = 64
    RALT     = 0b100000         # 1 << 5  = 32
    WIN      = 0b10000          # 1 << 4  = 16
    LWIN     = 0b1000           # 1 << 3  = 8
    RWIN     = 0b100            # 1 << 2  = 4
    FN       = 0b10             # 1 << 1  = 2
    BACKTICK = 0b1              # 1 << 0  = 1
    
    # Composite Modifier keys masks
    CTRL_ALT_WIN_FN = CTRL   | ALT   | WIN | FN  # 0b10000010010010 # 8338
    CTRL_ALT_WIN    = CTRL   | ALT   | WIN       # 0b10000010010000 # 8336
    CTRL_WIN_FN     = CTRL   | FN    | WIN       # 0b10000000010010 # 8210
    CTRL_SHIFT      = CTRL   | SHIFT             # 0b10010000000000 # 9216
    CTRL_FN         = CTRL   | FN                # 0b10000000000010 # 8194
    CTRL_WIN        = CTRL   | WIN               # 0b10000000001100 # 8204
    CTRL_BACKTICK   = CTRL   | BACKTICK          # 0b10000000000001 # 8193
    LCTRL_RCTRL     = LCTRL  | RCTRL             # 0b01100000000000 # 6144
    SHIFT_ALT       = SHIFT  | ALT               # 0b00010010000000 # 1152
    SHIFT_FN        = SHIFT  | FN                # 0b00010000000010 # 1026
    SHIFT_BACKTICK  = SHIFT  | BACKTICK          # 0b00010000000001 # 1025
    LSHIFT_RSHIFT   = LSHIFT | RSHIFT            # 0b00001100000000 # 768
    ALT_FN          = ALT    | FN                # 0b00000010000010 # 130
    ALT_BACKTICK    = ALT    | BACKTICK          # 0b00000010000001 # 129
    LALT_RALT       = LALT   | RALT              # 0b00000001100000 # 96
    WIN_FN          = WIN    | FN                # 0b00000000010010 # 18
    WIN_BACKTICK    = WIN    | BACKTICK          # 0b00000000000001 # 17
    LWIN_RWIN       = LWIN   | RWIN              # 0b00000000001100 # 12
    FN_BACKTICK     = FN     | BACKTICK          # 0b00000000000011 # 3
    
    
    
    locks = (win32api.GetKeyState(win32con.VK_CAPITAL) << 2) | \
            (win32api.GetKeyState(win32con.VK_SCROLL)  << 1) | \
            win32api.GetKeyState(win32con.VK_NUMLOCK)
    """An int packing the states of the keyboard lock keys (on or off)."""
    
    # Masks for extracting individual lock keys from the `locks` packed int.
    CAPITAL = 0b100
    SCROLL  = 0b10
    NUMLOCK = 0b1
    
    pressed_chars = ""
    """Stores the pressed character keys for the key expansion events."""
    
    heldMouseBtn = 0
    """Stores the mouse button that is currently being held down (by sending mouse events with keyboard shortcuts). Possible Values:
    Value | Button
    ------|-------
    0     | None
    1     | Left
    2     | Right
    3     | Middle"""
    
    abbreviations = configs.ABBREVIATIONS
    """A dictionary of abbreviations and their corresponding expansion."""
    
    locations = configs.LOCATIONS
    """A dictionary of abbreviations and their corresponding path address."""


cpdef inline void updateModifiersPress(commonUtils.KeyboardEvent event):
    """Updates the state of the `modifiers` packed for keyDown events with the specified event."""
    
    ControllerHouse.modifiers |= (
        ((event.KeyID == win32con.VK_LCONTROL) | (event.KeyID == win32con.VK_RCONTROL)) << 13 | # CTRL
        ((event.KeyID == win32con.VK_LSHIFT)   | (event.KeyID == win32con.VK_RSHIFT))   << 10 | # SHIFT
        ((event.KeyID == win32con.VK_LMENU)    | (event.KeyID == win32con.VK_RMENU))    << 7  | # ALT
        ((event.KeyID == win32con.VK_LWIN)     | (event.KeyID == win32con.VK_RWIN))     << 4  |  # WIM
        
        # (event.KeyID == win32con.VK_LCONTROL) << 12 | # LCTRL
        # (event.KeyID == win32con.VK_RCONTROL) << 11 | # RCTRL
        # (event.KeyID == win32con.VK_LSHIFT)   << 9  | # LSHIFT
        # (event.KeyID == win32con.VK_RSHIFT)   << 8  | # RSHIFT
        # (event.KeyID == win32con.VK_LMENU)    << 6  | # LALT
        # (event.KeyID == win32con.VK_RMENU)    << 5  | # RALT
        # (event.KeyID == win32con.VK_LWIN)     << 3  | # LWIN
        # (event.KeyID == win32con.VK_RWIN)     << 2  | # RWIN
        
        (event.KeyID == 255)                  << 1  | # FN
        (event.KeyID == kb_con.VK_BACKTICK)           # BACKTICK
    )
    
    # ControllerHouse.modifiers |= (
    #     ((ControllerHouse.modifiers & ControllerHouse.LCTRL_RCTRL)   != 0) << 13 | # CTRL
    #     ((ControllerHouse.modifiers & ControllerHouse.LSHIFT_RSHIFT) != 0) << 10 | # SHIFT
    #     ((ControllerHouse.modifiers & ControllerHouse.LALT_RALT)     != 0) << 7  | # ALT
    #     ((ControllerHouse.modifiers & ControllerHouse.LWIN_RWIN)     != 0) << 4    # WIM
    # )


cpdef inline void updateModifiersRelease(commonUtils.KeyboardEvent event):
    """Updates the state of the `modifiers` packed for keyUp events with the specified event."""
    
    ControllerHouse.modifiers &= ~(
        ((event.KeyID == win32con.VK_LCONTROL) | (event.KeyID == win32con.VK_RCONTROL)) << 13 | # CTRL
        ((event.KeyID == win32con.VK_LSHIFT)   | (event.KeyID == win32con.VK_RSHIFT))   << 10 | # SHIFT
        ((event.KeyID == win32con.VK_LMENU)    | (event.KeyID == win32con.VK_RMENU))    << 7  | # ALT
        ((event.KeyID == win32con.VK_LWIN)     | (event.KeyID == win32con.VK_RWIN))     << 4  | # WIN
        
        # (1 << 13) | (1 << 10) | (1 << 7) | (1 << 4) | # Reseting CTRL, SHIFT, ALT, WIN
        # (event.KeyID == win32con.VK_LCONTROL) << 12 | # LCTRL
        # (event.KeyID == win32con.VK_RCONTROL) << 11 | # RCTRL
        # (event.KeyID == win32con.VK_LSHIFT)   << 9  | # LSHIFT
        # (event.KeyID == win32con.VK_RSHIFT)   << 8  | # RSHIFT
        # (event.KeyID == win32con.VK_LMENU)    << 6  | # LALT
        # (event.KeyID == win32con.VK_RMENU)    << 5  | # RALT
        # (event.KeyID == win32con.VK_LWIN)     << 3  | # LWIN
        # (event.KeyID == win32con.VK_RWIN)     << 2  | # RWIN
        
        (event.KeyID == 255)                  << 1  | # FN
        (event.KeyID == kb_con.VK_BACKTICK)           # BACKTICK
    )
    
    # ControllerHouse.modifiers &= ~( # Reseeting CTRL, SHIFT, ALT, WIN
    #     ((ControllerHouse.modifiers & ControllerHouse.LCTRL_RCTRL)   != 0) << 13 | # CTRL
    #     ((ControllerHouse.modifiers & ControllerHouse.LSHIFT_RSHIFT) != 0) << 10 | # SHIFT
    #     ((ControllerHouse.modifiers & ControllerHouse.LALT_RALT)     != 0) << 7  | # ALT
    #     ((ControllerHouse.modifiers & ControllerHouse.LWIN_RWIN)     != 0) << 4    # WIM
    # )


cpdef inline void updateLocks(commonUtils.KeyboardEvent event):
    """Updates the state of the `locks` packed for keyDown events with the specified event (no need to call for keyUp events)."""
    
    ControllerHouse.locks ^= (
        (event.KeyID == win32con.VK_CAPITAL) << 2 | # CAPITAL
        (event.KeyID == win32con.VK_SCROLL)  << 1 | # SCROLL
        (event.KeyID == win32con.VK_NUMLOCK)        # NUMLOCK
    )


cpdef void sendMouseScroll(steps=1, direction=1, wheelDelta=40):
    """
    Description:
        Sends a mouse scroll event with the specified number of steps and direction.
    ---
    Parameters:
        `steps: int = 1`:
            The number of steps to scroll. Can take negative values.
        
        `direction: int = 1`:
            The direction to scroll in. `1` for vertical, `0` for horizontal.
        
        `wheelDelta: int = 40`:
            The amount of scroll per step.
    """
    
    # API doc: https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-mouse_event
    win32api.mouse_event((win32con.MOUSEEVENTF_HWHEEL, win32con.MOUSEEVENTF_WHEEL)[direction], 0, 0, steps * wheelDelta, 0) # win32con.WHEEL_DELTA


cpdef inline void printModifiers():
    """Prints the states of the modifier keys after extracting them from the packed int `modifiers`."""
    
    print(f"CTRL={bool(ControllerHouse.modifiers & ControllerHouse.CTRL)}", end=", ")
    print(f"SHIFT={bool(ControllerHouse.modifiers & ControllerHouse.SHIFT)}", end=", ")
    print(f"ALT={bool(ControllerHouse.modifiers & ControllerHouse.ALT)}", end=", ")
    print(f"WIN={bool(ControllerHouse.modifiers & ControllerHouse.WIN)}", end=", ")
    print(f"FN={bool(ControllerHouse.modifiers & ControllerHouse.FN)}", end=", ")
    print(f"BACKTICK={bool(ControllerHouse.modifiers & ControllerHouse.BACKTICK)}", end=" | ")


cpdef inline void printLockKeys():
    """Prints the states of the lock keys after extracting them from the packed int `locks`."""
    
    print(f"CAPSLOCK={bool(ControllerHouse.locks & ControllerHouse.CAPITAL)}", end=", ")
    print(f"SCROLL={bool(ControllerHouse.locks & ControllerHouse.SCROLL)}", end=", ")
    print(f"NUMLOCK={bool(ControllerHouse.locks & ControllerHouse.NUMLOCK)}")


class MouseHouse:
    """A class that stores information for managing and controlling mouse controller."""
    
    __slots__ = ()
    
    x, y = 0, 0
    
    delta = 0
    """The mouse wheel scroll distance. Can take negative values."""
    
    horizontal = False
    """Whether the mouse wheel is scrolling horizontally or not."""
    
    buttons = 0
    """A packed int that stores the states of the mouse buttons (pressed or not).
    Each of the least significant 5 bits (`0b0_0000`) represent a single button:
    Button Name | Value | Representation
    ------------|-------|-----------
    LButton     | 16    | 0b10000
    RButton     | 8     | 0b1000
    MButton     | 4     | 0b100
    X1Button    | 2     | 0b10
    X2Button    | 1     | 0b1
    """
    
    LButton  = 0b10000 # 1 << 4 = 16
    RButton  = 0b1000  # 1 << 3 = 8
    MButton  = 0b100   # 1 << 2 = 4
    X1Button = 0b10    # 1 << 1 = 2
    X2Button = 0b1     # 1 << 0 = 1
    
    # Composite button masks
    LRButton  = LButton | RButton
    LRX1X2Button = LRButton | X1Button | X2Button


cpdef inline void updateButtonsPress(commonUtils.MouseEvent event):
    """Updates the mouse postion, scroll delta and direction, and the state of the `buttons` packed for keyUp events with the specified event."""
    
    MouseHouse.x, MouseHouse.y = event.X, event.Y
    
    MouseHouse.delta = event.Delta
    
    MouseHouse.horizontal = event.IsWheelHorizontal
    
    MouseHouse.buttons |= (
        (event.PressedButton == MouseHouse.LButton)  << 4 |
        (event.PressedButton == MouseHouse.RButton)  << 3 |
        (event.PressedButton == MouseHouse.MButton)  << 2 |
        (event.PressedButton == MouseHouse.X1Button) << 1 |
        (event.PressedButton == MouseHouse.X2Button)
    )


cpdef inline void updateButtonsRelease(commonUtils.MouseEvent event):
    """Updates the mouse postion, scroll delta and direction, and the state of the `buttons` packed for keyUp events with the specified event."""
    
    MouseHouse.x, MouseHouse.y = event.X, event.Y
    
    MouseHouse.delta = event.Delta
    
    MouseHouse.horizontal = event.IsWheelHorizontal
    
    MouseHouse.buttons &= ~(
        (event.PressedButton == MouseHouse.LButton)  << 4 |
        (event.PressedButton == MouseHouse.RButton)  << 3 |
        (event.PressedButton == MouseHouse.MButton)  << 2 |
        (event.PressedButton == MouseHouse.X1Button) << 1 |
        (event.PressedButton == MouseHouse.X2Button)
    )


cpdef inline void printButtons():
    """Prints the states of the buttons after extracting them from the packed int `buttons`."""
    
    print(f"LB={(MouseHouse.buttons & MouseHouse.LButton) == MouseHouse.LButton},",
          f"RB={(MouseHouse.buttons & MouseHouse.RButton) == MouseHouse.RButton},",
          f"MB={(MouseHouse.buttons & MouseHouse.MButton) == MouseHouse.MButton},",
          f"X1={(MouseHouse.buttons & MouseHouse.X1Button) == MouseHouse.X1Button},",
          f"X2={(MouseHouse.buttons & MouseHouse.X2Button) == MouseHouse.X2Button}",
          f"| pos=({MouseHouse.x}, {MouseHouse.y}) | WDelta={MouseHouse.delta} {'' if not MouseHouse.delta else 'H' if MouseHouse.horizontal else 'V'}\n")


class Management:
    """A class for storing variables and functions used for managing to the script."""
    
    __slots__ = ()
    
    counter = 0
    """Counts the number of times a key has been pressed (mod 10k)."""
    
    silent = configs.SUPPRESS_TERMINAL_OUTPUT
    """For suppressing terminal output. Defaults to True."""
    
    suppressKbInputs = False
    """For suppressing all keyboard keys. Defaults to `False`. Note that the hotkeys and text expainsion operations will still work."""
    
    terminateEvent = threading.Event()
    """An `Event` object that is set when the hotkey for terminating the script is pressed."""
    
    @staticmethod
    def toggleSilentMode() -> None:
        """Toggles the `silent` variable."""
        
        Management.silent ^= 1
        
        if Management.silent:
            winsound.PlaySound(r"SFX\no-trespassing-368.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
        else:
            winsound.PlaySound(r"SFX\pedantic-490.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
    
    @staticmethod
    def toggleKeyboardInputs():
        """Toggles the `suppressKbInputs` variable."""
        
        Management.suppressKbInputs ^= 1
    
        if Management.suppressKbInputs:
            winsound.PlaySound(r"SFX\no-trespassing-368.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
        else:
            winsound.PlaySound(r"SFX\pedantic-490.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
    
    @staticmethod
    def logUncaughtExceptions(exc_type, exc_value, exc_traceback) -> int:
        """Logs uncaught exceptions to the console and displays an error message with the exception traceback."""
        
        message = f"{exc_type.__name__}: {exc_value}\n  {''.join(format_tb(exc_traceback)).strip()}"
        print(f'\nUncaught exception: """{message}"""', end="\n\n")
        
        return winHelper.showMessageBox(message)


class WindowHouse:
    """A class for storing window class names and their corresponding window handles."""
    
    __slots__ = ()
    
    classNames: dict[str, int] = {"MediaPlayerClassicW": 0}
    """Stores the class names and the corresponding window handle for some progarms."""
    
    closedExplorers: deque[str] = deque(maxlen=10)
    """Stores the addresses of the 10 most recently closed windows explorers."""
    
    @staticmethod
    def getHandleByClassName(className: str) -> int:
        """Returns the window handle of the specified class name from the `WindowHouse.classNames` dict."""
        
        return WindowHouse.classNames.get(className, 0)
    
    @staticmethod
    def setHandleByClassName(className: str, value: int) -> None:
        """Assigns the given window handle to the specified class name in the `WindowHouse.classNames` dict."""
        
        WindowHouse.classNames[className] = value
    
    @staticmethod
    def rememberActiveProcessTitle(fg_hwnd=0) -> None:
        """
        Stores the title of the active window into the `closedExplorers` variable.
        The title of a windows explorer window is the address of the opened directory.
        """
        ...
        
        if not fg_hwnd:
            fg_hwnd = win32gui.GetForegroundWindow()
        
        cdef str explorerAddress = win32gui.GetWindowText(fg_hwnd)
        
        if not explorerAddress:
            print(f"Error: Could not get the title of the active window (id={fg_hwnd}) or the window does not have a title.")
            
            return
        
        if explorerAddress in WindowHouse.closedExplorers:
            WindowHouse.closedExplorers.remove(explorerAddress)
        
        WindowHouse.closedExplorers.append(explorerAddress)


class ShellAutomationObjectWrapper:
    """Thread-safe wrapper class for accessing an Automation object in a multithreaded environment."""
    # The wrapper object acquires a lock before accessing the Automation object, ensuring that only one thread can access it at a time.
    # This can help prevent problems that can occur when multiple threads try to access the Automation object concurrently.
    # If two or more threads tried to access an automation object wrapper, only one would be able to acquire the lock to the automation object at a time.
    # The other threads would be blocked until the lock is released by the thread that currently holds it.
    # When a thread lock object is used in a `with` statement, the lock is released when the `with` block ends.
    # The lock is automatically acquired at the beginning of the block and released when the block ends.
    
    __slots__ = ()
    
    explorer = Dispatch("Shell.Application")
    """An explorer Automation object."""
    
    _lock = threading.Lock()
    """A lock object used to ensure that only one thread can access the Automation object at a time."""
    
    @staticmethod
    def __getattr__(name: str):
        # Acquire lock before accessing the Automation object
        with ShellAutomationObjectWrapper._lock:
            return getattr(ShellAutomationObjectWrapper.explorer, name)


# Source: https://stackoverflow.com/questions/6552097/threading-how-to-get-parent-id-name
class PThread(threading.Thread):
    """An extension of `threading.Thread`. The class adds these features:
    - A `parent` attribute to the thread object.
    - Propagates exceptions from the created threads to the calling one and show them using error messages.
    - Defines ways for thread communication.
    - Defines ways for controlling the frequency or timing of certain events.
    """
    
    _throttle_lock = threading.Lock()
    """A lock object used to ensure that only one thread can access the critical section in the `throttle` decorator."""
    
    mainThreadId = threading.main_thread().ident
    """The ID of the main thread."""
    
    kbMsgQueue: Queue[bool] = queue.Queue()
    """A queue used for message passing between keyboard-related threads."""
    
    msMsgQueue: Queue[bool] = queue.Queue()
    """A queue used for message passing between mouse-related threads."""
    
    def __init__(self, *args, **kwargs) -> None:
        threading.Thread.__init__(self, *args, **kwargs)
        self.parent = threading.current_thread()
        self.coInitializeCalled = False
    
    # Propagating exceptions from a thread and showing an error message.
    def run(self):
        cdef str error_msg, class_name
        
        try:
            self.ret = self._target(*self._args, **self._kwargs)
        
        except BaseException as e:
            error_msg = format_exc()
            class_name = str(type(e)).split("'")[1]
            print(f'➤ Warning! An error of type "{class_name}" occurred in thread {self.name}.\n\n→ Error message: {str(e)}\n\n→ {error_msg}\n{"="*50}\n\n')
            winHelper.showMessageBox(error_msg)
            
            # Management.logUncaughtExceptions(*sys.exc_info())
            # logging.error(f"Warning! An error occurred in thread {self.name}. Traceback:{str(e)}", exc_info=True)
    
    @staticmethod
    def inMainThread() -> bool:
        """Returns whether the current thread is the main thread for the current process."""
        
        return threading.get_ident() == PThread.mainThreadId
    
    @staticmethod
    def getParentThread() -> int:
        """
        Description:
            Returns the parent thread id of the current thread.
        ---
        Returns:
            - `int > 0` (the parent thread id): if the current thread is not the main thread.
            - `0`: if it is the main thread.
            - `-1`: if the parent thread is unknown (i.e., the current thread was not created using this class).
        """
        
        if threading.current_thread().name != "MainThread":
            if hasattr(threading.current_thread(), 'parent'):
                return threading.current_thread().parent.ident
            
            else:
                return -1 # Unknown. The current thread does not have a `parent` property.
        
        else:
            return 0 # "MainThread"
    
    @staticmethod
    def coInitialize() -> bool:
        """Initializes the COM library for the current thread if it was not previously initialized."""
        
        if not PThread.inMainThread() and not threading.current_thread().coInitializeCalled:
            print(f"CoInitialize called from: {threading.current_thread().name}")
            threading.current_thread().coInitializeCalled = True
            pythoncom.CoInitialize()
            
            return True
        
        return False
    
    @staticmethod
    def coUninitialize() -> None:
        """Uninitializes the COM library for the current thread if `initializer_called` is True."""
        if threading.current_thread().coInitializeCalled:
            print(f"CoUninitialize called from: {threading.current_thread().name}")
            pythoncom.CoUninitialize()
            threading.current_thread().coInitializeCalled = False
    
    # Source: https://github.com/salesforce/decorator-operations/blob/master/decoratorOperations/throttle_functions/throttle.py
    # For a comparison between Throttling and Debouncing: https://stackoverflow.com/questions/25991367/difference-between-throttling-and-debouncing-a-function
    @staticmethod
    def throttle(wait_time: float):
        """
        Description:
            Decorator function that ensures that a wrapped function is called only once in each time slot.
        """
        
        def decorator(function):
            def throttled(*args, **kwargs):
                def call_function():
                    return function(*args, **kwargs)
                
                if PThread._throttle_lock.acquire(blocking = False):
                    if time() - throttled._last_time_called >= wait_time:
                        call_function()
                        throttled._last_time_called = time()
                    PThread._throttle_lock.release()
                
                else: # Lock is already acquired by another thread, so return without calling the wrapped function.
                    return
            
            throttled._last_time_called = 0
            
            return throttled
        
        return decorator
    
    @staticmethod
    def debounce(wait_time: float):
        """
        Decorator that will debounce a function so that it is called after `wait_time` seconds.
        If the decorated function is called multiple times within a time slot, it will debounce
        (wait for a new time slot from the last call without executing the associated function).
        
        If no calls arrive after `wait_time` seconds from the last call, then execute the last call.
        
        A function decorated with this decorator can also be called immediately using `func_name.runNoWait()`.
        """
        
        def decorator(function: Callable):
            def debounced(*args, **kwargs):
                def call_function():
                    debounced._timer = None
                    return function(*args, **kwargs)
                
                if debounced._timer is not None:
                    debounced._timer.cancel()
                
                debounced._timer = threading.Timer(wait_time, call_function)
                debounced._timer.start()
            
            def runNoWait(*args, **kwargs):
                # If the timer is running, cancel it.
                if debounced._timer is not None:
                    debounced._timer.cancel()
                    debounced._timer = None
                
                # Now, no timer is running so call the function.
                function(*args, **kwargs)
            
            debounced._timer = None
            debounced.runNoWait = runNoWait
            
            return debounced
        
        return decorator


cpdef str readFromClipboard(int CF=win32clipboard.CF_TEXT): # CF: Clipboard format.
    """Reads the top of the clipboard if it was the same type as the specified."""
    
    win32clipboard.OpenClipboard()
    if win32clipboard.IsClipboardFormatAvailable(CF):
        clipboard_data = win32clipboard.GetClipboardData()
        win32clipboard.EmptyClipboard()
        win32clipboard.CloseClipboard()
    else:
        return ""
    
    return clipboard_data


cpdef void sendToClipboard(data, int CF=win32clipboard.CF_UNICODETEXT):
    """Copies the given data to the clipboard."""
    
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(CF, data)
    win32clipboard.CloseClipboard()
