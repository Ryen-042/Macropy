"""This module provides class definitions used in other modules."""

import threading, queue, win32clipboard, multiprocessing
from collections import deque
from enum import IntEnum
from win32com.client import CDispatch

import scriptConfigs as configs


class KB_Con(IntEnum):
    """
    Description:
        An enum class for storing keyboard `keyID` and `scancode` constants that are not present in the `win32con` module.
    ---
    Naming Scheme:
        `AS`: Ascii value.
        `VK`: Virtual keyID.
        `SC`: Scancode
    ---
    Notes:
        - Keys may send different `Ascii` values depending on the pressed modifier(s), but they send the same `keyID` and `scancode`.
            - If you need a code that is independent of the pressed modifiers, use `keyID`.
            - If you need a code that may have different values, use `Ascii`. Eg, Ascii of `=` is `61`, `+` (`Shift` + `=`) is `41`.
        - The class has only one copy of `KeyID` and `scancode` constants for each physical key.
            - These constants are named with respect to the pressed key with no modifiers.
            - To keep the naming scheme consistent with capitalizing all characters, letter keys are named using the uppercase version.
        - Capital letters (Shift + letter key) have Ascii values equal to their VK values. As such, the class only stores the Ascii of lowercase letters.
    """
    
    # Letter keys - Uppercase letters
    AS_a = 97; VK_A = 65;  SC_A = 30
    AS_b = 98; VK_B = 66;  SC_B = 48
    AS_c = 99; VK_C = 67;  SC_C = 46
    AS_d = 100; VK_D = 68;  SC_D = 32
    AS_e = 101; VK_E = 69;  SC_E = 18
    AS_f = 102; VK_F = 70;  SC_F = 33
    AS_g = 103; VK_G = 71;  SC_G = 34
    AS_h = 104; VK_H = 72;  SC_H = 35
    AS_i = 105; VK_I = 73;  SC_I = 23
    AS_j = 106; VK_J = 74;  SC_J = 36
    AS_k = 107; VK_K = 75;  SC_K = 37
    AS_l = 108; VK_L = 76;  SC_L = 38
    AS_m = 109; VK_M = 77;  SC_M = 50
    AS_n = 110; VK_N = 78;  SC_N = 49
    AS_o = 111; VK_O = 79;  SC_O = 24
    AS_p = 112; VK_P = 80;  SC_P = 25
    AS_q = 113; VK_Q = 81;  SC_Q = 16
    AS_r = 114; VK_R = 82;  SC_R = 19
    AS_s = 115; VK_S = 83;  SC_S = 31
    AS_t = 116; VK_T = 84;  SC_T = 20
    AS_u = 117; VK_U = 85;  SC_U = 22
    AS_v = 118; VK_V = 86;  SC_V = 47
    AS_w = 119; VK_W = 87;  SC_W = 17
    AS_x = 120; VK_X = 88;  SC_X = 45
    AS_y = 121; VK_Y = 89;  SC_Y = 21
    AS_z = 122; VK_Z = 90;  SC_Z = 44
    
    # Number keys
    VK_0 = 48;  SC_0 = 11
    VK_1 = 49;  SC_1 = 2
    VK_2 = 50;  SC_2 = 3
    VK_3 = 51;  SC_3 = 4
    VK_4 = 52;  SC_4 = 5
    VK_5 = 53;  SC_5 = 6
    VK_6 = 54;  SC_6 = 7
    VK_7 = 55;  SC_7 = 8
    VK_8 = 56;  SC_8 = 9
    VK_9 = 57;  SC_9 = 10
    
    # Symbol keys
    AS_EXCLAM               = 33
    AS_DOUBLE_QUOTES        = 34;   VK_SINGLE_QUOTES = 222;       SC_SINGLE_QUOTES = 40
    AS_HASH                 = 35
    AS_DOLLAR               = 36
    AS_PERCENT              = 37
    AS_AMPERSAND            = 38
    AS_SINGLE_QUOTE         = 39
    AS_OPEN_PAREN           = 40
    AS_CLOSE_PAREN          = 41
    AS_ASTERISK             = 42
    AS_PLUS                 = 43
    AS_COMMA                = 44;   VK_COMMA  = 188;                SC_COMMA  = 51
    AS_MINUS                = 45;   VK_MINUS  = 189;                SC_MINUS  = 12
    AS_PERIOD               = 46;   VK_PERIOD = 190;                SC_PERIOD = 52
    AS_SLASH                = 47;   VK_SLASH  = 191;                SC_SLASH  = 53
    
    AS_COLON                = 58
    AS_SEMICOLON            = 59;   VK_SEMICOLON = 186;             SC_SEMICOLON = 39
    AS_LESS_THAN            = 60
    AS_EQUALS               = 61;   VK_EQUALS = 187;                SC_EQUALS = 13
    AS_GREATER_THAN         = 62
    AS_QUESTION_MARK        = 63
    AS_AT                   = 64
    
    AS_OPEN_SQUARE_BRACKET  = 91;   VK_OPEN_SQUARE_BRACKET = 219;   SC_OPEN_SQUARE_BRACKET = 26
    AS_BACKSLASH            = 92;   VK_BACKSLASH = 220;             SC_BACKSLASH = 43
    AS_CLOSE_SQUARE_BRACKET = 93;   VK_CLOSE_SQUARE_BRACKET = 221;  SC_CLOSE_SQUARE_BRACKET = 27
    AS_CARET                = 94
    AS_UNDERSCORE           = 95
    AS_BACKTICK             = 96;   VK_BACKTICK = 192;              SC_BACKTICK = 41 # '`'
    
    AS_OPEN_CURLY_BRACE     = 123
    AS_PIPE                 = 124
    AS_CLOSE_CURLY_BRACE    = 125
    AS_TILDE                = 126
    AS_OEM_102_CTRL         = 28;   VK_AS_OEM_102 = 226;            SC_OEM_102 = 86 # '\' in European keyboards (between LShift and Z).
    
    # Miscellaneous keys
    SC_RETURN      = 28 # Enter
    SC_BACK        = 14 # Backspace
    SC_MENU        = 56 # 'LMenu' and 'RMenu'
    SC_HOME        = 71
    SC_UP          = 72
    SC_RIGHT       = 77
    SC_DOWN        = 80
    SC_LEFT        = 75
    SC_VOLUME_UP   = 48 # Make sure that this doesn't conflict with `KB_Con.SC_B` as both have the same value.
    SC_VOLUME_DOWN = 46


class BaseEvent:
    """
    Description:
        The base class for all windows hook events.
    ---
    Parameters:
        `EventID -> int`: The event ID (the message code).
        
        `EventName -> str`: The name of the event (message).
        
        `Flags -> int`: The flags associated with the event.
    """
    
    __slots__ = ("EventId", "Flags", "EventName")
    
    def __init__(self, event_id: int, event_name: str, flags: int):
        ...
    def __repr__(self) -> str:
        return f"EvtId={self.EventId}, EvtName='{self.EventName}', Flags={self.Flags}"


class KeyboardEvent(BaseEvent):
    """
    Description:
        Holds information about a keyboard event.
    ---
    Parameters:
        `EventID -> int`: The event ID (the message code).
        
        `EventName -> str`: The name of the event (message).
        
        `Key -> str`: The name of the key.
        
        `KeyID -> int`: The virtual key code.
        
        `Scancode -> int`: The key scancode.
        
        `Ascii -> int`: The ASCII value of the key.
        
        `Flags -> int`: The flags associated with the event.
        
        `Injected -> bool`: Whether or not the event was injected.
        
        `Extended -> bool`: Whether or not the event is an extended key event.
        
        `Shift -> bool`: Whether or not the shift key is pressed.
        
        `Alt -> bool`: Whether or not the alt key is pressed.
        
        `Transition -> bool`: Whether or not the key is transitioning from up to down.
    """
    
    __slots__ = ("KeyID", "Scancode", "Ascii", "Key", "Injected", "Extended", "Shift", "Alt", "Transition")
    
    def __init__(self, event_id: int, event_name: str, vkey_code: int, scancode: int, key_ascii: int, key_name: str,
                 flags: int, injected: bool, extended: bool, shift: bool, alt: bool, transition: bool):
        ...
    
    def __repr__(self) -> str:
        return f"Key={self.Key}, ID={self.KeyID}, SC={self.Scancode}, Asc={self.Ascii} | Inj={self.Injected}, Ext={self.Extended}, Shift={self.Shift}, Alt={self.Alt}, Trans={self.Transition}, Flags={self.Flags} | EvtId={self.EventId}, EvtName='{self.EventName}'"


buttonMapping = {16: "LB", 8: "RB", 4: "MB", 2: "X1", 1: "X2"}
"""Maps the mouse button id to the button name."""


class MouseEvent(BaseEvent):
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
    
    __slots__ = ("X", "Y", "MouseData", "ExtraInfo", # "Time",
                 "IsMouseAbsolute", "IsMouseInWindow", "Delta",
                 "IsWheelHorizontal", "PressedButton")
    
    def __init__(self, event_id: int, event_name: str, flags: int,
                 x: int, y: int, mouse_data: int, is_mouse_absolute: bool,
                 is_mouse_in_window: bool, wheel_delta: int, is_wheel_horizontal: bool,
                 pressed_button: int):
        ...
    
    def __repr__(self) -> str:
        return f"X={self.X}, Y={self.Y}, MouseData={self.MouseData}, IsMouseAbsolute={self.IsMouseAbsolute}, " \
               f"IsMouseInWindow={self.IsMouseInWindow}, Delta={self.Delta}, " \
               f"IsWheelHorizontal={self.IsWheelHorizontal}{f', PressedButton={buttonMapping[self.PressedButton]}' if self.PressedButton else ''}, " \
               f"EvtId={self.EventId}, EvtName='{self.EventName}', Flags={self.Flags}"


class ControllerHouse:
    """A class that stores information for managing and controlling keyboard and mouse."""
    
    __slots__ = ()
    
    modifiers: int
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
    CTRL = 0b10000000000000 # 1 << 13 = 8192
    """A mask for extracting the `CTRL` flag from the `modifiers` packed int."""
    
    LCTRL = 0b1000000000000 # 1 << 12 = 4096
    """A mask for extracting the `LCTRL` flag from the `modifiers` packed int."""
    
    RCTRL = 0b100000000000  # 1 << 11 = 2048
    """A mask for extracting the `RCTRL` flag from the `modifiers` packed int."""
    
    SHIFT = 0b10000000000   # 1 << 10 = 1024
    """A mask for extracting the `SHIFT` flag from the `modifiers` packed int."""
    
    LSHIFT = 0b1000000000   # 1 << 9  = 512
    """A mask for extracting the `LSHIFT` flag from the `modifiers` packed int."""
    
    RSHIFT = 0b100000000    # 1 << 8  = 256
    """A mask for extracting the `RSHIFT` flag from the `modifiers` packed int."""
    ALT = 0b10000000        # 1 << 7  = 128
    """A mask for extracting the `ALT` flag from the `modifiers` packed int."""
    
    LALT = 0b1000000        # 1 << 6  = 64
    """A mask for extracting the `LALT` flag from the `modifiers` packed int."""
    
    RALT = 0b100000         # 1 << 5  = 32
    """A mask for extracting the `RALT` flag from the `modifiers` packed int."""
    
    WIN = 0b10000           # 1 << 4  = 16
    """A mask for extracting the `WIN` flag from the `modifiers` packed int."""
    
    LWIN = 0b1000           # 1 << 3  = 8
    """A mask for extracting the `LWIN` flag from the `modifiers` packed int."""
    
    RWIN = 0b100            # 1 << 2  = 4
    """A mask for extracting the `RWIN` flag from the `modifiers` packed int."""
    
    FN = 0b10               # 1 << 1  = 2
    """A mask for extracting the `FN` flag from the `modifiers` packed int."""
    
    BACKTICK = 0b1          # 1 << 0  = 1
    """A mask for extracting the `BACKTICK` flag from the `modifiers` packed int."""
    
    # Composite Modifier keys masks
    CTRL_SHIFT_ALT  = 0b10010010000000 # 9216
    """A mask for extracting the `CTRL`, `SHIFT` and `ALT` flags from the `modifiers` packed int."""
    
    CTRL_ALT_WIN_FN = 0b10000010010010 # 8338
    """A mask for extracting the `CTRL`,` ALT`, `WIN` and `FN` flags from the `modifiers` packed int."""
    
    CTRL_ALT_WIN    = 0b10000010010000 # 8336
    """A mask for extracting the `CTRL`, `ALT` and `WIN` flags from the `modifiers` packed int."""
    
    CTRL_WIN_FN     = 0b10000000010010 # 8210
    """A mask for extracting the `CTRL`, `WIN` and `FN` flags from the `modifiers` packed int."""
    
    CTRL_SHIFT      = 0b10010000000000 # 9216
    """A mask for extracting the `CTRL` and `SHIFT` flags from the `modifiers` packed int."""
    
    CTRL_ALT        = 0b10000010000000 # 8320
    """A mask for extracting the `CTRL` and `ALT` flags from the `modifiers` packed int."""
    
    CTRL_FN         = 0b10000000000010 # 8194
    """A mask for extracting the `CTRL` and `FN` flags from the `modifiers` packed int."""
    
    CTRL_WIN        = 0b10000000001100 # 8204
    """A mask for extracting the `CTRL` and `WIN` flags from the `modifiers` packed int."""
    
    CTRL_BACKTICK   = 0b10000000000001 # 8193
    """A mask for extracting the `CTRL` and `BACKTICK` flags from the `modifiers` packed int."""
    
    LCTRL_RCTRL     = 0b01100000000000 # 6144
    """A mask for extracting the `LCTRL` and `RCTRL` flags from the `modifiers` packed int."""
    
    SHIFT_ALT       = 0b00010010000000 # 1152
    """A mask for extracting the `SHIFT` and `ALT` flags from the `modifiers` packed int."""
    
    SHIFT_FN        = 0b00010000000010 # 1026
    """A mask for extracting the `SHIFT` and `FN` flags from the `modifiers` packed int."""
    
    SHIFT_BACKTICK  = 0b00010000000001 # 1025
    """A mask for extracting the `SHIFT` and `BACKTICK` flags from the `modifiers` packed int."""
    
    LSHIFT_RSHIFT   = 0b00001100000000 # 768
    """A mask for extracting the `LSHIFT` and `RSHIFT` flags from the `modifiers` packed int."""
    
    ALT_FN          = 0b00000010000010 # 130
    """A mask for extracting the `ALT` and `FN` flags from the `modifiers` packed int."""
    
    ALT_BACKTICK    = 0b00000010000001 # 129
    """A mask for extracting the `ALT` and `BACKTICK` flags from the `modifiers` packed int."""
    
    LALT_RALT       = 0b00000001100000 # 96
    """A mask for extracting the `LALT` and `RALT` flags from the `modifiers` packed int."""
    
    WIN_FN          = 0b00000000010010 # 18
    """A mask for extracting the `WIN and `FN` flags from the `modifiers` packed int."""
    
    WIN_BACKTICK    = 0b00000000000001 # 17
    """A mask for extracting the `WIN` and `BACKTICK` flags from the `modifiers` packed int."""
    
    LWIN_RWIN       = 0b00000000001100 # 12
    """A mask for extracting the `LWIN` and `RWIN` flags from the `modifiers` packed int."""
    
    FN_BACKTICK     = 0b00000000000011 # 3
    """A mask for extracting the `FN` and `BACKTICK` flags from the `modifiers` packed int."""
    
    # Boolean variables for storing the state of the lock keys.
    CAPITAL: bool
    """A mask for extracting the `CAPITAL` (CAPSLOCK) flag from the `locks` packed int."""
    
    SCROLL: bool
    """A mask for extracting the `SCROLL` flag from the `locks` packed int."""
    
    NUMLOCK: bool
    """A mask for extracting the `NUMLOCK` flag from the `locks` packed int."""
    
    # locks : int
    # """An int packing the states of the keyboard lock keys (on or off)."""
    
    # Masks for extracting individual lock keys from the `locks` packed int.
    CAPITAL_MASK = 0b100
    SCROLL_MASK  = 0b10
    NUMLOCK_MASK = 0b1
    
    pressed_chars : str
    """Stores the pressed character keys for the key expansion events."""
    
    pressed_chars_backup : str
    """Stores a backup that lasts one keystroke to undo replacement of text if backspace is pressed immediately after expainsion."""
    
    heldMouseBtn = 0
    """Stores the mouse button that is currently being held down (by sending mouse events with keyboard shortcuts). Possible Values:
    Value | Button
    ------|-------
    0     | None
    1     | Left
    2     | Right
    3     | Middle"""
    
    abbreviations = configs.ABBREVIATIONS
    """A dictionary of aliases and their corresponding expansion."""
    
    non_prefixed_abbreviations = configs.NON_PERFIXED_ABBREVIATIONS
    """A dictionary of aliases and their corresponding expansion that are not prefixed with one of the defined prefixes."""
    
    locations = configs.LOCATIONS
    """A dictionary of aliases and their corresponding path address expansions."""
    
    max_alias_length = configs.MAX_ALIAS_LENGTH
    """The maximum length of an alias."""
    
    burstClicksActive = False
    """A boolean that indicates whether burst-inputs is enable or not"""
    
    @staticmethod
    def getModifiersStates() -> list[bool]:
        """
        Returns a list of boolean values representing the states of the modifier keys. The order of the returned values is:
        `[CTRL, SHIFT, ALT, WIN, FN, BACKTICK]`
        """
        ...


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


class Management:
    """A class for storing variables and functions used for managing to the script."""
    
    __slots__ = ()
    
    counter : int
    """Counts the number of times a key has been pressed (mod 10k)."""
    
    silent : bool
    """For suppressing terminal output. Defaults to True."""
    
    suppressKbInputs : bool
    """For suppressing all keyboard keys. Defaults to `False`. Note that the hotkeys and text expainsion operations will still work."""
    
    terminateEvent: threading.Event
    """An `Event` object that is set when the hotkey for terminating the script is pressed."""
    
    isBacktickTheOnlyModiferPressed: bool
    """Whether the backtick key is pressed with other modifiers or not."""
    
    mouseVolumeControlSVar = multiprocessing.RawValue("b", False)
    """A shared boolean value used for specifying whether the mouse volume control hotkey is pressed or not"""
    
    @staticmethod
    def toggleSilentMode() -> None:
        """Toggles the `silent` variable."""
        ...
    
    @staticmethod
    def toggleKeyboardInputs():
        """Toggles the `suppressKbInputs` variable."""
        ...
    
    @staticmethod
    def logUncaughtExceptions(exc_type, exc_value, exc_traceback) -> int:
        """Logs uncaught exceptions to the console and displays an error message with the exception traceback."""
        ...


class WindowHouse:
    """A class for storing window class names and their corresponding window handles."""
    
    __slots__ = ()
    
    classNames: dict[str, int]
    """Stores the class names and the corresponding window handle for some progarms."""
    
    closedExplorers: deque[str]
    """Stores the addresses of the 10 most recently closed windows explorers."""
    
    @staticmethod
    def getHandleByClassName(className: str) -> int:
        """Returns the window handle of the specified class name from the `WindowHouse.classNames` dict."""
        ...
    
    @staticmethod
    def setHandleByClassName(className: str, value: int) -> None:
        """Assigns the given window handle to the specified class name in the `WindowHouse.classNames` dict."""
        ...
    
    @staticmethod
    def rememberActiveProcessTitle(fg_hwnd=0) -> None:
        """
        Stores the title of the active window into the `closedExplorers` variable.
        The title of a windows explorer window is the address of the opened directory.
        """
        ...


class ShellAutomationObjectWrapper:
    """Thread-safe wrapper class for accessing an Automation object in a multithreaded environment."""
    
    __slots__ = ()
    
    explorer: CDispatch
    """An explorer Automation object."""
    
    _lock: threading.Lock
    """A lock object used to ensure that only one thread can access the Automation object at a time."""


# Source: https://stackoverflow.com/questions/6552097/threading-how-to-get-parent-id-name
class PThread(threading.Thread):
    """An extension of `threading.Thread`. The class adds these features:
    - A `parent` attribute to the thread object.
    - Propagates exceptions from the created threads to the calling one and show them using error messages.
    - Defines ways for thread communication.
    - Defines ways for controlling the frequency or timing of certain events.
    """
    
    mainThreadId = threading.main_thread().ident
    """The ID of the main thread."""
    
    kbMsgQueue: queue.Queue[bool]
    """A queue used for message passing between threads."""
    
    def __init__(self, *args, **kwargs) -> None:
        ...
    
    @staticmethod
    def inMainThread() -> bool:
        """Returns whether the current thread is the main thread for the current process."""
        ...
    
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
        ...
    
    @staticmethod
    def coInitialize() -> bool:
        """Initializes the COM library for the current thread if it was not previously initialized."""
        ...
    
    @staticmethod
    def coUninitialize() -> None:
        """Uninitializes the COM library for the current thread if `initializer_called` is True."""
        ...
    
    # Source: https://github.com/salesforce/decorator-operations/blob/master/decoratorOperations/throttle_functions/throttle.py
    # For a comparison between Throttling and Debouncing: https://stackoverflow.com/questions/25991367/difference-between-throttling-and-debouncing-a-function
    @staticmethod
    def throttle(wait_time: float):
        """
        Description:
            Decorator function that ensures that a wrapped function is called only once in each time slot.
        """
        ...
    
    @staticmethod
    def debounce(wait_time: float):
        """
        Decorator that will debounce a function so that it is called after `wait_time` seconds.
        If the decorated function is called multiple times within a time slot, it will debounce
        (wait for a new time slot from the last call without executing the associated function).
        
        If no calls arrive after `wait_time` seconds from the last call, then execute the last call.
        
        A function decorated with this decorator can also be called immediately using `func_name.runNoWait()`.
        """
        ...


def readFromClipboard(CF= win32clipboard.CF_TEXT) -> str: # CF: Clipboard format.
    """Reads the top of the clipboard if it was the same type as the specified."""
    ...


def sendToClipboard(data, CF=win32clipboard.CF_UNICODETEXT):
    """Copies the given data to the clipboard."""
    ...
