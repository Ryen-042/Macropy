import ctypes
import ctypes.wintypes
from typing import Callable, Any
from enum import IntEnum
from cythonExtensions.commonUtils.commonUtils import KeyboardEvent

# https://learn.microsoft.com/en-us/windows/win32/winmsg/about-hooks
class HookTypes(IntEnum):
    # Constants that represent different types of Windows hooks that can be set using the SetWindowsHookEx function.
    WH_MSGFILTER       = -1
    """A hook type that allows you to monitor messages sent to a window or dialog box procedure. This is also the minmum value for a hook type."""

    WH_JOURNALRECORD   = 0
    """A hook type that is used to record input events."""

    WH_JOURNALPLAYBACK = 1
    """A hook type that is used to replay the input events recorded by the WH_JOURNALRECORD hook type."""

    WH_KEYBOARD        = 2
    """A hook type that is used to monitor keystrokes. It can be used to implement a keylogger or to perform other keyboard-related tasks."""

    WH_GETMESSAGE      = 3
    """A hook type that is used to monitor messages in the message queue of a thread. It can be used to modify or filter messages before they are dispatched."""

    WH_CALLWNDPROC     = 4
    """A hook type that is used to monitor messages sent to a window or dialog box procedure before they are processed by the procedure."""

    WH_CBT             = 5
    """A hook type that is used to monitor a variety of system events, including window creation and destruction, window activation, and window focus changes."""

    WH_SYSMSGFILTER    = 6
    """A hook type that is similar to WH_MSGFILTER, but it monitors messages sent to all windows, not just a specific window or dialog box procedure."""

    WH_HARDWARE        = 8
    """A hook type that is used to monitor low-level hardware events, such as mouse and keyboard input."""

    WH_DEBUG           = 9
    """A hook type that is used for debugging purposes."""

    WH_SHELL           = 10
    """A hook type that is used to monitor shell-related events, such as when the shell is about to activate a new window."""

    WH_FOREGROUNDIDLE  = 11
    """A hook type that is used to monitor when the system enters an idle state and the foreground application is not processing input."""

    WH_CALLWNDPROCRET  = 12
    """A hook type that is similar to WH_CALLWNDPROC, but it monitors messages after they have been processed by the procedure."""

    WH_KEYBOARD_LL     = 13
    """A hook type that is similar to WH_KEYBOARD, but it is a low-level keyboard hook that can be used to monitor keystrokes from all processes."""

    WH_MAX             = 15
    """The maximum value for a hook type. It is not a valid hook type itself."""

class KB_MsgTypes(IntEnum):
    # Constants that represent different types of keyboard-related Windows messages that can be received by a window or a message loop.
    WM_KEYFIRST    = 0x0100
    """Defines the minimum value for the range of keyboard-related messages. """

    WM_KEYDOWN     = 0x0100
    """A keyboard key was pressed."""

    WM_KEYUP       = 0x0101
    """A keyboard key was released."""

    WM_CHAR        = 0x0102
    """A keyboard key was pressed down and released, and it represents a printable character."""

    WM_DEADCHAR    = 0x0103
    """A keyboard key was pressed down and released, and it represents a dead character."""

    WM_SYSKEYDOWN  = 0x0104
    """A keyboard key was pressed while the ALT key was held down."""

    WM_SYSKEYUP    = 0x0105
    """A keyboard key was released after the ALT key was held down."""

    WM_SYSCHAR     = 0x0106
    """A keyboard key was pressed down and released, and it represents a printable character while the ALT key was held down."""

    WM_SYSDEADCHAR = 0x0107
    """A keyboard key was pressed down and released, and it represents a dead character while the ALT key was held down."""

    WM_KEYLAST     = 0x0108
    """Defines the maximum value for the range of keyboard-related messages."""


vKeyNameToId = {
    "VK_LBUTTON":           0x01,  "VK_RBUTTON":          0x02,  "VK_CANCEL":              0x03,  "VK_MBUTTON":        0x04, 
    "VK_BACK":              0x08,  "VK_TAB":              0x09,  "VK_CLEAR":               0x0C,  "VK_RETURN":         0x0D,  "VK_SHIFT":        0x10,
    "VK_CONTROL":           0x11,  "VK_MENU":             0x12,  "VK_PAUSE":               0x13,  "VK_CAPITAL":        0x14,  "VK_KANA":         0x15,
    "VK_HANGEUL":           0x15,  "VK_HANGUL":           0x15,  "VK_JUNJA":               0x17,  "VK_FINAL":          0x18,  "VK_HANJA":        0x19,
    "VK_KANJI":             0x19,  "VK_ESCAPE":           0x1B,  "VK_CONVERT":             0x1C,  "VK_NONCONVERT":     0x1D,  "VK_ACCEPT":       0x1E,
    "VK_MODECHANGE":        0x1F,  "VK_SPACE":            0x20,  "VK_PRIOR":               0x21,  "VK_NEXT":           0x22,  "VK_END":          0x23,
    "VK_HOME":              0x24,  "VK_LEFT":             0x25,  "VK_UP":                  0x26,  "VK_RIGHT":          0x27,  "VK_DOWN":         0x28,
    "VK_SELECT":            0x29,  "VK_PRINT":            0x2A,  "VK_EXECUTE":             0x2B,  "VK_SNAPSHOT":       0x2C,  "VK_INSERT":       0x2D,
    "VK_DELETE":            0x2E,  "VK_HELP":             0x2F,  "VK_LWIN":                0x5B,  "VK_RWIN":           0x5C,  "VK_APPS":         0x5D,
    "VK_NUMPAD0":           0x60,  "VK_NUMPAD1":          0x61,  "VK_NUMPAD2":             0x62,  "VK_NUMPAD3":        0x63,  "VK_NUMPAD4":      0x64,
    "VK_NUMPAD5":           0x65,  "VK_NUMPAD6":          0x66,  "VK_NUMPAD7":             0x67,  "VK_NUMPAD8":        0x68,  "VK_NUMPAD9":      0x69,
    "VK_MULTIPLY":          0x6A,  "VK_ADD":              0x6B,  "VK_SEPARATOR":           0x6C,  "VK_SUBTRACT":       0x6D,  "VK_DECIMAL":      0x6E,  "VK_DIVIDE": 0x6F, 
    "VK_F1":                0x70,  "VK_F2":               0x71,  "VK_F3":                  0x72,  "VK_F4":             0x73,  "VK_F5":           0x74,
    "VK_F6":                0x75,  "VK_F7":               0x76,  "VK_F8":                  0x77,  "VK_F9":             0x78,  "VK_F10":          0x79,  "VK_F11":    0x7A,
    "VK_F12":               0x7B,  "VK_F13":              0x7C,  "VK_F14":                 0x7D,  "VK_F15":            0x7E,  "VK_F16":          0x7F,  "VK_F17":    0x80,
    "VK_F18":               0x81,  "VK_F19":              0x82,  "VK_F20":                 0x83,  "VK_F21":            0x84,  "VK_F22":          0x85,  "VK_F23":    0x86,
    "VK_F24":               0x87,  "VK_NUMLOCK":          0x90,  "VK_SCROLL":              0x91,  "VK_LSHIFT":         0xA0,  "VK_RSHIFT":       0xA1,
    "VK_LCONTROL":          0xA2,  "VK_RCONTROL":         0xA3,  "VK_LMENU":               0xA4,  "VK_RMENU":          0xA5,  "VK_PROCESSKEY":   0xE5,
    "VK_ATTN":              0xF6,  "VK_CRSEL":            0xF7,  "VK_EXSEL":               0xF8,  "VK_EREOF":          0xF9,  "VK_PLAY":         0xFA,
    "VK_ZOOM":              0xFB,  "VK_NONAME":           0xFC,  "VK_PA1":                 0xFD,  "VK_OEM_CLEAR":      0xFE,  "VK_BROWSER_BACK": 0xA6,
    "VK_BROWSER_FORWARD":   0xA7,  "VK_BROWSER_REFRESH":  0xA8,  "VK_BROWSER_STOP":        0xA9,  "VK_BROWSER_SEARCH": 0xAA,
    "VK_BROWSER_FAVORITES": 0xAB,  "VK_BROWSER_HOME":     0xAC,  "VK_VOLUME_MUTE":         0xAD,  "VK_VOLUME_DOWN":    0xAE,
    "VK_VOLUME_UP":         0xAF,  "VK_MEDIA_NEXT_TRACK": 0xB0,  "VK_MEDIA_PREV_TRACK":    0xB1,  "VK_MEDIA_STOP":     0xB2,
    "VK_MEDIA_PLAY_PAUSE":  0xB3,  "VK_LAUNCH_MAIL":      0xB4,  "VK_LAUNCH_MEDIA_SELECT": 0xB5,  "VK_LAUNCH_APP1":    0xB6,  "VK_LAUNCH_APP2":  0xB7,
    
    "VK_OEM_1":             0xBA,  "VK_OEM_PLUS":         0xBB,  "VK_OEM_COMMA":           0xBC,  "VK_OEM_MINUS":      0xBD,
    "VK_OEM_PERIOD":        0xBE,  "VK_OEM_2":            0xBF,  "VK_OEM_3":               0xC0,  "VK_OEM_4":          0xDB,  "VK_OEM_5":        0xDC,
    "VK_OEM_6":             0xDD,  "VK_OEM_7":            0xDE,  "VK_OEM_8":               0xDF,  "VK_OEM_102":        0xE2,
    
    "VK_PROCESSKEY":        0xE5,  "VK_PACKET":           0xE7,  "VK_FN":                  0xFF}
"""Mapping of virtual key names to key codes. Most of the function or utility keys in this dict are not affected by shift. Only the OEM keys are."""

# Note that OEM_102 have the same character value as VK_OEM_5.
oemCodeToCharName          = {0xBA: ";", 0xBB: "=", 0xBC: ",", 0xBD: "-", 0xBE: ".", 0xBF: "/", 0xC0: "`", 0xDB: "[", 0xDC: "\\", 0xDD: "]", 0xDE: "'", 0xDF: chr(0xDF), 0xE2: "\\"}
"""Mapping of OEM key codes to character names. These are the keys that are affected by shift."""

oemCodeToCharNameWithShift = {0xBA: ":", 0xBB: "+", 0xBC: "<", 0xBD: "_", 0xBE: ">", 0xBF: "?", 0xC0: "~", 0xDB: "{", 0xDC: "|",  0xDD: "}", 0xDE: '"', 0xDF: chr(0xDF), 0xE2: "|"}
"""Mapping of OEM key codes to character names when shift is pressed."""

numRowCodeToSymbol = {48: ")", 49: "!", 50: "@", 51: "#", 52: "$", 53: "%", 54: "^", 55: "&", 56: "*", 57: "("}
"""Mapping of number row key codes to their symbols when shift is pressed."""

# Inverse of previous mapping. Used to convert virtual key codes to their names.
vKeyCodeToName= {v: k for k, v in vKeyNameToId.items()}
"""Mapping of virtual key codes to their names."""

eventIdToName = {
    KB_MsgTypes.WM_KEYDOWN    : "key down",      KB_MsgTypes.WM_KEYUP       : "key up",
    KB_MsgTypes.WM_CHAR       : "key char",      KB_MsgTypes.WM_DEADCHAR    : "key dead char",
    KB_MsgTypes.WM_SYSKEYDOWN : "key sys down",  KB_MsgTypes.WM_SYSKEYUP    : "key sys up",
    KB_MsgTypes.WM_SYSCHAR    : "key sys char",  KB_MsgTypes.WM_SYSDEADCHAR : "key sys dead char"}
"""Mapping of event codes to their names."""

def GetKeyAsciiAndName(vkey_code: int, shiftPressed=False) -> tuple[int, str]:
    """
    Description:
        Returns the ascii value and key name for the given key code.
    ---
    Parameters:
        `vkey_code -> int`: A virtual key code.
        
        `scancode -> int`: The key scancode.
        
        `shiftPressed -> bool`: Whether or not the shift key is pressed.
    
    ---
    Returns:
        `tuple[int, str]`: The ascii and name of the virtual key.
    """
    ...


# Define the KBDLLHOOKSTRUCT structure. Docs: https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-kbdllhookstruct.
class KBDLLHOOKSTRUCT(ctypes.Structure):
    """A structure that contains information about a low-level keyboard input event."""
    
    _fields_ = [
        ("vkCode", ctypes.c_ulong),
        ("scanCode", ctypes.c_ulong),
        ("flags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))
    ]

# By TwhK/Kheldar. Source: http://www.hackerthreads.org/Topic-42395
class HookManager:
    """
    Description:
        A class that manages windows event hooks.
    ---
    Methods:
        `InstallHook(self, callBack: Callable[[int, int, KBDLLHOOKSTRUCT], Any], hookType=HookTypes.WH_KEYBOARD_LL) -> bool`:
            Installs the specified callback hook. Returns `True` if everything was successful, and `False` if it failed.
        
        `BeginListening()`:
            Starts listening for keyboard Windows. Must be called from the main thread.
        
        `UninstallHook()`:
            "Uninstalls the hook specified by the hookId.
    """
    
    def __init__(self):
        self.hookId: int
        self.hookPtr: Callable[[int, int, KBDLLHOOKSTRUCT], Any]

    def InstallHook(self, callBack: Callable[[int, int, KBDLLHOOKSTRUCT], Any], hookType=HookTypes.WH_KEYBOARD_LL) -> bool:
        """Installs the specified hook. Returns True if everything was successful, and False if it failed."""
        ...

    def BeginListening(self):
        """Starts listening for keyboard Windows. Must be called from the main thread."""
        ...

    def UninstallHook(self):
        """Uninstalls the hook specified by the `hookId`."""
        ...

class KeyboardHookManager:
    """A class for managing keyboard hooks and their event listeners."""
    
    def __init__(self):
        self.keyDownListeners: list[Callable[[KeyboardEvent], bool]]
        self.keyUpListeners  : list[Callable[[KeyboardEvent], bool]]
        self.hookId: int

    def addKeyDownListener(self, listener: Callable[[KeyboardEvent], bool]):
        ...

    def addKeyUpListener(self, listener: Callable[[KeyboardEvent], bool]):
        ...

    def removeKeyDownListener(self, listener: Callable[[KeyboardEvent], bool]):
        ...

    def removeKeyUpListener(self, listener: Callable[[KeyboardEvent], bool]):
        ...

    def KeyboardHook(self, nCode: int, wParam: int, lParam: KBDLLHOOKSTRUCT) -> bool:
        """
        Description:
            Processes a low level windows keyboard event.
        ---
        Parameters:
            - `nCode`: The hook code passed to the hook procedure. The value of the hook code depends on the type of hook associated with the procedure.
            - `wParam`: The identifier of the keyboard message (event id). This parameter can be one of the following messages: `WM_KEYDOWN`, `WM_KEYUP`, `WM_SYSKEYDOWN`, or `WM_SYSKEYUP`.
            - `lParam`: A pointer to a `KBDLLHOOKSTRUCT` structure.
        """
        ...
