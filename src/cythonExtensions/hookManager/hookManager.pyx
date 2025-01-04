# cython: language_level = 3str

"""This extension module contains functions and classes for managing Windows hooks."""

cimport cython
from cythonExtensions.commonUtils.commonUtils cimport KeyboardEvent, MouseEvent
from cythonExtensions.hookManager.hookManager cimport HookTypes, KbMsgIds, MsMsgIds, RawMouse

import ctypes, win32gui, win32api, win32con, atexit
import ctypes.wintypes
from cythonExtensions.commonUtils.commonUtils import KB_Con as kbcon, ControllerHouse as ctrlHouse, MouseHouse as msHouse, PThread


cdef dict vKeyNameToId = {
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
        "VK_LCONTROL":          0xA2,  "VK_RCONTROL":         0xA3,  "VK_LMENU":               0xA4,  "VK_RMENU":          0xA5,
        "VK_ATTN":              0xF6,  "VK_CRSEL":            0xF7,  "VK_EXSEL":               0xF8,  "VK_EREOF":          0xF9,  "VK_PLAY":         0xFA,
        "VK_ZOOM":              0xFB,  "VK_NONAME":           0xFC,  "VK_PA1":                 0xFD,  "VK_OEM_CLEAR":      0xFE,  "VK_BROWSER_BACK": 0xA6,
        "VK_BROWSER_FORWARD":   0xA7,  "VK_BROWSER_REFRESH":  0xA8,  "VK_BROWSER_STOP":        0xA9,  "VK_BROWSER_SEARCH": 0xAA,
        "VK_BROWSER_FAVORITES": 0xAB,  "VK_BROWSER_HOME":     0xAC,  "VK_VOLUME_MUTE":         0xAD,  "VK_VOLUME_DOWN":    0xAE,
        "VK_VOLUME_UP":         0xAF,  "VK_MEDIA_NEXT_TRACK": 0xB0,  "VK_MEDIA_PREV_TRACK":    0xB1,  "VK_MEDIA_STOP":     0xB2,
        "VK_MEDIA_PLAY_PAUSE":  0xB3,  "VK_LAUNCH_MAIL":      0xB4,  "VK_LAUNCH_MEDIA_SELECT": 0xB5,  "VK_LAUNCH_APP1":    0xB6,  "VK_LAUNCH_APP2":  0xB7,
        "Reserved1":            0xB8,  "Reserved2":           0xB9,  # Although these two are reserved, one time `0xB8` was sent (before I add it here) for some unknown reason and I couldn't reproduce it.
        "VK_OEM_1":             0xBA,  "VK_OEM_PLUS":         0xBB,  "VK_OEM_COMMA":           0xBC,  "VK_OEM_MINUS":      0xBD,
        "VK_OEM_PERIOD":        0xBE,  "VK_OEM_2":            0xBF,  "VK_OEM_3":               0xC0,  "VK_OEM_4":          0xDB,  "VK_OEM_5":        0xDC,
        "VK_OEM_6":             0xDD,  "VK_OEM_7":            0xDE,  "VK_OEM_8":               0xDF,  "VK_OEM_102":        0xE2,
        
        "VK_PROCESSKEY":        0xE5,  "VK_PACKET":           0xE7,  "VK_FN":                  0xFF}
"""Mapping of virtual key names to key codes. Most of the function or utility keys in this dict are not affected by shift. Only the OEM keys are."""

# Note that OEM_102 have the same character value as VK_OEM_5.
cdef dict oemCodeToCharName          = {0xBA: ";", 0xBB: "=", 0xBC: ",", 0xBD: "-", 0xBE: ".", 0xBF: "/", 0xC0: "`", 0xDB: "[", 0xDC: "\\", 0xDD: "]", 0xDE: "'", 0xDF: chr(0xDF), 0xE2: "\\"}
"""Mapping of OEM key codes to character names. These are the keys that are affected by shift."""

cdef dict oemCodeToCharNameWithShift = {0xBA: ":", 0xBB: "+", 0xBC: "<", 0xBD: "_", 0xBE: ">", 0xBF: "?", 0xC0: "~", 0xDB: "{", 0xDC: "|",  0xDD: "}", 0xDE: '"', 0xDF: chr(0xDF), 0xE2: "|"}
"""Mapping of OEM key codes to character names when shift is pressed."""

cdef dict numRowCodeToSymbol = {48: ")", 49: "!", 50: "@", 51: "#", 52: "$", 53: "%", 54: "^", 55: "&", 56: "*", 57: "("}
"""Mapping of number row key codes to their symbols when shift is pressed."""

# Inverse of previous mapping. Used to convert virtual key codes to their names.
cdef dict vKeyCodeToName = {v: k for k, v in vKeyNameToId.items()}
"""Mapping of virtual key codes to their names."""

cdef dict numPadCodeToName = {
    0x60: "0", 0x61: "1", 0x62: "2",     0x63: "3", 0x64: "4",
    0x65: "5", 0x66: "6", 0x67: "7",     0x68: "8", 0x69: "9",
    0x6A: "*", 0x6B: "+", 0x6C: "ENTER", # 0x6C is the ENTER key on the number pad. It has the same code as the regular ENTER key.
    0x6D: "-", 0x6E: ".", 0x6F: "/"
}
"""Mapping of number pad key codes to their names."""

cdef dict kbMsgIdToName = {
    KbMsgIds.WM_KEYDOWN    : "key down",      KbMsgIds.WM_KEYUP       : "key up",
    KbMsgIds.WM_CHAR       : "key char",      KbMsgIds.WM_DEADCHAR    : "key dead char",
    KbMsgIds.WM_SYSKEYDOWN : "key sys down",  KbMsgIds.WM_SYSKEYUP    : "key sys up",
    KbMsgIds.WM_SYSCHAR    : "key sys char",  KbMsgIds.WM_SYSDEADCHAR : "key sys dead char"
}
"""Maps the Windows Message Ids for keyboard events to their respective names."""

@cython.wraparound(False)
cdef tuple[int, object] getKeyAsciiAndName(int vkey_code, bint shiftPressed=False):
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
    
    # If the key is a number, return its ascii and string representation with respect to `shiftPressed`.
    if (0x30 <= vkey_code <= 0x39):
        if shiftPressed:
            text = numRowCodeToSymbol[vkey_code]
            return (ord(text), text)
        
        # VK_0 : VK_9 have the same code as the ASCII of "0" : "9" (0x30 : 0x39).
        return (vkey_code, chr(vkey_code))
    
    # A number pad key.
    elif 0x60 <= vkey_code <= 0x6F:
        return (vkey_code, numPadCodeToName[vkey_code])
    
    # The key is letter.
    elif 0x41 <= vkey_code <= 0x5A:
        if shiftPressed or win32api.GetKeyState(win32con.VK_CAPITAL): # A capital letter.
            # VK_A : VK_Z have the same code as the ASCII of "A" : "Z" (0x41 : 0x5A).
            return (vkey_code, chr(vkey_code))
        
        # VK_a : VK_z have an ASCII value equal to the virtual code + 32.
        return (vkey_code + 32, chr(vkey_code + 32))
    
    # The key is not a letter or number. It must be one of the other VK_ keys.
    else:
        text = vKeyCodeToName[vkey_code]
        
        if text.startswith("VK_OEM_"):
            if shiftPressed:
                text = oemCodeToCharNameWithShift[vkey_code]
                return (ord(text), text)
            
            text = oemCodeToCharName[vkey_code]
            return (ord(text), text)
        
        # A function or utility key.
        else: # text.startswith("VK_"):
            return (win32api.MapVirtualKey(vkey_code, 2), text[3:].title())


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
cdef class HookManager:
    """A class that manages windows event hooks."""
    
    cdef int kbHookId, msHookId
    cdef kbHookPtr, msHookPtr
    
    def __init__(self):
        self.kbHookId = 0
        self.kbHookPtr = None
        self.msHookId = 0
        self.msHookPtr = None
    
    cdef bint installHook(self, callBack, int hookType=HookTypes.WH_KEYBOARD_LL):
        """Installs a hook of the specified type. `hookType` can be one of the following:
        Value            | Receive messsages for
        -----------------|----------------------
        `WH_KEYBOARD_LL` | `Keyboard`
        `WH_MOUSE_LL`    | `Mouse`
        """
        
        if hookType not in (HookTypes.WH_KEYBOARD_LL, HookTypes.WH_MOUSE_LL):
            print("Error, the hookType is not recognized.")
            
            return False
        
        # Defining a type signature for the given low level hook/handler.
        CMPFUNC = ctypes.CFUNCTYPE(ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.POINTER(ctypes.c_void_p))
        
        # Configuring the windows hook argtypes for 64-bit Python compatibility.
        ctypes.windll.user32.SetWindowsHookExW.argtypes = (
            ctypes.c_int,
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.c_uint
        )
        
        if (hookType == HookTypes.WH_KEYBOARD_LL and self.kbHookId) or (hookType == HookTypes.WH_MOUSE_LL and self.msHookId):
            print(f"Warning: A {'keyboard' if hookType == HookTypes.WH_KEYBOARD_LL else 'mouse'} hook is already installed.")
            
            return False
        
        # Converting the Python hook into a C pointer.
        cdef callbackPtr = CMPFUNC(callBack)
        
        # Setting a windows hook with the given hook type.
        cdef int callbackHookId = ctypes.windll.user32.SetWindowsHookExW( # SetWindowsHookExA
            hookType,                       # Hook type.
            callbackPtr,                    # Callback pointer.
            win32gui.GetModuleHandle(None), # Handle to the current process.
            0)                              # Thread id (0 = current/main thread).
        
        # Check if the hook was installed successfully.
        if not callbackHookId:
            print("Error, the hook could not be installed.")
            
            return False
        
        # Register a function that is called upon termination (when the interpreter exits) to remove the hook.
        # Note that methods registered using atexit will be called in reverse order of registration.
        # Also, the registered functions are not called if the interpreter is terminated by a signal not handled by Python,
        # a Python fatal internal error is detected, or when os._exit() is called.
        atexit.register(ctypes.windll.user32.UnhookWindowsHookEx, callbackHookId)
        
        if hookType == HookTypes.WH_KEYBOARD_LL:
            self.kbHookId = callbackHookId
            self.kbHookPtr = callbackPtr
        
        else:
            self.msHookId = callbackHookId
            self.msHookPtr = callbackPtr
        
        return True
    
    cdef bint beginListening(self):
        """Starts listening for Windows events. Must be called from the main thread."""
        
        if not PThread.inMainThread():
            raise RuntimeError("Error! This method can only be called from the main thread.")
        
        if not self.kbHookId and not self.msHookId:
            print("Cannot start receiving event messages as no hook is installed yet.")
            
            return False
        
        # Creating a message loop to keep the hook alive.
        msg = ctypes.wintypes.MSG()
        
        # If wMsgFilterMin and wMsgFilterMax are both zero, GetMessage returns all available messages (that is, no range filtering is performed).
        # Keep in mind that any application that wishes to receive notifications of windows events must have a message queue.
        # Also, each thread has its own message queue, and, as far as I know, the GetMessage can only receive System messages from the main thread.
        # Docs: https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getmessagew#:~:text=the%20WM_INPUT%20messages.-,If%20wMsgFilterMin%20and%20wMsgFilterMax%20are%20both%20zero%2C%20GetMessage%20returns%20all%20available%20messages%20(that%20is%2C%20no%20range%20filtering%20is%20performed).,-%5Bin%5D%20wMsgFilterMax
        while ctypes.windll.user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
            ctypes.windll.user32.TranslateMessage(ctypes.byref(msg))
            ctypes.windll.user32.DispatchMessageW(ctypes.byref(msg))
        return True
    
    cdef bint uninstallHook(self, int hookType=HookTypes.WH_KEYBOARD_LL):
        """Uninstalls the hook specified by the hook type:
        Value            | Receive messsages for
        -----------------|----------------------
        `WH_KEYBOARD_LL` | `Keyboard`
        `WH_MOUSE_LL`    | `Mouse`
        """
        
        if hookType not in (HookTypes.WH_KEYBOARD_LL, HookTypes.WH_MOUSE_LL):
            print("Error, the hook type is not recognized.")
            
            return False
        
        if (hookType == HookTypes.WH_KEYBOARD_LL and not self.kbHookId) or (hookType == HookTypes.WH_MOUSE_LL and not self.msHookId):
            print(f"Warning: No {'keyboard' if hookType == HookTypes.WH_KEYBOARD_LL else 'mouse'} hook has been installed yet.")
            
            return False
        
        if hookType == HookTypes.WH_KEYBOARD_LL:
            ctypes.windll.user32.UnhookWindowsHookEx(self.kbHookId)
            self.kbHookId = 0
            
            # Unregister the function that was registered using atexit.
            # Note that if the function given to `atexit.unregister` has been registered more than once, every occurrence
            # of that function in the atexit call stack will be removed, as equality comparison (==) is used internally.
            if not self.msHookId:
                atexit.unregister(ctypes.windll.user32.UnhookWindowsHookEx)
        
        else:
            ctypes.windll.user32.UnhookWindowsHookEx(self.msHookId)
            self.msHookId = 0
            
            if not self.kbHookId:
                atexit.unregister(ctypes.windll.user32.UnhookWindowsHookEx)
        
        return True

# ctypedef bint (*keyboardCallbackPtr)(KeyboardEvent)

cdef class KeyboardHookManager:
    """A class for managing keyboard hooks and their event listeners."""
    
    cdef public list keyDownListeners, keyUpListeners
    cdef public int hookId
    
    def __init__(self):
        self.keyDownListeners = []
        self.keyUpListeners = []
        self.hookId = 0
    
    cpdef bint keyboardCallback(self, int nCode, int wParam, lParam):
        """
        Description:
            Processes a low level windows keyboard event.
        ---
        Parameters:
            - `nCode`: The hook code passed to the hook procedure. The value of the hook code depends on the type of hook associated with the procedure.
            - `wParam`: The identifier of the keyboard message (event id). This parameter can be one of the following messages: `WM_KEYDOWN`, `WM_KEYUP`, `WM_SYSKEYDOWN`, or `WM_SYSKEYUP`.
            - `lParam`: A pointer to a `KBDLLHOOKSTRUCT` structure.
        """
        
        cdef int vkey_code, scancode, flags
        cdef bint injected, extended, transition, shiftPressed, altPressed, suppressKeyPress
        cdef KeyboardEvent keyboardEvent
        
        suppressKeyPress = False
        
        # Checking if the event is valid. Docs: https://stackoverflow.com/questions/64449078/c-keyboard-hook-what-does-the-parameter-ncode-mean
        if nCode == win32con.HC_ACTION:
            # Casting lParam to KBDLLHOOKSTRUCT.
            lParamStruct_ptr = ctypes.cast(lParam, ctypes.POINTER(KBDLLHOOKSTRUCT))
            lParamStruct = lParamStruct_ptr.contents
            
            # Accessing fields of KBDLLHOOKSTRUCT.
            vkey_code = lParamStruct.vkCode
            
            # Check if the key code is valid.
            if not vkey_code:
                # Note that, if you returned a boolean value instead of calling CallNextHookEx, you are
                # effectively preventing any further processing of the keystroke by other hooks in the chain.
                return ctypes.windll.user32.CallNextHookEx(None, nCode, wParam, lParam)
            
            scancode = lParamStruct.scanCode
            
            # eventTime = lParamStruct.time
            flags = lParamStruct.flags
            
            # Extracting key state flags from the packed int `flags'.
            # Docs: https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-kbdllhookstruct#:~:text=The%20following%20table%20describes%20the%20layout%20of%20this%20value.
            extended   = bool(flags & 0x1)
            injected   = bool(flags & 0x10)
            altPressed = bool(flags & 0x20)
            transition = bool(flags & 0x82)
            
            # To get the correct key ascii value, we need first to check if the shift is pressed.
            shiftPressed = ((win32api.GetAsyncKeyState(win32con.VK_SHIFT) & 0x8000) >> 15) | (vkey_code in (win32con.VK_LSHIFT, win32con.VK_RSHIFT))
            keyAscii, keyName= getKeyAsciiAndName(vkey_code, shiftPressed)
            
            eventName = kbMsgIdToName[wParam]
            
            # Creating a keyboard event object.
            keyboardEvent = KeyboardEvent(event_id=wParam, event_name=eventName, vkey_code=vkey_code,
                                          scancode=scancode, key_ascii=keyAscii, key_name=keyName,
                                          flags=flags, injected=injected, extended=extended,
                                          shift=shiftPressed, alt=altPressed, transition=transition)
            
            # Key down/press event.
            if wParam in (win32con.WM_KEYDOWN, win32con.WM_SYSKEYDOWN):
                # Update the state of the lock keys to reflect their current state.
                ctrlHouse.CAPITAL = win32api.GetKeyState(win32con.VK_CAPITAL)
                ctrlHouse.SCROLL  = win32api.GetKeyState(win32con.VK_SCROLL)
                ctrlHouse.NUMLOCK = win32api.GetKeyState(win32con.VK_NUMLOCK)
                
                #! Distinguish between real user input and keyboard input generated by programs/scripts.
                if not injected:
                    # Update the state of the modifier keys to reflect the current state of being pressed.
                    ctrlHouse.modifiers |= (
                        ((keyboardEvent.KeyID == win32con.VK_LCONTROL) | (keyboardEvent.KeyID == win32con.VK_RCONTROL)) << 13 | # CTRL
                        ((keyboardEvent.KeyID == win32con.VK_LSHIFT)   | (keyboardEvent.KeyID == win32con.VK_RSHIFT))   << 10 | # SHIFT
                        ((keyboardEvent.KeyID == win32con.VK_LMENU)    | (keyboardEvent.KeyID == win32con.VK_RMENU))    << 7  | # ALT
                        ((keyboardEvent.KeyID == win32con.VK_LWIN)     | (keyboardEvent.KeyID == win32con.VK_RWIN))     << 4  | # WIM
                        
                        # (keyboardEvent.KeyID == win32con.VK_LCONTROL) << 12 | # LCTRL
                        # (keyboardEvent.KeyID == win32con.VK_RCONTROL) << 11 | # RCTRL
                        # (keyboardEvent.KeyID == win32con.VK_LSHIFT)   << 9  | # LSHIFT
                        # (keyboardEvent.KeyID == win32con.VK_RSHIFT)   << 8  | # RSHIFT
                        # (keyboardEvent.KeyID == win32con.VK_LMENU)    << 6  | # LALT
                        # (keyboardEvent.KeyID == win32con.VK_RMENU)    << 5  | # RALT
                        # (keyboardEvent.KeyID == win32con.VK_LWIN)     << 3  | # LWIN
                        # (keyboardEvent.KeyID == win32con.VK_RWIN)     << 2  | # RWIN
                        
                        (keyboardEvent.KeyID == 255)                    << 1  | # FN
                        (keyboardEvent.KeyID == kbcon.VK_BACKTICK)              # BACKTICK
                    )
                    
                    # Propagate the event to the registered keyDown listeners.
                    for listener in self.keyDownListeners:
                        PThread(target=listener, args=(keyboardEvent,)).start()
                    
                    # One of the keyDown listeners must put a signal in the PThread.kbMsgQueue to specify whether to return or suppress the pressed key.
                    if self.keyDownListeners and PThread.kbMsgQueue.get():
                        suppressKeyPress = True
            
            # Key up event.
            else:
                if not injected:
                    # Update the state of the modifier keys to reflect their current state of being released.
                    ctrlHouse.modifiers &= ~(
                        ((keyboardEvent.KeyID == win32con.VK_LCONTROL) | (keyboardEvent.KeyID == win32con.VK_RCONTROL)) << 13 | # CTRL
                        ((keyboardEvent.KeyID == win32con.VK_LSHIFT)   | (keyboardEvent.KeyID == win32con.VK_RSHIFT))   << 10 | # SHIFT
                        ((keyboardEvent.KeyID == win32con.VK_LMENU)    | (keyboardEvent.KeyID == win32con.VK_RMENU))    << 7  | # ALT
                        ((keyboardEvent.KeyID == win32con.VK_LWIN)     | (keyboardEvent.KeyID == win32con.VK_RWIN))     << 4  | # WIN
                        
                        # (1 << 13) | (1 << 10) | (1 << 7) | (1 << 4) | # Reseting CTRL, SHIFT, ALT, WIN
                        # (keyboardEvent.KeyID == win32con.VK_LCONTROL) << 12 | # LCTRL
                        # (keyboardEvent.KeyID == win32con.VK_RCONTROL) << 11 | # RCTRL
                        # (keyboardEvent.KeyID == win32con.VK_LSHIFT)   << 9  | # LSHIFT
                        # (keyboardEvent.KeyID == win32con.VK_RSHIFT)   << 8  | # RSHIFT
                        # (keyboardEvent.KeyID == win32con.VK_LMENU)    << 6  | # LALT
                        # (keyboardEvent.KeyID == win32con.VK_RMENU)    << 5  | # RALT
                        # (keyboardEvent.KeyID == win32con.VK_LWIN)     << 3  | # LWIN
                        # (keyboardEvent.KeyID == win32con.VK_RWIN)     << 2  | # RWIN
                        
                        (keyboardEvent.KeyID == 255)                    << 1  | # FN
                        (keyboardEvent.KeyID == kbcon.VK_BACKTICK)              # BACKTICK
                    )
                    
                    # Propagate the event to the registered keyUp listeners.
                    for listener in self.keyUpListeners:
                        PThread(target=listener, args=(keyboardEvent,)).start()
            
            ctypes.windll.user32.CallNextHookEx(None, nCode, wParam, lParam)
            
            return suppressKeyPress

# ======================================================================================================================

cdef dict msMsgIdToName = {
    MsMsgIds.WM_MOUSEMOVE       : "MOVE",
    MsMsgIds.WM_LBUTTONDOWN     : "LB CLK",
    MsMsgIds.WM_LBUTTONUP       : "LB UP",
    MsMsgIds.WM_LBUTTONDBLCLK   : "LB DBL CLK",
    MsMsgIds.WM_RBUTTONDOWN     : "RB CLK",
    MsMsgIds.WM_RBUTTONUP       : "RB UP",
    MsMsgIds.WM_RBUTTONDBLCLK   : "RB DBL CLK",
    MsMsgIds.WM_MBUTTONDOWN     : "MB CLK",
    MsMsgIds.WM_MBUTTONUP       : "MB UP",
    MsMsgIds.WM_MBUTTONDBLCLK   : "MB DBL CLK",
    MsMsgIds.WM_MOUSEWHEEL      : "WHEEL SCRL",
    MsMsgIds.WM_XBUTTONDOWN     : "XB CLK",
    MsMsgIds.WM_XBUTTONUP       : "XB UP",
    MsMsgIds.WM_XBUTTONDBLCLK   : "XB DBL CLK",
    MsMsgIds.WM_MOUSEHWHEEL     : "WHEEL H SCRL",
    MsMsgIds.WM_NCXBUTTONDOWN   : "NC XB CLK",
    MsMsgIds.WM_NCXBUTTONUP     : "NC XB UP",
    MsMsgIds.WM_NCXBUTTONDBLCLK : "NC XB DBL CLK",
}
"""Maps the Windows Message Ids for mouse events to their respective names."""


# Docs: https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-msllhookstruct
class MSLLHOOKSTRUCT(ctypes.Structure):
    """
    Description:
        A structure that contains information about a low-level mouse input event.
    ---
    Fields:
        'pt: POINT`:
            The x- and y-coordinates of the cursor, in screen coordinates.
        
        `mouseData: DWORD`:
            Contains additional information about the mouse event. Can have one of the following values depending on the event type:
            Mouse Wheel | Meaning
            ------------|---------
            `+ve value` | The mouse wheel was scrolled forward/upward.
            `-ve value` | The mouse wheel was scrolled backward/downward.
            `Note`      | Each unit of scrolling typically corresponds to a fixed value (e.g., 120):
            
            Mouse Button | Meaning
            -------------|---------
            `0x0000` | No mouse button change
            `0x0001` | Left mouse button was pressed or released.
            `0x0002` | Right mouse button was pressed or released.
            `0x0004` | Middle mouse button was pressed or released.
        
        `flags: DWORD`:
            Whether the event was injected. The value is `1` if it was injected; otherwise, it is `0`.
        
        `time: DWORD`:
            The time stamp for this message.
        
        `dwExtraInfo: ULONG_PTR`:
            An additional value associated with the mouse event.
    """
    
    _fields_ = [
        ("pt", ctypes.wintypes.POINT),
        ("mouseData", ctypes.c_ulong),
        ("flags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))
    ]


cdef class MouseHookManager:
    """A class for managing mouse hooks and their event listeners."""
    
    cdef public list mouseButtonDownListeners, mouseButtonUpListeners
    cdef public int hookId
    
    def __init__(self):
        self.mouseButtonDownListeners = []
        self.mouseButtonUpListeners = []
        self.hookId = 0
    
    cpdef mouseCallback(self, int nCode, int wParam, lParam):
        """
        Description:
            Processes a low level windows mouse event.
        ---
        Parameters:
            - `nCode`: The hook code passed to the hook procedure. The value of the hook code depends on the type of hook associated with the procedure.
            - `wParam`: The identifier of the mouse message (event id).
            - `lParam`: A pointer to a `MSLLHOOKSTRUCT` structure.
        """
        
        # Filter out mouse move events.
        if wParam == MsMsgIds.WM_MOUSEMOVE:
            return ctypes.windll.user32.CallNextHookEx(None, nCode, wParam, lParam)
        
        cdef int x, y, flags, pressedButton, wheelDelta # time, dwExtraInfo,
        cdef bint isMouseAbsolute, isMouseInWindow, isWheelHorizontal, suppressInput
        cdef MouseEvent mouseEvent
        
        suppressInput = False
        if nCode == win32con.HC_ACTION:
            lParamStruct_ptr = ctypes.cast(lParam, ctypes.POINTER(MSLLHOOKSTRUCT))
            lParamStruct = lParamStruct_ptr.contents
            
            x = lParamStruct.pt.x
            y = lParamStruct.pt.y
            mouseData = lParamStruct.mouseData
            flags = lParamStruct.flags
            # time = datetime.datetime.fromtimestamp(ctypes.c_long(lParamStruct.time).value).astimezone()
            # dwExtraInfo = lParamStruct.dwExtraInfo
            
            isMouseAbsolute = bool(flags & RawMouse.MOUSE_MOVE_ABSOLUTE)
            isMouseInWindow = bool(flags & RawMouse.MOUSE_MOVE_NOCOALESCE)
            isWheelHorizontal = wParam == MsMsgIds.WM_MOUSEHWHEEL
            
            # isLeftButtonPressed   = wParam == win32con.WM_LBUTTONDOWN    # wParam == win32con.WM_LBUTTONUP
            # isRightButtonPressed  = wParam == win32con.WM_RBUTTONDOWN    # wParam == win32con.WM_RBUTTONUP
            # isMiddleButtonPressed = wParam == win32con.WM_MBUTTONDOWN    # wParam == win32con.WM_MBUTTONUP
            # isXButton1Pressed, isXButton2Pressed = (wParam == MsMsgIds.WM_XBUTTONDOWN, False) if (mouseData >> 16 == 1) else (False, wParam == MsMsgIds.WM_XBUTTONDOWN)# wParam == WM_XBUTTONUP
            
            pressedButton = (
                (wParam in (MsMsgIds.WM_LBUTTONDOWN, MsMsgIds.WM_LBUTTONUP)) << 4 |
                (wParam in (MsMsgIds.WM_RBUTTONDOWN, MsMsgIds.WM_RBUTTONUP)) << 3 |
                (wParam in (MsMsgIds.WM_MBUTTONDOWN, MsMsgIds.WM_MBUTTONUP)) << 2 |
                (wParam in (MsMsgIds.WM_XBUTTONDOWN, MsMsgIds.WM_XBUTTONUP)) << (1 if (mouseData >> 16 == 1) else 0) # (mouseData >> 17 == 1)
            )
            
            wheelDelta = 0
            if wParam in (MsMsgIds.WM_MOUSEWHEEL, MsMsgIds.WM_MOUSEHWHEEL):
                wheelDelta = ctypes.c_short((mouseData >> 16) & 0xFFFF).value
            
            mouseEvent = MouseEvent(
                event_id=wParam,
                event_name=msMsgIdToName[wParam],
                flags=flags,
                x=x,
                y=y,
                mouse_data=ctypes.c_int(mouseData).value,
                is_mouse_absolute=isMouseAbsolute,
                is_mouse_in_window=isMouseInWindow,
                wheel_delta=wheelDelta,
                is_wheel_horizontal=isWheelHorizontal,
                pressed_button=pressedButton
            )
            
            # Button down/press event.
            if wParam in (MsMsgIds.WM_LBUTTONDOWN, MsMsgIds.WM_RBUTTONDOWN, MsMsgIds.WM_MBUTTONDOWN, MsMsgIds.WM_XBUTTONDOWN, MsMsgIds.WM_NCXBUTTONDOWN, MsMsgIds.WM_MOUSEWHEEL, MsMsgIds.WM_MOUSEHWHEEL):
                # Updating the state of the mouse for the button down and wheel movement events.
                msHouse.delta = mouseEvent.Delta
                
                msHouse.horizontal = mouseEvent.IsWheelHorizontal
                
                msHouse.buttons |= (
                    (mouseEvent.PressedButton == msHouse.LButton)  << 4 |
                    (mouseEvent.PressedButton == msHouse.RButton)  << 3 |
                    (mouseEvent.PressedButton == msHouse.MButton)  << 2 |
                    (mouseEvent.PressedButton == msHouse.X1Button) << 1 |
                    (mouseEvent.PressedButton == msHouse.X2Button)
                )
                
                
                # Propagateing the event to the registered butDown listeners.
                for listener in self.mouseButtonDownListeners:
                    PThread(target=listener, args=(mouseEvent,)).start()
                
                # One of the buttonDown listeners must put a signal in the PThread.msMsgQueue to specify whether to return or suppress the mouse input.
                if self.mouseButtonDownListeners and PThread.msMsgQueue.get():
                    suppressInput = True
            
            # Button up event.
            else:
                # Updating the state of the mouse for the button up, mouse movement, and other events.
                msHouse.x, msHouse.y = mouseEvent.X, mouseEvent.Y
                
                msHouse.delta = mouseEvent.Delta
                
                msHouse.horizontal = mouseEvent.IsWheelHorizontal
                
                msHouse.buttons &= ~(
                    (mouseEvent.PressedButton == msHouse.LButton)  << 4 |
                    (mouseEvent.PressedButton == msHouse.RButton)  << 3 |
                    (mouseEvent.PressedButton == msHouse.MButton)  << 2 |
                    (mouseEvent.PressedButton == msHouse.X1Button) << 1 |
                    (mouseEvent.PressedButton == msHouse.X2Button)
                )
                
                # Propagating the event to the registered buttonUp listeners.
                for listener in self.mouseButtonUpListeners:
                    PThread(target=listener, args=(mouseEvent,)).start()
        
        ctypes.windll.user32.CallNextHookEx(None, nCode, wParam, lParam)
        
        return suppressInput
