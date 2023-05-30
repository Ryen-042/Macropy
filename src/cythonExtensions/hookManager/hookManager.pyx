# cython: embedsignature = True
# cython: language_level = 3str

import ctypes, win32gui, win32api, win32con, atexit
import ctypes.wintypes
from cythonExtensions.commonUtils.commonUtils cimport KeyboardEvent
from cythonExtensions.commonUtils.commonUtils import PThread, KeyboardEvent

# https://learn.microsoft.com/en-us/windows/win32/winmsg/about-hooks
cdef enum HookTypes:
    # Constants that represent different types of Windows hooks that can be set using the SetWindowsHookEx function.
    WH_MSGFILTER       = -1   # A hook type that allows you to monitor messages sent to a window or dialog box procedure. This is also the minmum value for a hook type.
    WH_JOURNALRECORD   = 0    # A hook type that is used to record input events.
    WH_JOURNALPLAYBACK = 1    # A hook type that is used to replay the input events recorded by the WH_JOURNALRECORD hook type.
    WH_KEYBOARD        = 2    # A hook type that is used to monitor keystrokes. It can be used to implement a keylogger or to perform other keyboard-related tasks.
    WH_GETMESSAGE      = 3    # A hook type that is used to monitor messages in the message queue of a thread. It can be used to modify or filter messages before they are dispatched.
    WH_CALLWNDPROC     = 4    # A hook type that is used to monitor messages sent to a window or dialog box procedure before they are processed by the procedure.
    WH_CBT             = 5    # A hook type that is used to monitor a variety of system events, including window creation and destruction, window activation, and window focus changes.
    WH_SYSMSGFILTER    = 6    # A hook type that is similar to WH_MSGFILTER, but it monitors messages sent to all windows, not just a specific window or dialog box procedure.
    WH_HARDWARE        = 8    # A hook type that is used to monitor low-level hardware events, such as mouse and keyboard input.
    WH_DEBUG           = 9    # A hook type that is used for debugging purposes.
    WH_SHELL           = 10   # A hook type that is used to monitor shell-related events, such as when the shell is about to activate a new window.
    WH_FOREGROUNDIDLE  = 11   # A hook type that is used to monitor when the system enters an idle state and the foreground application is not processing input.
    WH_CALLWNDPROCRET  = 12   # A hook type that is similar to WH_CALLWNDPROC, but it monitors messages after they have been processed by the procedure.
    WH_KEYBOARD_LL     = 13   # A hook type that is similar to WH_KEYBOARD, but it is a low-level keyboard hook that can be used to monitor keystrokes from all processes.
    WH_MAX             = 15   # The maximum value for a hook type. It is not a valid hook type itself.

cdef enum KB_MsgTypes:
    # Constants that represent different types of keyboard-related Windows messages that can be received by a window or a message loop.
    WM_KEYDOWN     = 0x0100   # A keyboard key was pressed.
    WM_KEYUP       = 0x0101   # A keyboard key was released.
    WM_CHAR        = 0x0102   # A keyboard key was pressed down and released, and it represents a printable character.
    WM_DEADCHAR    = 0x0103   # A keyboard key was pressed down and released, and it represents a dead character.
    WM_SYSKEYDOWN  = 0x0104   # A keyboard key was pressed while the ALT key was held down.
    WM_SYSKEYUP    = 0x0105   # A keyboard key was released after the ALT key was held down.
    WM_SYSCHAR     = 0x0106   # A keyboard key was pressed down and released, and it represents a printable character while the ALT key was held down.
    WM_SYSDEADCHAR = 0x0107   # A keyboard key was pressed down and released, and it represents a dead character while the ALT key was held down.
    WM_KEYLAST     = 0x0108   # Defines the maximum value for the range of keyboard-related messages.

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
cdef dict vKeyCodeToName= {v: k for k, v in vKeyNameToId.items()}
"""Mapping of virtual key codes to their names."""

cdef dict eventIdToName = {
    KB_MsgTypes.WM_KEYDOWN    : "key down",      KB_MsgTypes.WM_KEYUP       : "key up",
    KB_MsgTypes.WM_CHAR       : "key char",      KB_MsgTypes.WM_DEADCHAR    : "key dead char",
    KB_MsgTypes.WM_SYSKEYDOWN : "key sys down",  KB_MsgTypes.WM_SYSKEYUP    : "key sys up",
    KB_MsgTypes.WM_SYSCHAR    : "key sys char",  KB_MsgTypes.WM_SYSDEADCHAR : "key sys dead char"}
"""Mapping of event codes to their names."""

cdef inline tuple GetKeyAsciiAndName(int vkey_code, bint shiftPressed=False):
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

# ctypedef bint (*HookPtr)(int, int, KBDLLHOOKSTRUCT)


# By TwhK/Kheldar. Source: http://www.hackerthreads.org/Topic-42395
cdef class HookManager:
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
    
    cdef int hookId
    cdef hookPtr
    # cdef HookPtr hook_ptr
    
    def __init__(self):
        self.hookId = 0
        self.hookPtr = None
    
    cpdef bint InstallHook(self, callBack, int hookType=HookTypes.WH_KEYBOARD_LL):
        """Installs the specified hook. Returns True if everything was successful, and False if it failed."""
        
        # Defining a type signature for the given low level hook/handler.
        CMPFUNC = ctypes.CFUNCTYPE(ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.POINTER(ctypes.c_void_p))
        
        # Configuring the windows hook argtypes for 64-bit Python compatibility.
        ctypes.windll.user32.SetWindowsHookExW.argtypes = (
            ctypes.c_int,
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.c_uint
        )
        
        # Converting the Python hook into a C pointer.
        self.hookPtr = CMPFUNC(callBack)
        
        # Setting the windows hook with the given hook.
        # Hook both key up and key down events for common keys (non-system).
        self.hookId = ctypes.windll.user32.SetWindowsHookExW( # SetWindowsHookExA
            hookType,                       # Hook type.
            self.hookPtr,                   # Callback pointer.
            win32gui.GetModuleHandle(None), # Handle to the current process.
            0)                              # Thread id (0 = current/main thread).
        
        # Check if the hook was installed successfully.
        if not self.hookId:
            return False
        
        # Register a function that is called upon termination (when the interpreter exits) to remove the hook.
        # Note that methods registered using atexit will be called in reverse order of registration.
        # Also, the registered functions are not called if the interpreter is terminated by a signal not handled by Python,
        # a Python fatal internal error is detected, or when os._exit() is called.
        atexit.register(ctypes.windll.user32.UnhookWindowsHookEx, self.hookId)
        
        return True
    
    cpdef void BeginListening(self):
        """Starts listening for keyboard Windows. Must be called from the main thread."""
        
        if not PThread.InMainThread():
            raise RuntimeError("Waranning! This method can only be called from the main thread.")
        
        if self.hookId is None or not self.hookId:
            "Warning: No hook is installed yet."
            return
        
        # Creating a message loop to keep the hook alive.
        msg = ctypes.wintypes.MSG()
        
        # If wMsgFilterMin and wMsgFilterMax are both zero, GetMessage returns all available messages (that is, no range filtering is performed).
        # Keep in mind that any application that wishes to receive notifications of windows events must have a message queue.
        # Also, each thread has its own message queue, and, as far as I know, the GetMessage can only receive System messages from the main thread.
        # Docs: https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getmessagew#:~:text=the%20WM_INPUT%20messages.-,If%20wMsgFilterMin%20and%20wMsgFilterMax%20are%20both%20zero%2C%20GetMessage%20returns%20all%20available%20messages%20(that%20is%2C%20no%20range%20filtering%20is%20performed).,-%5Bin%5D%20wMsgFilterMax
        while ctypes.windll.user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
            ctypes.windll.user32.TranslateMessage(ctypes.byref(msg))
            ctypes.windll.user32.DispatchMessageW(ctypes.byref(msg))
    
    cpdef bint UninstallHook(self):
        """Uninstalls the hook specified by the `hookId`."""

        if self.hookId is None or not self.hookId:
            print("Warning: No hook is installed yet.")
            return False
        
        ctypes.windll.user32.UnhookWindowsHookEx(self.hookId)
        
        # Unregister the function that was registered using atexit.
        # Note that if the function given to `atexit.unregister` has been registered more than once, every occurrence
        # of that function in the atexit call stack will be removed, as equality comparison (==) is used internally.
        atexit.unregister(ctypes.windll.user32.UnhookWindowsHookEx)
        
        self.hookId = 0
        
        return True

# ctypedef bint (*EventListenerPtr)(KeyboardEvent)

cdef class KeyboardHookManager:
    """A class for managing keyboard hooks and their event listeners."""
    
    cdef public list keyDownListeners, keyUpListeners
    cdef public int hookId
    
    def __init__(self):
        self.keyDownListeners = []
        self.keyUpListeners = []
        self.hookId = 0

    cpdef void addKeyDownListener(self, listener):
        self.keyDownListeners.append(listener)

    cpdef void addKeyUpListener(self, listener):
        self.keyUpListeners.append(listener)

    cpdef void removeKeyDownListener(self, listener):
        self.keyDownListeners.remove(listener)

    cpdef removeKeyUpListener(self, listener):
        self.keyUpListeners.remove(listener)

    cpdef bint KeyboardHook(self, int nCode, int wParam, lParam):
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
        cdef bint injected, extended, transition, shiftPressed, altPressed
        cdef KeyboardEvent keyboardEvent
        
        # Checking if the event is valid. Docs: https://stackoverflow.com/questions/64449078/c-keyboard-hook-what-does-the-parameter-ncode-mean
        if nCode == win32con.HC_ACTION:
            # Casting lParam to KBDLLHOOKSTRUCT.
            lParamStruct_ptr = ctypes.cast(lParam, ctypes.POINTER(KBDLLHOOKSTRUCT))
            lParamStruct = lParamStruct_ptr.contents
            
            # Accessing fields of KBDLLHOOKSTRUCT.
            vkey_code = lParamStruct.vkCode
            
            # Check if the key code is valid.
            if not vkey_code:
                return False
            
            scancode = lParamStruct.scanCode
            
            # eventTime = lParamStruct.time
            flags = lParamStruct.flags
            
            # Extracting key state flags from the packed int `flags`.
            # Docs: https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-kbdllhookstruct#:~:text=The%20following%20table%20describes%20the%20layout%20of%20this%20value.
            extended = bool(flags & 0x1)
            injected = bool(flags & 0x10)
            altPressed = bool(flags & 0x20)
            transition = bool(flags & 0x82) # Always `False`?
            
            # To get the correct key ascii value, we need first to check if the shift is pressed.
            shiftPressed = ((win32api.GetAsyncKeyState(win32con.VK_SHIFT) & 0x8000) >> 15) | (vkey_code in [win32con.VK_LSHIFT, win32con.VK_RSHIFT])
            keyAscii, keyName= GetKeyAsciiAndName(vkey_code, shiftPressed)
            
            eventName = eventIdToName[wParam]
            
            # Creating a keyboard event object.
            keyboardEvent = KeyboardEvent(event_id=wParam, event_name=eventName, vkey_code=vkey_code,
                                          scancode=scancode, key_ascii=keyAscii, key_name=keyName,
                                          flags=flags, injected=injected, extended=extended,
                                          shift=shiftPressed, alt=altPressed, transition=transition)
            
            # Key down/press event.
            if wParam in [win32con.WM_KEYDOWN, win32con.WM_SYSKEYDOWN]:
                # Propagate the event to the registered keyDown listeners.
                for listener in self.keyDownListeners:
                    PThread(target=listener, args=[keyboardEvent]).start()
                
                # One of the keyDown listeners must put a signal in the PThread.msgQueue to specify whether to return or suppress the pressed key.
                if self.keyDownListeners:
                    # Return True to suppress the key, False to propagate it.
                    # Note that, if you returned a boolean value instead of calling CallNextHookEx, you are effectively preventing any further processing of the keystroke by other hooks.
                    return PThread.msgQueue.get()
            
            # Key up event.
            else:
                # Propagate the event to the registered keyUp listeners.
                for listener in self.keyUpListeners:
                    PThread(target=listener, args=[keyboardEvent]).start()
                
                return False

# ctypedef bint (*EventListenerPtr)(KeyboardEvent)

# cdef bint test(EventListenerPtr listener):
    # keyboardEvent = KeyboardEvent(event_id=256, event_name="key down", key_name="Escape",
                                    # vkey_code=27, scancode=1,
                                    # vkey_ascii=27, flags=0,
                                    # injected=False, extended=False,
                                    # shift=True, alt=False, transition=False)
    
    # return(listener(keyboardEvent))

# cdef bint MyInner(KeyboardEvent keyboardEvent):
    # print("MyInner called.")
    # print(keyboardEvent)
    # return True

# cdef EventListenerPtr myPtr = &MyInner

# print(test(myPtr))

# ###############################################################
# # Define a function pointer type
# ctypedef bint (*FuncPtr)(int, int)

# # Define a function that takes a function pointer as an argument
# cdef bint CallFunction(FuncPtr func, int value1, int value2):
    # return func(value1, value2)

# # Define a function that will be passed as a function pointer
# cdef bint MyFunction(int value1, int value2):
    # print("MyFunction called with value:", value1, value2)
    # return True

# # Create a function pointer variable and assign MyFunction to it
# cdef FuncPtr myFuncPtr = &MyFunction

# # Call the function and pass the function pointer
# print(CallFunction(myFuncPtr, 42, 77))
