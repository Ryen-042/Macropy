"""This module provides class definitions used in other modules."""

from win32com.client import Dispatch
import win32api, win32con, win32clipboard, pythoncom, ctypes
import keyboard, threading, queue
from enum import IntEnum
from traceback import format_tb, format_exc
from pynput.keyboard import Controller as kbController
from pynput.mouse import Controller as mController
from time import time, sleep
from collections import deque


class KB_Con(IntEnum):
    """
    Description:
        An enum class for storing keyboard `keyID` and `scancode` constants not present in the `win32con` module.
    ---
    Naming Scheme:
        `AS`: Ascii value.
        `VK`: Virtual keyID.
        `SC`: Scancode
    ---
    Notes:
        - No Ascii constant stored for keys that have Ascii value equal to their VK value.
        - Keys may send different `Ascii` values depending on the pressed modifiers, but they send the same `keyID` and `scancode`.
            - If you need a code that is independent of the pressed modifiers, use `keyID`.
            - If you need a code that may give different values, use `Ascii` (ex, `=` would give `61`, `+` would give `41`.)
        - `KeyID` and `scancode` constants are stored only one time for each physical key.
            - These constants are named with respect to the sent key when the corresponding physical key is pressed with no modifiers.
            - Letter keys are named with the uppercase letter, not the lowercase letters unlike what mentioned above.
    """
    
    # Letter keys - Uppercase letters
    VK_A = 65;  SC_A = 30
    VK_B = 66;  SC_B = 48
    VK_C = 67;  SC_C = 46
    VK_D = 68;  SC_D = 32
    VK_E = 69;  SC_E = 18
    VK_F = 70;  SC_F = 33
    VK_G = 71;  SC_G = 34
    VK_H = 72;  SC_H = 35
    VK_I = 73;  SC_I = 23
    VK_J = 74;  SC_J = 36
    VK_K = 75;  SC_K = 37
    VK_L = 76;  SC_L = 38
    VK_M = 77;  SC_M = 50
    VK_N = 78;  SC_N = 49
    VK_O = 79;  SC_O = 24
    VK_P = 80;  SC_P = 25
    VK_Q = 81;  SC_Q = 16
    VK_R = 82;  SC_R = 19
    VK_S = 83;  SC_S = 31
    VK_T = 84;  SC_T = 20
    VK_U = 85;  SC_U = 22
    VK_V = 86;  SC_V = 47
    VK_W = 87;  SC_W = 17
    VK_X = 88;  SC_X = 45
    VK_Y = 89;  SC_Y = 21
    VK_Z = 90;  SC_Z = 44
    
    # Lowercase letters
    AS_a = 97
    AS_b = 98
    AS_c = 99
    AS_d = 100
    AS_e = 101
    AS_f = 102
    AS_g = 103
    AS_h = 104
    AS_i = 105
    AS_j = 106
    AS_k = 107
    AS_l = 108
    AS_m = 109
    AS_n = 110
    AS_o = 111
    AS_p = 112
    AS_q = 113
    AS_r = 114
    AS_s = 115
    AS_t = 116
    AS_u = 117
    AS_v = 118
    AS_w = 119
    AS_x = 120
    AS_y = 121
    AS_z = 122
    
    
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
    AS_DOUBLE_QUOTES        = 34;   VK_DOUBLE_QUOTES = 222;       SC_DOUBLE_QUOTES = 40
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
    
    
    # Miscellaneous keys
    SC_RETURN      = 28 # Enter
    SC_BACK        = 14 # Backspace
    SC_MENU        = 56 # 'LMenu' and 'RMenu'
    SC_HOME        = 71
    SC_UP          = 72
    SC_RIGHT       = 77
    SC_DOWN        = 80
    SC_LEFT        = 75
    SC_VOLUME_UP   = 48
    SC_VOLUME_DOWN = 46


class GenericDataClass:
    """A generic data class that can be customized at initialization."""
    
    def __init__(self, name=None, **kwargs):
        self.class_name = name
        
        for key, value in kwargs.items():
            setattr(self, key, value)
    
    def __iter__(self):
        return iter([v for k, v in vars(self).items() if k != 'class_name'])
    
    def __repr__(self):
        attributes = ', '.join(f'{key}={value}' for key, value in vars(self).items() if key != "name")
        return f'{self.class_name or self.__class__.__name__}({attributes})'


# Source: https://stackoverflow.com/questions/6760685/creating-a-singleton-in-python
class Singleton(type):
    """A metaclass that, when used by other classes, returns the same instance of this class whenever it is instantiated."""
    
    _instances = {}
    
    def __call__(cls, *args, **kwargs):
        if cls not in cls._instances:
            cls._instances[cls] = super(Singleton, cls).__call__(*args, **kwargs)
        # To run __init__ every time the class is called, uncomment the else statement.
        #else:
         #   cls._instances[cls].__init__(*args, **kwargs)
        return cls._instances[cls]


def static_class(cls):
    """Class decorator that turns a class into a static class."""
    
    class StaticClass(type):
        def __call__(cls, *args, **kwargs):
            raise TypeError("This is a static class, it cannot be instantiated.")
    return StaticClass(cls.__name__, (cls,), {})

@static_class
class DebuggingHouse:
    counter           = 0     # For debugging purposes only.
    silent            = False # For suppressing terminal output.
    prev_event        = 0     # {1: "Hotkey event", 0: "No event"}
    suppress_all_keys = 0     # For suppressing all keyboard keys.
    terminate_script  = False
    
    @staticmethod
    def ShowErrorMessage(msg: str) -> int:
        """
        Description:
            Display an error message window with the message.
        ---
        Parameters:
            `msg -> str`:
                The error message to display.
        ---
        Returns:
            `int`: `1` is returned when the error message is closed.
        """
        
        # https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-messageboxa
        return ctypes.windll.user32.MessageBoxW(0, msg, "Warning", win32con.MB_OK | win32con.MB_ICONERROR | win32con.MB_TOPMOST)
    
    @staticmethod
    def LogUncaughtExceptions(exc_type, exc_value, exc_traceback):
        """Logs uncaught exceptions to the console and displays an error message with the exception traceback."""
        
        message = f"{exc_type.__name__}: {exc_value}\n  "
        message += ''.join(format_tb(exc_traceback)).strip()
        print(f'\nUncaught exception: """{message}"""', end="\n\n")
        
        return DebuggingHouse.ShowErrorMessage(message)
    
    @staticmethod
    def TogglePrevEvent():
        DebuggingHouse.prev_event ^= 1
        return DebuggingHouse.prev_event

@static_class
class WindowHouse:
    """A dataclass for storing class names and a corresponding single window handle for each of them."""
    
    classNames: dict[str, int] = {"MediaPlayerClassicW": None}
    
    closedExplorers = deque(maxlen=10)
    
    @staticmethod
    def GetClassNameHandle(className) -> int | None:
        return WindowHouse.classNames.get(className, None)
    
    @staticmethod
    def SetClassNameHandle(className, value):
        WindowHouse.classNames[className] = value

@static_class
class ControllerHouse:
    """A dataclass used for storing information related to keyboard keys."""
    
    modifiers = GenericDataClass(name="Modifiers", CTRL=False, SHIFT=False, ALT=False, WIN=False, FN=False, BACKTICK=False)
    
    # GetAsyncKeyState vs GetKeyState: https://stackoverflow.com/questions/17770753/getkeystate-vs-getasynckeystate-vs-getch
    locks = GenericDataClass(name="Locks", CAPITAL=win32api.GetKeyState(win32con.VK_CAPITAL),
                             SCROLL=win32api.GetKeyState(win32con.VK_SCROLL),
                             NUMLOCK=win32api.GetKeyState(win32con.VK_NUMLOCK))  # [CapsLk, ScrLk, NumLk]
    
    # Holds the pressed character keys.
    pressed_chars = ""
    
    # Abbreviations for expansion:
    abbreviations = {":py"    : "python",
                     ":name" : "name place holder",
                     ":gmail" : "testmail123@host.com"}
    
    locations = {"!cmd":    r"C:\Windows\System32\cmd.exe",
                 "!paint":  r"C:\Windows\System32\mspaint.exe",
                 "!prog":   r"C:\Program Files"}
    
    
    
    # https://nitratine.net/blog/post/simulate-keypresses-in-python/
    pynput_kb = kbController()
    pynput_mouse = mController()
    keyboard_send = keyboard.send
    keyboard_write = keyboard.write
    
    operations = {
        "|": lambda x, y: x | y,
        "&": lambda x, y: x & y
    }
    
    @staticmethod
    def UpdateLockKeys(event) -> None:
        """Used for updating the states of Lock keys."""
        
        ControllerHouse.locks.CAPITAL ^= event.KeyID == win32con.VK_CAPITAL  # 'Capital' => CapsLock
        ControllerHouse.locks.SCROLL  ^= event.KeyID == win32con.VK_SCROLL   # 'Scroll'  => ScrollLock
        ControllerHouse.locks.NUMLOCK ^= event.KeyID == win32con.VK_NUMLOCK  # 'NumLock'
    
    @staticmethod
    def UpdateModifierKeys(event, operator="|") -> None:
        """Used for updating the states of modifier keys. Pass the operator `|` in `KeyPress` events, and `&` in `KeyRelease` events."""
        
        # Updating the state of a modifier in `KeyPress` event is different than in `KeyRelease` event.
        # The difference is just the opperator ('|' for `KeyPress`, '&' for `KeyRelease`), and the check is inverted in case of `KeyRelease`.
        # We can divide this function into two like this:
            # KeyPress version:   ctrlhouse.modifiers.CTRL |= event.KeyID     in [win32con.VK_LCONTROL, win32con.VK_RCONTROL] # ['Lcontrol', 'Rcontrol']
            # KeyRelease version: ctrlhouse.modifiers.CTRL &= event.KeyID not in [win32con.VK_LCONTROL, win32con.VK_RCONTROL] # ['Lcontrol', 'Rcontrol']
        
        ControllerHouse.modifiers.CTRL     = ControllerHouse.operations[operator](ControllerHouse.modifiers.CTRL,     (operator=="&") ^ (event.KeyID in [win32con.VK_LCONTROL, win32con.VK_RCONTROL])) # ['Lcontrol', 'Rcontrol']
        ControllerHouse.modifiers.SHIFT    = ControllerHouse.operations[operator](ControllerHouse.modifiers.SHIFT,    (operator=="&") ^ (event.KeyID in [win32con.VK_LSHIFT,   win32con.VK_RSHIFT]))   # ['Lshift', 'Rshift']
        ControllerHouse.modifiers.ALT      = ControllerHouse.operations[operator](ControllerHouse.modifiers.ALT,      (operator=="&") ^ (event.KeyID in [win32con.VK_LMENU,    win32con.VK_RMENU]))    # ['Lmenu', 'Rmenu']
        ControllerHouse.modifiers.WIN      = ControllerHouse.operations[operator](ControllerHouse.modifiers.WIN,      (operator=="&") ^ (event.KeyID in [win32con.VK_LWIN,     win32con.VK_RWIN]))     # ['Lwin', 'Rwin']
        ControllerHouse.modifiers.FN       = ControllerHouse.operations[operator](ControllerHouse.modifiers.FN,       (operator=="&") ^ (event.KeyID == 255)) # FN Key
        ControllerHouse.modifiers.BACKTICK = ControllerHouse.operations[operator](ControllerHouse.modifiers.BACKTICK, (operator=="&") ^ (event.KeyID == KB_Con.VK_BACKTICK)) # 'Oem_3'


@static_class
class ShellAutomationObjectWrapper:
    """Thread-safe wrapper class for accessing an Automation object in a multithreaded environment."""
    #> The wrapper object acquires a lock before accessing the Automation object, ensuring that only one thread can access it at a time.
     # This can help preventting problems that can occur when multiple threads try to access the Automation object concurrently.
    #> If two or more threads tried to access an automation object wrapper, only one would be able to acquire the lock to the automation object at a time. 
     # The other threads would be blocked until the lock is released by the thread that currently holds it.
    #> When a threading.Lock object is used in a `with` statement, the lock is released when the `with` block ends.
     # The lock is automatically acquired at the beginning of the block and released when the block ends.
    
    explorer = Dispatch("Shell.Application")
    _lock = threading.Lock()
    
    # @staticmethod
    # def set_obj(obj):
    #     """For resetting and changing the automation object type."""
    #     ShellAutomationObjectWrapper.explorer = obj
    
    @staticmethod
    def __getattr__(name):
        # Acquire lock before accessing the Automation object
        with ShellAutomationObjectWrapper._lock:
            return getattr(ShellAutomationObjectWrapper.explorer, name)


# Source: https://stackoverflow.com/questions/6552097/threading-how-to-get-parent-id-name
class PThread(threading.Thread):
    """An extension of threading.Thread. Stores some relevant information about the threads."""
    
    # Used in the Throttle decorator to make it thread-safe.
    computer_state_lock = threading.Lock()
    
    outputQueue = queue.Queue()
    
    def __init__(self, *args, **kwargs):
        threading.Thread.__init__(self, *args, **kwargs)
        self.parent = threading.current_thread()
        self.coInitializeCalled = False
    
    # Propagating exceptions from a thread.
    def run(self):
        try:
            self.ret = self._target(*self._args, **self._kwargs)
        except BaseException as e:
            error_msg = format_exc()
            class_name = str(type(e)).split("'")[1]
            print(f'Warning! An error occurred in thread {self.name}.\nClass: {class_name}\nValue: {str(e)}\n\n"""\n{error_msg}"""')
            DebuggingHouse.ShowErrorMessage(error_msg)
            # DebuggingHouse.LogUncaughtExceptions(*sys.exc_info()) 
            # logging.error(f"Warning! An error occurred in thread {self.name}. Traceback:{str(e)}", exc_info=True)
    
    
    @staticmethod
    def InMainThread():
        """Returns where the current thread is the main thread for the current process."""
        
        return threading.get_ident() == threading.main_thread().ident
    
    @staticmethod
    def GetParentThread() -> int:
        """
        Description:
            Returns the parent thread id of the current thread.
        ---
        Returns:
            - The parent thread id: if the current thread is not the main thread.
            - `0`: if it was the main thread.
            - `-1`: if the parent thread is unknown (i.e., the current thread was not created using this class).
        """
        
        if threading.current_thread().name != "MainThread":
            if hasattr(threading.current_thread(), 'parent'):
                return threading.current_thread().parent.ident
            else:
                return -1 # Unknown. Current thread does not have `parent` property.
        else:
            return 0 # "MainThread"
    
    @staticmethod
    def CoInitialize():
        """Initializes the COM library for the current thread if it was not previously initialized."""
        
        if not PThread.InMainThread() and not threading.current_thread().coInitializeCalled:
            print(f"CoInitialize called from: {threading.current_thread().name}")
            threading.current_thread().coInitializeCalled = True
            pythoncom.CoInitialize()
            return True
        return False
    
    
    @staticmethod
    def CoUninitialize(initializer_called: bool):
        """Un-initializes the COM library for the current thread if `initializer_called` is True."""
        
        if initializer_called:
            print(f"CoUninitialize called from: {threading.current_thread().name}")
            pythoncom.CoUninitialize()
            threading.current_thread().coInitializeCalled = False
            return True
        return False
    
    
    # Source: https://github.com/salesforce/decorator-operations/blob/master/decoratorOperations/throttle_functions/throttle.py
    # For a comparison between Throttling and Debouncing: https://stackoverflow.com/questions/25991367/difference-between-throttling-and-debouncing-a-function
    @staticmethod
    def Throttle(wait_time):
        """
        Description:
            Decorator function that ensures that a wrapped function is called only one time each time slot.
        """
        
        def decorator(function):
            def throttled(*args, **kwargs):
                def call_function():
                    return function(*args, **kwargs)
                
                if PThread.computer_state_lock.acquire(blocking = False):
                    if time() - throttled._last_time_called >= wait_time:
                        call_function()
                        throttled._last_time_called = time()
                    PThread.computer_state_lock.release()
                
                else: # Lock is already acquired by another thread, so return without calling the wrapped function
                    return
            
            throttled._last_time_called = 0
            
            return throttled
        
        return decorator
    
    # Source: https://github.com/salesforce/decorator-operations/blob/master/decoratorOperations/debounce_functions/debounce.py
    @staticmethod
    def Debounce(wait_time):
        """
        Decorator that will debounce a function so that it is called after `wait_time` seconds.
        If the decorated function is called multiple times within a time slot, it will debounce
        (wait for a new time slot from the last call).
        
        If no calls arrive after wait_time seconds from the last call, then execute the last call.
        """
        
        def decorator(function):
            def debounced(*args, **kwargs):
                def call_function():
                    debounced._timer = None
                    return function(*args, **kwargs)
                
                if debounced._timer is not None:
                    debounced._timer.cancel()
                
                debounced._timer = threading.Timer(wait_time, call_function)
                debounced._timer.start()
            
            debounced._timer = None
            return debounced
        
        return decorator

def ReadFromClipboard(CF=win32clipboard.CF_TEXT): # CF: Clipboard format.
    """Reads the top of the clipboard if it was the same type as the specified."""
    
    win32clipboard.OpenClipboard()
    if win32clipboard.IsClipboardFormatAvailable(CF):
        clipboard_data = win32clipboard.GetClipboardData()
        # win32clipboard.EmptyClipboard()
        win32clipboard.CloseClipboard()
    else:
        return
    
    return clipboard_data


def SendToClipboard(data, CF=win32clipboard.CF_UNICODETEXT):
    """Copy the given data to the clipboard."""
    
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(CF, data)
    win32clipboard.CloseClipboard()


if __name__ == "__main__":
    def test_debounce():
        print("started")
        
        @PThread.Debounce(3)
        def sleeping(i):
            print(f"Sleeping {i}...")
        
        sleeping(1)
        print(1)
        
        sleeping(2)
        print(2)
        
        sleeping(3)
        print(3)
        
        sleeping(4)  # Only this executes! (The last call before the time slot expires)
        print(4)
        
        sleep(4)
        
        sleeping(5)
        print(5)
        
        sleeping(6)
        print(6)
        
        sleeping(7)
        print(7)
        
        while True:
            sleeping(0)
    
    def test_throttle():
        print("started")
        
        @PThread.Throttle(3)
        def sleeping(i):
            print(f"Sleeping {i}...")
        
        print(threading.Thread(sleeping(1)).start()) # Only this executes!
        print(threading.Thread(sleeping(2)).start())
        print(threading.Thread(sleeping(3)).start())
        
        # sleep(4)
        # while True:
        #     threading.Thread(sleeping(4)).start()
    
    # test_debounce()
    # test_throttle()
