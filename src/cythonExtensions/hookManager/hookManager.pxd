# https://learn.microsoft.com/en-us/windows/win32/winmsg/about-hooks
cdef enum HookTypes:
    # Constants that represent different types of Windows hooks that can be set using the SetWindowsHookEx function.
    
    WH_MSGFILTER       = -1   # A hook type that monitors messages sent to a window as a result of an input event in a dialog box, message box, menu, or scroll bar. Also represents the minmum value for a hook type.
    
    # Warning: Journaling Hooks APIs are unsupported starting in Windows 11 and will be removed in a future release. Use  the SendInput TextInput API instead.
    WH_JOURNALRECORD   = 0    # A hook type used to record input messages posted to the system message queue. This hook is useful for recording macros.
    WH_JOURNALPLAYBACK = 1    # A hook type used to replay/post the input messages recorded by the WH_JOURNALRECORD hook type.
    
    WH_KEYBOARD        = 2    # A hook type that monitors keystrokes. It can be used to implement a keylogger or to perform other keyboard-related tasks.
    WH_GETMESSAGE      = 3    # A hook type that monitors messages in the message queue of a thread. It can be used to modify or filter messages before they are dispatched.
    WH_CALLWNDPROC     = 4    # A hook type that monitors messages sent to a window before the system sends them to the destination window procedure.
    WH_CBT             = 5    # A hook type that monitors a variety of system events, including window creation and destruction, window activation, and window focus changes.
    WH_SYSMSGFILTER    = 6    # A hook type similar to WH_MSGFILTER but monitors messages sent to all windows, not just a specific window or dialog box procedure.
    WH_MOUSE           = 7    # A hook type that monitors mouse events.
    
    # Warning: Not Implemented hook. May have been intended for system activities such as watching for disk or other hardware activity to occur.
    WH_HARDWARE        = 8    # A hook type that monitors low-level hardware events other than a mouse or keyboard events.
    
    WH_DEBUG           = 9    # A hook type useful for debugging other hook procedures.
    WH_SHELL           = 10   # A hook type that monitors shell-related events, such as when the shell is about to activate a new window.
    WH_FOREGROUNDIDLE  = 11   # A hook type that monitors when the system enters an idle state and the foreground application is not processing input.
    WH_CALLWNDPROCRET  = 12   # A hook type that similar to WH_CALLWNDPROC but monitors messages after they have been processed by the destination window procedure.
    WH_KEYBOARD_LL     = 13   # A hook type that similar to WH_KEYBOARD but monitors low-level keyboard input events.
    WH_MOUSE_LL        = 14   # A hook type that similar to WH_MOUSE but monitors low-level mouse input events.
    WH_MAX             = 15   # The maximum value for a hook type. It is not a valid hook type itself.


cdef enum KbMsgIds:
    # Contains the Windows Message Ids for keyboard events.
    WM_KEYDOWN     = 0x0100   # A keyboard key was pressed.
    WM_KEYUP       = 0x0101   # A keyboard key was released.
    WM_CHAR        = 0x0102   # A keyboard key was pressed down and released, and it represents a printable character.
    WM_DEADCHAR    = 0x0103   # A keyboard key was pressed down and released, and it represents a dead character.
    WM_SYSKEYDOWN  = 0x0104   # A keyboard key was pressed while the ALT key was held down.
    WM_SYSKEYUP    = 0x0105   # A keyboard key was released after the ALT key was held down.
    WM_SYSCHAR     = 0x0106   # A keyboard key was pressed down and released, and it represents a printable character while the ALT key was held down.
    WM_SYSDEADCHAR = 0x0107   # A keyboard key was pressed down and released, and it represents a dead character while the ALT key was held down.
    WM_KEYLAST     = 0x0108   # Defines the maximum value for the range of keyboard-related messages.


cdef tuple getKeyAsciiAndName(int vkey_code, bint shiftPressed=*)


cdef class HookManager:
    cdef int kbHookId, msHookId
    cdef kbHookPtr, msHookPtr

    cdef bint installHook(self, callBack, int hookType=*)
    
    cdef bint beginListening(self)
    
    cdef bint uninstallHook(self, int hookType=*)


cdef class KeyboardHookManager:
    cdef list keyDownListeners, keyUpListeners
    cdef int hookId
    
    cdef bint keyboardCallback(self, int nCode, int wParam, void * lParam)


# Docs: https://learn.microsoft.com/en-us/windows/win32/inputdev/about-mouse-input
cdef enum MsMsgIds:
    # Contains the Windows Message Ids for mouse events.
    
    WM_MOUSEMOVE       = 0x0200 # The mouse was moved.
    WM_LBUTTONDOWN     = 0x0201 # The left mouse button was pressed.
    WM_LBUTTONUP       = 0x0202 # The left mouse button was released.
    WM_LBUTTONDBLCLK   = 0x0203 # The left mouse button was double-clicked.
    WM_RBUTTONDOWN     = 0x0204 # The right mouse button was pressed.
    WM_RBUTTONUP       = 0x0205 # The right mouse button was released.
    WM_RBUTTONDBLCLK   = 0x0206 # The right mouse button was double-clicked.
    WM_MBUTTONDOWN     = 0x0207 # The middle mouse button was pressed.
    WM_MBUTTONUP       = 0x0208 # The middle mouse button was released.
    WM_MBUTTONDBLCLK   = 0x0209 # The middle mouse button was double-clicked.
    WM_MOUSEWHEEL      = 0x020A # The mouse wheel was rotated.
    WM_XBUTTONDOWN     = 0x020B # A mouse extended button was pressed.
    WM_XBUTTONUP       = 0x020C # A mouse extended button was released.
    WM_XBUTTONDBLCLK   = 0x020D # A mouse extended button was double-clicked.
    WM_MOUSEHWHEEL     = 0x020E # The mouse wheel was rotated horizontally.
    WM_NCXBUTTONDOWN   = 0x00AB # Non-client area extended button down event.
    WM_NCXBUTTONUP     = 0x00AC # Non-client area extended button up event.
    WM_NCXBUTTONDBLCLK = 0x00AD # Non-client area extended button double-click event.


# https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-rawmouse
# https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-mouseinput
cdef enum RawMouse:
    # Contains information about the state of the mouse.
    
    MOUSE_MOVE_RELATIVE      = 0x00 # Movement data is relative to the last mouse position.
    MOUSE_MOVE_ABSOLUTE      = 0x01 # Movement data is based on absolute position.
    MOUSE_VIRTUAL_DESKTOP    = 0x02 # Coordinates are mapped to the virtual desktop (for a multiple monitor system).
    MOUSE_ATTRIBUTES_CHANGED = 0x04 # Attributes changed; application needs to query the mouse attributes.
    MOUSE_MOVE_NOCOALESCE    = 0x08 # Mouse movement event was not coalesced. Mouse movement events can be coalesced by default. This value is not supported for `Windows XP`/`2000`.


cdef class MouseHookManager:
    """A class for managing mouse hooks and their event listeners."""
    
    cdef list mouseButtonDownListeners, mouseButtonUpListeners
    cdef int hookId
    
    
    cdef bint mouseCallback(self, int nCode, int wParam, void * lParam)