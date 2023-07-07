cdef class BaseEvent:
    cdef public int EventId, Flags
    cdef public str EventName

cdef class KeyboardEvent(BaseEvent):
    cdef public int KeyID, Scancode, Ascii
    cdef public str Key
    cdef public bint Injected, Extended, Shift, Alt, Transition

cdef class MouseEvent(BaseEvent):
    cdef public int X, Y, MouseData, Delta, PressedButton #, "Time"
    cdef public bint IsMouseAbsolute, IsMouseInWindow, IsWheelHorizontal
