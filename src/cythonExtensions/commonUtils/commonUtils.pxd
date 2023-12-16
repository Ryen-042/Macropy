cdef class BaseEvent:
    cdef int EventId, Flags
    cdef EventName

cdef class KeyboardEvent(BaseEvent):
    cdef int KeyID, Scancode, Ascii
    cdef bint Injected, Extended, Shift, Alt, Transition
    cdef Key

cdef class MouseEvent(BaseEvent):
    cdef int X, Y, MouseData, Delta, PressedButton #, "Time"
    cdef bint IsMouseAbsolute, IsMouseInWindow, IsWheelHorizontal
