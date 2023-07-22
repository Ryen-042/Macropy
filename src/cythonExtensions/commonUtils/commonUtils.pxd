cdef class BaseEvent:
    cdef int EventId, Flags
    cdef str EventName

cdef class KeyboardEvent(BaseEvent):
    cdef int KeyID, Scancode, Ascii
    cdef str Key
    cdef bint Injected, Extended, Shift, Alt, Transition

cdef class MouseEvent(BaseEvent):
    cdef int X, Y, MouseData, Delta, PressedButton #, "Time"
    cdef bint IsMouseAbsolute, IsMouseInWindow, IsWheelHorizontal

cdef enum KB_Con:
    # """
    # Description:
    #     An enum class for storing keyboard `keyID` and `scancode` constants that are not present in the `win32con` module.
    # ---
    # Naming Scheme:
    #     `AS`: Ascii value.
    #     `VK`: Virtual keyID.
    #     `SC`: Scancode
    # ---
    # Notes:
    #     - Keys may send different `Ascii` values depending on the pressed modifier(s), but they send the same `keyID` and `scancode`.
    #         - If you need a code that is independent of the pressed modifiers, use `keyID`.
    #         - If you need a code that may have different values, use `Ascii`. Eg, Ascii of `=` is `61`, `+` (`Shift` + `=`) is `41`.
    #     - The class has only one copy of `KeyID` and `scancode` constants for each physical key.
    #         - These constants are named with respect to the pressed key with no modifiers.
    #         - To keep the naming scheme consistent with capitalizing all characters, letter keys are named using the uppercase version.
    #     - Capital letters (Shift + letter key) have Ascii values equal to their VK values. As such, the class only stores the Ascii of lowercase letters.
    # """
    
    # Letter keys - Uppercase letters
    AS_a = 97,  VK_A = 65,  SC_A = 30
    AS_b = 98,  VK_B = 66,  SC_B = 48
    AS_c = 99,  VK_C = 67,  SC_C = 46
    AS_d = 100, VK_D = 68,  SC_D = 32
    AS_e = 101, VK_E = 69,  SC_E = 18
    AS_f = 102, VK_F = 70,  SC_F = 33
    AS_g = 103, VK_G = 71,  SC_G = 34
    AS_h = 104, VK_H = 72,  SC_H = 35
    AS_i = 105, VK_I = 73,  SC_I = 23
    AS_j = 106, VK_J = 74,  SC_J = 36
    AS_k = 107, VK_K = 75,  SC_K = 37
    AS_l = 108, VK_L = 76,  SC_L = 38
    AS_m = 109, VK_M = 77,  SC_M = 50
    AS_n = 110, VK_N = 78,  SC_N = 49
    AS_o = 111, VK_O = 79,  SC_O = 24
    AS_p = 112, VK_P = 80,  SC_P = 25
    AS_q = 113, VK_Q = 81,  SC_Q = 16
    AS_r = 114, VK_R = 82,  SC_R = 19
    AS_s = 115, VK_S = 83,  SC_S = 31
    AS_t = 116, VK_T = 84,  SC_T = 20
    AS_u = 117, VK_U = 85,  SC_U = 22
    AS_v = 118, VK_V = 86,  SC_V = 47
    AS_w = 119, VK_W = 87,  SC_W = 17
    AS_x = 120, VK_X = 88,  SC_X = 45
    AS_y = 121, VK_Y = 89,  SC_Y = 21
    AS_z = 122, VK_Z = 90,  SC_Z = 44
    
    # Number keys
    VK_0 = 48,  SC_0 = 11
    VK_1 = 49,  SC_1 = 2
    VK_2 = 50,  SC_2 = 3
    VK_3 = 51,  SC_3 = 4
    VK_4 = 52,  SC_4 = 5
    VK_5 = 53,  SC_5 = 6
    VK_6 = 54,  SC_6 = 7
    VK_7 = 55,  SC_7 = 8
    VK_8 = 56,  SC_8 = 9
    VK_9 = 57,  SC_9 = 10
    
    # Symbol keys
    AS_EXCLAM               = 33
    AS_DOUBLE_QUOTES        = 34,   VK_SINGLE_QUOTES = 222,       SC_SINGLE_QUOTES = 40
    AS_HASH                 = 35
    AS_DOLLAR               = 36
    AS_PERCENT              = 37
    AS_AMPERSAND            = 38
    AS_SINGLE_QUOTE         = 39
    AS_OPEN_PAREN           = 40
    AS_CLOSE_PAREN          = 41
    AS_ASTERISK             = 42
    AS_PLUS                 = 43
    AS_COMMA                = 44,   VK_COMMA  = 188,                SC_COMMA  = 51
    AS_MINUS                = 45,   VK_MINUS  = 189,                SC_MINUS  = 12
    AS_PERIOD               = 46,   VK_PERIOD = 190,                SC_PERIOD = 52
    AS_SLASH                = 47,   VK_SLASH  = 191,                SC_SLASH  = 53
    
    AS_COLON                = 58
    AS_SEMICOLON            = 59,   VK_SEMICOLON = 186,             SC_SEMICOLON = 39
    AS_LESS_THAN            = 60
    AS_EQUALS               = 61,   VK_EQUALS = 187,                SC_EQUALS = 13
    AS_GREATER_THAN         = 62
    AS_QUESTION_MARK        = 63
    AS_AT                   = 64
    
    AS_OPEN_SQUARE_BRACKET  = 91,   VK_OPEN_SQUARE_BRACKET = 219,   SC_OPEN_SQUARE_BRACKET = 26
    AS_BACKSLASH            = 92,   VK_BACKSLASH = 220,             SC_BACKSLASH = 43
    AS_CLOSE_SQUARE_BRACKET = 93,   VK_CLOSE_SQUARE_BRACKET = 221,  SC_CLOSE_SQUARE_BRACKET = 27
    AS_CARET                = 94
    AS_UNDERSCORE           = 95
    AS_BACKTICK             = 96,   VK_BACKTICK = 192,              SC_BACKTICK = 41 # '`'
    
    AS_OPEN_CURLY_BRACE     = 123
    AS_PIPE                 = 124
    AS_CLOSE_CURLY_BRACE    = 125
    AS_TILDE                = 126
    AS_OEM_102_CTRL         = 28,   VK_AS_OEM_102 = 226,            SC_OEM_102 = 86 # '\' in European keyboards (between LShift and Z).
    
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
