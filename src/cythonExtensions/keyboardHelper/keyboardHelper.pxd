cdef void simulateKeyPress(int key_id, int key_scancode=*, int times=*)

cdef void simulateKeyPressSequence(tuple keys_list, float delay=*)

cdef int getCaretPosition(text, caret=*)