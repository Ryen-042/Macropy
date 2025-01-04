# https://stackoverflow.com/questions/5353883/python3-ctype-createwindowex-simple-example

import ctypes
from ctypes import wintypes
import win32con

try:
    from cythonExtensions.guiHelper.guiBase import *
except ImportError:
    from guiBase import *

# Make the application DPI aware to prevent scaling issues
ctypes.windll.shcore.SetProcessDpiAwareness(2)


class SimpleWindow:
    def __init__(self, title: str, label_width=150, input_field_width=250, itemsHeight=20, x_sep=10, y_sep=10, x_offset=10, y_offset=10):
        self.classname = "SimpleWindowClass"
        self.title = title
        self.computed_width = 500
        self.computed_height = 300
        self.label_width = label_width
        self.input_field_width = input_field_width
        self.itemsHeight = itemsHeight
        self.x_sep = x_sep
        self.y_sep = y_sep
        self.x_offset = x_offset
        self.y_offset = y_offset
        self.hwnd: int = None
        self.hButtonOK = None
        self.hEditFields = []
        self.userInputs: list[str] = []
        
        # Key capture fields
        self.hKeyDisplay = None
        self.hCaptureButton = None
        self.isCapturing = False
        self.capturedKeyVK = None
        self.capturedKeyScanCode = None
        self.capturedKeyName = None
        self.showKeyCaptureField = False
        self.CAPTURE_BUTTON_ID = 2  # Command id for capture button

    def WndProc(self, hwnd, message, wParam, lParam):
        if message == win32con.WM_DESTROY:
            user32.PostQuitMessage(0)
            user32.UnregisterClassW(self.classname, 0)
            return 0
        
        elif message == win32con.WM_COMMAND:
            control_id = LOWORD(wParam)  # Get control id from wParam
            
            if control_id == 1:  # OK button # if lParam == self.hButtonOK:
                for hEdit in self.hEditFields:
                    buffer = ctypes.create_unicode_buffer(256)  # Adjust the buffer size as needed
                    user32.GetWindowTextW(hEdit, buffer, len(buffer))
                    self.userInputs.append(buffer.value)
                
                self.destroyWindow()
                return 0
            
            elif control_id == self.CAPTURE_BUTTON_ID:  # Capture button
                self.isCapturing = True
                user32.SetWindowTextW(self.hCaptureButton, "Press any key...")
                
                # Set focus to main window to capture keystrokes
                user32.SetFocus(self.hwnd)
                
                return 0
        
        elif message == win32con.WM_KEYDOWN:
            # if self.isCapturing:
            # Convert virtual key code to key name
            buffer = ctypes.create_unicode_buffer(32)
            scanCode = user32.MapVirtualKeyW(wParam, 0) << 16

            # Add extended key flag if needed
            if (lParam & (1 << 24)):
                scanCode |= (1 << 24)

            if wParam == win32con.VK_TAB:
                user32.SetFocus(user32.GetNextDlgTabItem(hwnd, 0, False))
                return 0

            result = user32.GetKeyNameTextW(scanCode, buffer, 32)
            if result > 0:
                self.capturedKeyVK = wParam
                self.capturedKeyScanCode = scanCode >> 16
                self.capturedKeyName = buffer.value
                user32.SetWindowTextW(self.hKeyDisplay, self.capturedKeyName)
                user32.SetWindowTextW(self.hCaptureButton, "Press to Capture Key")
                self.isCapturing = False

            elif wParam == win32con.VK_RETURN:
                button_hwnd = user32.GetDlgItem(self.hwnd, 1)
                user32.SendMessageW(button_hwnd, win32con.BM_CLICK, 0, 0)
                return 0

            elif wParam == win32con.VK_ESCAPE:
                self.destroyWindow()
                return 0

            return 0
        
        elif message == win32con.WM_SYSKEYDOWN:
            scanCode = user32.MapVirtualKeyW(wParam, 0) << 16
            if (lParam & (1 << 24)):
                scanCode |= (1 << 24)
            
            # Special handling for Alt key
            if wParam == win32con.VK_MENU:
                self.capturedKeyVK = wParam
                self.capturedKeyScanCode = scanCode >> 16
                self.capturedKeyName = "Alt"
                user32.SetWindowTextW(self.hKeyDisplay, self.capturedKeyName)
                user32.SetWindowTextW(self.hCaptureButton, "Press to Capture Key")
                self.isCapturing = False
                return 1  # Prevent system processing

        # Handle the window's resizing restrictions.
        elif message == win32con.WM_GETMINMAXINFO:
            minmaxinfo = ctypes.cast(lParam, ctypes.POINTER(MINMAXINFO)).contents
            minmaxinfo.ptMinTrackSize.x = self.computed_width  # Minimum width
            minmaxinfo.ptMinTrackSize.y = self.computed_height # Minimum height
            minmaxinfo.ptMaxTrackSize.y = self.computed_height # Maximum height - Comment this line to allow resizing in the vertical direction
            
            return 0
        
        # Adjust the width of the edit boxes and button when the window is resized.
        elif message == win32con.WM_SIZE:
            # Get new width and height
            width = LOWORD(lParam)
            # self.computed_height = HIWORD(lParam)
            
            # Update captured key display field width
            if self.showKeyCaptureField:
                user32.SetWindowPos(
                    self.hCaptureButton,
                    None,
                    0, 0,
                    width - (self.label_width + self.x_sep + 2 * self.x_offset),
                    self.itemsHeight,
                    win32con.SWP_NOMOVE | win32con.SWP_NOZORDER
                )
            
            # Update the width of the other fields
            for hEdit in self.hEditFields:
                user32.SetWindowPos(
                    hEdit,
                    None,
                    0, 0,
                    width - (self.label_width + self.x_sep + 2 * self.x_offset),
                    self.itemsHeight,
                    win32con.SWP_NOMOVE | win32con.SWP_NOZORDER
                )
            
            # Update the width of the OK button
            user32.SetWindowPos(
                self.hButtonOK,
                None,
                0, 0,
                width - 2 * self.x_offset,
                self.itemsHeight + 5,
                win32con.SWP_NOMOVE | win32con.SWP_NOZORDER
            )
            return 0
        
        # Call the default window procedure for messages not handled above
        return user32.DefWindowProcW(hwnd, message, wParam, lParam)

    def destroyWindow(self) -> None:
        """Destroy the window and unregister the window class"""
        
        user32.DestroyWindow(self.hwnd)
        user32.UnregisterClassW(self.classname, 0)

    def createDynamicInputWindow(self, input_labels: list[str], placeholders: list[str]=None, enable_key_capture=False):
        self.showKeyCaptureField = enable_key_capture
        
        if placeholders and len(placeholders) != len(input_labels):
            print("Warning! The number of placeholders should be equal to the number of input labels. Continuing without placeholders...")
            placeholders = None

        # Window class definition
        wndclass = WNDCLASSW()
        wndclass.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
        wndclass.lpfnWndProc = WNDPROC(self.WndProc)
        wndclass.cbClsExtra = wndclass.cbWndExtra = 0
        wndclass.hInstance = kernel32.GetModuleHandleW(None)
        wndclass.hIcon = user32.LoadIconW(None, IDI_APPLICATION)
        wndclass.hCursor = user32.LoadCursorW(None, IDC_ARROW)
        wndclass.hbrBackground = gdi32.GetStockObject(1)  # Darker background color
        wndclass.lpszMenuName = None
        wndclass.lpszClassName = self.classname

        # Register Window Class
        try:
            user32.RegisterClassW(ctypes.byref(wndclass))
        except OSError:
            print("Failed to register window class as it already exists. Continuing...")

        # Adjust window dimensions
        self.computed_width = self.label_width + self.input_field_width + self.x_sep + 2 * self.x_offset
        capturedKeyHeight = self.itemsHeight + self.y_sep if enable_key_capture else 0
        self.computed_height = 60 + capturedKeyHeight + (self.itemsHeight + self.y_sep) * len(input_labels) + self.itemsHeight + 2 * self.y_offset

        # Create main window
        self.hwnd = user32.CreateWindowExW(
            win32con.WS_EX_TOPMOST,
            wndclass.lpszClassName,
            self.title,
            win32con.WS_OVERLAPPEDWINDOW,
            win32con.CW_USEDEFAULT,
            win32con.CW_USEDEFAULT,
            self.computed_width,
            self.computed_height,
            None,
            None,
            wndclass.hInstance,
            None
        )

        if not placeholders:
            placeholders = [""] * len(input_labels)

        y_pos = self.y_offset

        # Create key capture field if enabled
        if enable_key_capture:
            # Captured key display field
            self.hKeyDisplay = user32.CreateWindowExW(
                0,
                "STATIC",
                "",
                win32con.WS_CHILD | win32con.WS_VISIBLE | win32con.SS_CENTER | win32con.SS_CENTERIMAGE,
                self.x_offset,
                y_pos,
                self.label_width,
                self.itemsHeight,
                self.hwnd,
                None,
                kernel32.GetModuleHandleW(None),
                None
            )

            # Capture key button
            self.hCaptureButton = user32.CreateWindowExW(
                0,
                "Button",
                "Press to Capture Key",
                win32con.WS_CHILD | win32con.WS_VISIBLE | win32con.BS_PUSHBUTTON | win32con.WS_TABSTOP,
                self.x_offset + self.label_width + self.x_sep,
                y_pos,
                self.input_field_width,
                self.itemsHeight,
                self.hwnd,
                ctypes.c_void_p(self.CAPTURE_BUTTON_ID),  # Command ID
                kernel32.GetModuleHandleW(None),
                None
            )
            
            comctl32.SetWindowSubclass(self.hCaptureButton, SubclsProc, self.CAPTURE_BUTTON_ID, self.hwnd)
            y_pos += self.itemsHeight + self.y_sep

        # Create Label and Edit controls dynamically
        for label, placeholder in zip(input_labels, placeholders):
            user32.CreateWindowExW(
                0,
                "STATIC",
                label,
                win32con.WS_CHILD | win32con.WS_VISIBLE | win32con.SS_CENTER | win32con.SS_CENTERIMAGE,
                self.x_offset,
                y_pos,
                self.label_width,
                self.itemsHeight,
                self.hwnd,
                None,
                kernel32.GetModuleHandleW(None),
                None
            )
            
            hEdit = user32.CreateWindowExW(
                0,
                "Edit",
                str(placeholder),
                win32con.WS_CHILD | win32con.WS_VISIBLE | win32con.WS_BORDER | win32con.ES_AUTOHSCROLL | win32con.ES_LEFT | win32con.WS_TABSTOP,
                self.x_offset + self.label_width + self.x_sep,
                y_pos,
                self.input_field_width,
                self.itemsHeight,
                self.hwnd,
                None,
                kernel32.GetModuleHandleW(None),
                None
            )
            self.hEditFields.append(hEdit)
            
            # Subclass the edit control to handle tab key press
            comctl32.SetWindowSubclass(hEdit, SubclsProc, 1, self.hwnd)
            
            y_pos += self.itemsHeight + self.y_sep
        
        # Create OK button
        self.hButtonOK = user32.CreateWindowExW(
            0,
            "Button",
            "OK",
            win32con.WS_CHILD | win32con.WS_VISIBLE | win32con.BS_PUSHBUTTON | win32con.WS_TABSTOP,
            self.x_offset,
            y_pos,
            self.label_width + self.input_field_width + self.x_sep,
            self.itemsHeight,
            self.hwnd,
            ctypes.c_void_p(1),  # Assigning the button an command id
            kernel32.GetModuleHandleW(None),
            None
        )
        
        comctl32.SetWindowSubclass(self.hButtonOK, SubclsProc, 1, self.hwnd)

        # Show window
        user32.ShowWindow(self.hwnd, win32con.SW_SHOWNORMAL)
        user32.UpdateWindow(self.hwnd)

        # Message loop
        msg = wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))


if __name__ == '__main__':
    window = SimpleWindow("Dynamic Input Window", 150, 250, 50, 20, 10, 50, 50)
    window.createDynamicInputWindow(
        ["Character", "Delay", "Another Input"],
        placeholders=["A", "0", "Another Placeholder"],
        enable_key_capture=True  # Enable the key capture field
    )
    print(window.userInputs)
    print("Captured VK:", window.capturedKeyVK)
    print("Captured ScanCode:", window.capturedKeyScanCode)
    print("Captured Key Name:", window.capturedKeyName)
