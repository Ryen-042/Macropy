import ctypes, win32con, win32gui, win32api, os, pythoncom, time, hashlib
from threading import Thread
from win32com.client import Dispatch
from ctypes import wintypes
from PIL import Image

try:
    from cythonExtensions.guiHelper.guiBase import *
    from cythonExtensions.guiHelper.kbHook import *
    from cythonExtensions.guiHelper.guiTester import windows_context_menu_file
except ImportError:
    from guiBase import *
    from kbHook import *
    from guiTester import windows_context_menu_file

pythoncom.CoInitialize()

# Make the application DPI aware to prevent scaling issues
ctypes.windll.shcore.SetProcessDpiAwareness(2)

SC_TITLE_DOUBLECLICK_MAXIMIZE = 61490
SC_TITLE_DOUBLECLICK_RESTORE  = 61730
# WM_USER vs WM_APP: https://stackoverflow.com/questions/30843497/wm-user-vs-wm-app
WM_SETIMAGE = win32con.WM_APP + 1  # Custom message to set the image of the window

MASKS = {
"CTRL"     : 0b10000000000000, # 1 << 13 = 8192
"LCTRL"    : 0b1000000000000,  # 1 << 12 = 4096
"RCTRL"    : 0b100000000000,   # 1 << 11 = 2048
"SHIFT"    : 0b10000000000,    # 1 << 10 = 1024
"LSHIFT"   : 0b1000000000,     # 1 << 9  = 512
"RSHIFT"   : 0b100000000,      # 1 << 8  = 256
"ALT"      : 0b10000000,       # 1 << 7  = 128
"LALT"     : 0b1000000,        # 1 << 6  = 64
"RALT"     : 0b100000,         # 1 << 5  = 32
"WIN"      : 0b10000,          # 1 << 4  = 16
"LWIN"     : 0b1000,           # 1 << 3  = 8
"RWIN"     : 0b100,            # 1 << 2  = 4
"FN"       : 0b10,             # 1 << 1  = 2
"BACKTICK" : 0b1,              # 1 << 0  = 1
}


class BuildString:
    """Builds a string from individual bytes received via PostMessage."""
    
    def __init__(self, hwnd, msg_code, encoding='utf-8'):
        self.hwnd = hwnd
        self.msg_code = msg_code
        self.encoding = encoding
        self.temp_bytes = b''

    def buildString(self, lparam):
        if lparam:
            # Accumulate received bytes
            self.temp_bytes += bytes([lparam & 0xFF])
            return None
        
        else:
            # Decode the accumulated bytes into a string
            image_path = self.temp_bytes.decode(self.encoding)
            self.temp_bytes = b''
            
            return image_path

    def sendString(self, image_path):
        encoded_path = image_path.encode(self.encoding)
        
        # Send each byte of the encoded string using PostMessage
        for byte in encoded_path:
            user32.PostMessageW(self.hwnd, self.msg_code, 0, byte)

        # Send the terminator (0) to indicate the end of the string
        user32.PostMessageW(self.hwnd, self.msg_code, 0, 0)

class ImageViewer():
    def __init__(self, title: str, image_path: str):
        self.classname = "ImageViewerClass"
        self.title = title
        self.width = 800
        self.height = 600
        self.hwnd = None
        
        self.image_path = None
        self.openImage(image_path)
        
        self.hDC = None
        self.memory_dc = None
        self.hBitmap = None
        self.hOldBitmap = None
        
        # For centering the image when the window is maximized or restored
        self.maximized_or_restored = False
        
        # Semi-related: https://devblogs.microsoft.com/oldnewthing/20110218-00/?p=11453
        self.click_through = False
        
        self.modifiers = 0
        self.CAPITAL = win32api.GetKeyState(win32con.VK_CAPITAL)
        self.SCROLL  = win32api.GetKeyState(win32con.VK_SCROLL)
        self.NUMLOCK = win32api.GetKeyState(win32con.VK_NUMLOCK)
        
        self.hHook = None
        self.installHook()
        
        self.build_string = BuildString(self.hwnd, WM_SETIMAGE)
        
        self.run()
    
    def openImage(self, image_path: str):
        if self.image_path == image_path:
            return self.destroyWindow()
        
        self.image_path = image_path
        self.og_image = Image.open(image_path)
        
        self.scale = 1.0
        self.center_x = self.width // 2
        self.center_y = self.height // 2
        self.cursor_x = 0
        self.cursor_y = 0
        
        # make the image fit the window
        if self.og_image.width > self.width or self.og_image.height > self.height:
            self.scale = min(self.width / self.og_image.width, self.height / self.og_image.height)
            self.image = self.og_image.resize((int(self.og_image.width * self.scale), int(self.og_image.height * self.scale)), Image.Resampling.LANCZOS)
        else:
            self.image = self.og_image.copy()
        
        # Modify the window title
        user32.SetWindowTextW(self.hwnd, image_path)
        
        user32.RedrawWindow(self.hwnd, 0, 0, win32con.RDW_INVALIDATE | win32con.RDW_ERASE | win32con.RDW_UPDATENOW)
    
    def WndProc(self, hwnd, message, wParam, lParam):
        if message == win32con.WM_DESTROY:
            self.destroyWindow()
            return 0
        
        # Exit the application when the Escape/Space key is pressed
        elif message == win32con.WM_KEYDOWN:
            if wParam == win32con.VK_ESCAPE:
                self.destroyWindow()
            
            # Toggle click-through mode
            elif wParam == ord('T'):
                self.click_through: bool = not self.click_through
                self.toggle_click_through()
                
                # Force redraw to show/hide the red square
                user32.RedrawWindow(hwnd, 0, 0, win32con.RDW_INVALIDATE | win32con.RDW_ERASE | win32con.RDW_UPDATENOW)
            
            return 0
        
        elif message == win32con.WM_PAINT:
            ps = PAINTSTRUCT()
            self.hDC = user32.BeginPaint(hwnd, ctypes.byref(ps))
            
            # Use memory DC to reduce flicker
            self.update_memory_dc()
            gdi32.BitBlt(self.hDC, 0, 0, self.width, self.height, self.memory_dc, 0, 0, win32con.SRCCOPY)
            
            user32.EndPaint(hwnd, ctypes.byref(ps))
            return 0
        
        # Zoom in/out
        elif message == win32con.WM_MOUSEWHEEL:
            delta = HIWORD(wParam)
            
            # Check if shift key is pressed
            factor = 1.2
            if win32api.GetAsyncKeyState(win32con.VK_SHIFT) < 0:
                factor = 1.05
            
            if delta >> 15:
                self.scale /= factor
            else:
                self.scale *= factor

            # Ensure scale is within the limits (1% to 400% of the window size)
            min_scale = max(0.01 * self.width / self.og_image.width, 0.01 * self.height / self.og_image.height)
            max_scale = min(4.0 * self.width / self.og_image.width, 4.0 * self.height / self.og_image.height)
            self.scale = max(min_scale, min(max_scale, self.scale))
            
            new_size = (int(self.og_image.width * self.scale), int(self.og_image.height * self.scale))
            
            # Check if the new size is less than the window size, if so, reset the center of the image to the center of the window
            if new_size[0] < self.width and new_size[1] < self.height:
                self.center_x = self.width // 2
                self.center_y = self.height // 2
            
            self.image = self.og_image.resize(new_size, Image.Resampling.LANCZOS)
            
            area = wintypes.RECT(0, 0, self.width, self.height)
            user32.RedrawWindow(hwnd, ctypes.byref(area), 0, win32con.RDW_INVALIDATE | win32con.RDW_ERASE | win32con.RDW_UPDATENOW)
            
            return 0
        
        elif message == win32con.WM_LBUTTONDOWN:
            self.cursor_x = LOWORD(lParam)
            self.cursor_y = HIWORD(lParam)
            return 0
        
        elif message == win32con.WM_RBUTTONDOWN:
            windows_context_menu_file(self.image_path)
        
        elif message == win32con.WM_MOUSEMOVE:
            if wParam & win32con.MK_LBUTTON and self.cursor_x != -1:
                x = LOWORD(lParam)
                y = HIWORD(lParam)
                self.center_x += x - self.cursor_x
                self.center_y += y - self.cursor_y
                self.cursor_x = x
                self.cursor_y = y
                
                # Redraw the window
                # user32.InvalidateRect(hwnd, None, True)
                # user32.UpdateWindow(hwnd)
                user32.RedrawWindow(hwnd, 0, 0, win32con.RDW_INVALIDATE | win32con.RDW_ERASE | win32con.RDW_UPDATENOW)
            
            return 0
        
        elif message == WM_SETIMAGE:
            # Convert the LPARAM to a string
            # new_image_path = ctypes.cast(lParam, wintypes.LPCWSTR).value
            image_path = self.build_string.buildString(lParam)
            
            if image_path:
                self.openImage(image_path)
            
            return 0
        
        # Prevent background erasure to eliminate flickering
        elif message == win32con.WM_ERASEBKGND:
            return 0
        
        # Increasing the window size
        elif message == win32con.WM_SIZE:
            self.width = LOWORD(lParam)
            self.height = HIWORD(lParam)
            
            if self.maximized_or_restored:
                self.center_x = self.width // 2
                self.center_y = self.height // 2
                self.maximized_or_restored = False
            
            return 0
        
        # Reset the center of the image when the window is maximized
        elif message == win32con.WM_SYSCOMMAND:
            if wParam in (win32con.SC_MAXIMIZE, win32con.SC_RESTORE, SC_TITLE_DOUBLECLICK_MAXIMIZE, SC_TITLE_DOUBLECLICK_RESTORE):
                self.maximized_or_restored = True
            
            # Prevent the image from being dragged when clicking on the title bar
            self.cursor_x = -1
        
        return user32.DefWindowProcW(hwnd, message, wParam, lParam)
    
    def update_memory_dc(self):
        # Create a device-independent bitmap (DIB)
        if self.memory_dc:
            gdi32.SelectObject(self.memory_dc, self.hOldBitmap)
            gdi32.DeleteObject(self.hBitmap)
            gdi32.DeleteDC(self.memory_dc)
        
        self.memory_dc = gdi32.CreateCompatibleDC(self.hDC)
        self.hBitmap = gdi32.CreateCompatibleBitmap(self.hDC, self.width, self.height)
        self.hOldBitmap = gdi32.SelectObject(self.memory_dc, self.hBitmap)
        
        # Clear the memory DC with a transparent color
        # gdi32.PatBlt(self.memory_dc, 0, 0, self.width, self.height, win32con.BLACKNESS)
        
        # Draw a red square if click-through mode is enabled
        if self.click_through:
            red_brush = gdi32.CreateSolidBrush(win32api.RGB(255, 0, 0))
            gdi32.SelectObject(self.memory_dc, red_brush)
            gdi32.Ellipse(self.memory_dc, 10, 10, 30, 30)
            gdi32.DeleteObject(red_brush)

            # Draw text to indicate click-through mode
            # gdi32.SetBkMode(self.memory_dc, win32con.TRANSPARENT)
            # gdi32.SetTextColor(self.memory_dc, win32api.RGB(255, 255, 255))
            # text = "Click-through ON"
            # gdi32.TextOutW(self.memory_dc, 40, 10, text, len(text))
        
        # Convert PIL Image to DIB and copy to memory DC
        dib_dc = self.create_dib_section(self.image)
        gdi32.BitBlt(self.memory_dc, self.center_x - self.image.width // 2, self.center_y - self.image.height // 2, self.image.width, self.image.height, dib_dc, 0, 0, win32con.SRCCOPY)
        
        # Clean up
        gdi32.DeleteDC(dib_dc)

    def create_dib_section(self, image):
        # Create a device context compatible with the screen
        compatible_dc = gdi32.CreateCompatibleDC(self.hDC)

        # Define the bitmap info
        bi = BITMAPINFO()
        bi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        bi.bmiHeader.biWidth = image.width
        bi.bmiHeader.biHeight = -image.height  # Negative for top-down
        bi.bmiHeader.biPlanes = 1
        bi.bmiHeader.biBitCount = 32
        bi.bmiHeader.biCompression = 0  # BI_RGB
        bi.bmiHeader.biSizeImage = 0
        bi.bmiHeader.biXPelsPerMeter = 0
        bi.bmiHeader.biYPelsPerMeter = 0
        bi.bmiHeader.biClrUsed = 0
        bi.bmiHeader.biClrImportant = 0

        # Create the DIB section
        ppvBits = ctypes.c_void_p()
        self.hBitmap = gdi32.CreateDIBSection(compatible_dc, ctypes.byref(bi), 0, ctypes.byref(ppvBits), None, 0)
        
        # Select the bitmap into the compatible DC
        self.hOldBitmap = gdi32.SelectObject(compatible_dc, self.hBitmap)

        # Convert PIL Image to RGBA and copy data to DIB section
        img_data = image.convert("RGBA").tobytes("raw", "BGRA")
        ctypes.memmove(ppvBits.value, img_data, len(img_data))

        return compatible_dc
    
    def toggle_click_through(self):
        ex_style = user32.GetWindowLongW(self.hwnd, win32con.GWL_EXSTYLE)
        if self.click_through:
            user32.SetWindowLongW(self.hwnd, win32con.GWL_EXSTYLE, ex_style | win32con.WS_EX_LAYERED | win32con.WS_EX_TRANSPARENT)
        else:
            user32.SetWindowLongW(self.hwnd, win32con.GWL_EXSTYLE, ex_style & ~win32con.WS_EX_TRANSPARENT)
        
        user32.SetLayeredWindowAttributes(self.hwnd, 0, 255, win32con.LWA_ALPHA)
    
    @staticmethod
    def openImageFromActiveExplorer(img_viewer):
        pythoncom.CoInitialize()
        
        # Early binding/Dynamic dispatch
        # import win32com.client.gencache
        # explorer = win32com.client.gencache.EnsureDispatch("Shell.Application")
        
        # Late binding/Static dispatch
        explorer = Dispatch("Shell.Application")
        
        explorer_windows = explorer.Windows()
        
        fg_hwnd = win32gui.GetForegroundWindow()
        
        if not fg_hwnd:
            return pythoncom.CoUninitialize()
        
        try:
            curr_className = win32gui.GetClassName(fg_hwnd)
        except Exception as e:
            print(f"Error: {e}\nA problem occurred while retrieving the className of the foreground window.\n")
            
            return pythoncom.CoUninitialize()
        
        # Check other explorer windows if any.
        active_explorer = None
        if curr_className in ("WorkerW", "Progman"):
            active_explorer = explorer_windows.Item()
        
        else:
            for explorer_window in explorer_windows:
                if explorer_window.HWND == fg_hwnd:
                    active_explorer = explorer_window
                    break
        
        if not active_explorer:
            return pythoncom.CoUninitialize()
        
        imageFiles = []
        for selected_item in active_explorer.Document.SelectedItems():
            if selected_item.Path.endswith(('.png', '.jpg', '.jpeg', '.ico', ".bmp", '.gif', '.webp')):
                imageFiles.append(selected_item.Path)
        
        if not imageFiles:
            return pythoncom.CoUninitialize()
        
        img_viewer.openImage(imageFiles[0])
        
        pythoncom.CoUninitialize()
    
    def LowLevelKeyboardProc(self, nCode: int, wParam: int, lParam):
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
            extended   = bool(flags & 0x1)
            injected   = bool(flags & 0x10)
            altPressed = bool(flags & 0x20)
            transition = bool(flags & 0x82)
            
            # To get the correct key ascii value, we need first to check if the shift is pressed.
            shiftPressed = ((win32api.GetAsyncKeyState(win32con.VK_SHIFT) & 0x8000) >> 15) | (vkey_code in (win32con.VK_LSHIFT, win32con.VK_RSHIFT))
            keyAscii, keyName= getKeyAsciiAndName(vkey_code, shiftPressed)
            
            # Key down/press event.
            if wParam in (win32con.WM_KEYDOWN, win32con.WM_SYSKEYDOWN):
                if vkey_code in (win32con.VK_LEFT, win32con.VK_UP, win32con.VK_RIGHT, win32con.VK_DOWN):
                    Thread(target=self.openImageFromActiveExplorer, args=(self,)).start()
        
        return ctypes.windll.user32.CallNextHookEx(None, nCode, wParam, lParam)
    
    def installHook(self):
        CMPFUNC = ctypes.CFUNCTYPE(ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.POINTER(ctypes.c_void_p))
        self.hook_proc = CMPFUNC(self.LowLevelKeyboardProc)
        ctypes.windll.user32.SetWindowsHookExW.argtypes = (
            ctypes.c_int,
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.c_uint
        )
        self.hHook = ctypes.windll.user32.SetWindowsHookExW(
            win32con.WH_KEYBOARD_LL,
            self.hook_proc,
            win32gui.GetModuleHandle(None),
            0
        )
        if not self.hHook:
            print("Failed to install hook!")
    
    def uninstallHook(self):
        if self.hHook:
            user32.UnhookWindowsHookEx(self.hHook)
            self.hHook = None
    
    def run(self):
        # Define Window Class
        wndclass = WNDCLASSW()
        wndclass.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
        wndclass.lpfnWndProc = WNDPROC(self.WndProc)
        wndclass.cbClsExtra = wndclass.cbWndExtra = 0
        wndclass.hInstance = kernel32.GetModuleHandleW(None)
        wndclass.hIcon = user32.LoadIconW(None, IDI_APPLICATION)
        wndclass.hCursor = user32.LoadCursorW(None, IDC_ARROW)
        wndclass.hbrBackground = gdi32.GetStockObject(win32con.WHITE_BRUSH)
        wndclass.lpszMenuName = None
        wndclass.lpszClassName = self.classname

        # Register Window Class
        try:
            user32.RegisterClassW(ctypes.byref(wndclass))
        except OSError:
            print("Failed to register window class as it already exists. Continuing...")
        
        # Get the size of the window borders and title bar
        rect = wintypes.RECT()
        user32.AdjustWindowRectEx(ctypes.byref(rect), win32con.WS_OVERLAPPEDWINDOW, False, 0)
        border_width = rect.right - rect.left
        border_height = rect.bottom - rect.top

        # Create Window with adjusted size, accounting for the title bar and borders:
        self.hwnd = user32.CreateWindowExW(
            # Make the window always on top without stealing keyboard focus
            win32con.WS_EX_TOPMOST | win32con.WS_EX_NOACTIVATE,
            wndclass.lpszClassName,
            self.title,
            win32con.WS_OVERLAPPEDWINDOW,
            win32con.CW_USEDEFAULT,
            win32con.CW_USEDEFAULT,
            self.width + border_width,
            self.height + border_height,
            None, None, wndclass.hInstance, None
        )
        
        if not self.hwnd:
            raise ctypes.WinError(ctypes.get_last_error())
        
        # Show Window
        user32.ShowWindow(self.hwnd, win32con.SW_SHOWNORMAL)
        user32.UpdateWindow(self.hwnd)
        
        self.build_string.hwnd = self.hwnd
        
        # Pump Messages
        msg = wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))

    def destroyWindow(self) -> None:
        self.uninstallHook()
        if self.memory_dc:
            gdi32.SelectObject(self.memory_dc, self.hOldBitmap)
            gdi32.DeleteObject(self.hBitmap)
            gdi32.DeleteDC(self.memory_dc)
        
        user32.DestroyWindow(self.hwnd)
        user32.UnregisterClassW(self.classname, 0)
        user32.PostQuitMessage(0)

