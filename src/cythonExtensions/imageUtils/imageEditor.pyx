# cython: language_level = 3str

import cv2, PIL.Image, PIL.ImageOps, PIL.ImageGrab, numpy as np
import os, ctypes, win32ui, win32con, winsound, win32clipboard
from io import BytesIO
from datetime import datetime as dt
from glob import glob


cdef list openFileDialog(int dialog_type, initial_dir, default_extension, filter, bint multiselect=False, title="File Dialog"):
    """
    Description:
        Opens a file selection dialog window and returns a list of paths to the selected files/folders.
    ---
    Parameters:
        `dialog_type -> int`:
            Specify the dialog type. `0` for a file save dialog, `1` for a file select dialog.
        
        `Initial_dir -> str`:
            A path to a directory the file dialog will open in.
        
        `default_extension -> str`:
            An extension that will be auto selected for filtering.
        
        `filter -> srt`:
            A string containing a filter that specifies acceptable files.
            
            Ex: `"Text Files (*.txt)|*.txt|All Files (*.*)|*.*|"`.
        
        `multiselect -> bool`:
            Allow selection of multiple files.
        
        `title -> str`:
            The title of the file dialog window.
    ---
    Usage:
        >>> # Api -> CreateFileDialog(FileSave_0/FileOpen_1, DefaultExtension, InitialFilename, Flags, Filter)
        >>> o = win32ui.CreateFileDialog(1, ".txt", "default.txt", 0, "Text Files (*.txt)|*.txt|All Files (*.*)|*.*|")
    """
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
    
    cdef int dialog_flags = win32con.OFN_OVERWRITEPROMPT | win32con.OFN_FILEMUSTEXIST | win32con.OFN_EXPLORER
    
    if multiselect and dialog_type:
        dialog_flags|=win32con.OFN_ALLOWMULTISELECT
    
    dialog_window = win32ui.CreateFileDialog(dialog_type, default_extension, "", dialog_flags, filter)
    dialog_window.SetOFNTitle(title)
    
    if initial_dir:
        dialog_window.SetOFNInitialDir(initial_dir)
    
    if dialog_window.DoModal()!=win32con.IDOK:
        return []
    
    return dialog_window.GetPathNames()


cdef enum kbcon:
    AS_a = 97,  VK_A = 65,  SC_A = 30
    AS_c = 99,  VK_C = 67,  SC_C = 46
    AS_k = 107, VK_K = 75,  SC_K = 37
    AS_l = 108, VK_L = 76,  SC_L = 38
    AS_n = 110, VK_N = 78,  SC_N = 49
    AS_o = 111, VK_O = 79,  SC_O = 24
    AS_r = 114, VK_R = 82,  SC_R = 19
    AS_s = 115, VK_S = 83,  SC_S = 31
    AS_t = 116, VK_T = 84,  SC_T = 20
    AS_v = 118, VK_V = 86,  SC_V = 47
    AS_w = 119, VK_W = 87,  SC_W = 17
    
    VK_0 = 48,  VK_1 = 49,  VK_2 = 50
    VK_3 = 51,  VK_4 = 52,  VK_5 = 53
    VK_6 = 54,  VK_7 = 55,  VK_8 = 56
    VK_9 = 57
    
    AS_EQUALS     = 61
    AS_PLUS       = 43
    AS_MINUS      = 45
    AS_UNDERSCORE = 95


cdef class ImageEditor:
    """
    Description:
        A simple class for managing and manipulating images.
    ---
    Attributes:
        - `pil_image -> PIL.Image.Image`: The current image as a PIL Image object.
        - `cv2_image -> np.ndarray`: The current image as a numpy ndarray.
        - `original_image -> PIL.Image.Image`: The original image as a PIL Image object.
        - `window_name -> str`: The title of the editor's windows.
        - `save_directory -> str`: The directory to save images to.
        - `shape -> tuple[int, int]`: The shape of the image that the mouse helper is used with.
        - `state -> int`: The state of the image. {0: "Equal", 1: "pil_image is newer", 2: "cv2_image is newer"}
        - `color_picker_switch -> bool`: Whether the live color picker is currently on or off.
        - `color -> tuple[int, int, int]`: The current color of the drawing tool.
        - `color_palette -> tuple[tuple[int, int, int]]`: A tuple of the colors in the color palette.
        - `line_size -> int`: The size of the drawing tool.
        - `drawing -> bool`: Whether the user is currently drawing on an image or not.
        - `cropping -> bool`: Whether the user is currently cropping an image or not.
        - `x, y -> (int, int)`: The current x and y coordinates of the mouse.
        - `px, py -> (int, int)`: The previous x and y coordinates of the mouse.
        - `mouse_selection_start -> None | tuple[int, int]`: The starting point of the mouse selection for cropping.
        - `mouse_selection_end -> None, tuple[int, int]`: The ending point of the mouse selection for cropping.
    ---
    Methods:
        - `getImageFromClipboard`: Fetches an image from the top of the clipboard and stores it into `self.pil_image`.
        - `sendToClipboard`: Copies the image from `self.pil_image` to the clipboard.
        - `pilImageToBmp`: Converts the image from `self.pil_image` to a BMP format byte array and returns the byte array.
        - `pilToCv`: Converts `self.pil_image` to a numpy `ndarray` and stores it into `self.cv2_image`.
        - `cvToPil`: Converts the `self.cv2_image` to a PIL Image and stores it into `self.cv2_image`.
        - `imageInvert`: Inverts the colors of the `self.pil_image` and sends it to the clipboard.
        - `makeTransparent`: Converts the `self.pil_image` to an RGBA image and make the background transparent.
        - `toggleColorPicker`: Toggles the live color picker on/off.
        - `updateColorBar`: Update the top strip of the `self.cv2_image` to the color of the pixel under the mouse cursor, and put text displaying the pixel color.
        - `invertColors`: Inverts the colors of `self.cv2_image`.
        - `cropImage`: Crops `self.cv2_image` based on the current mouse selection.
        - `rotateImage`: Rotates `self.cv2_image` 90 degrees clockwise.
        - `scaleImage`: Scales `cv2_image` based on the provided scale factor `(fx, fy)`.
        - `saveImage`: Saves `cv2_image` to the `save_directory`.
        - `openSaveDirectory`: Opens the specified save directory in file explorer.
        - `openImageSelectDialog`: Opens a file dialog to select an image to open and stores it into `self.original_image` and `self.pil_image`.
        - `updateMouseData`: Updates the state of the mouse cursor based on the given mouse event.
        - `runEditor`: A blocking function that runs the image editor and shows the image window.
    """
    
    cdef int x, y, px, py, state, line_size
    cdef bint save_near_module, cropping, drawing, color_picker_switch
    cdef tuple[int, int] mouse_selection_start, mouse_selection_end
    cdef tuple[int, int, int] color
    cdef tuple color_palette
    cdef pil_image, cv2_image, original_image, window_name, save_directory

    
    def __init__(self, image: PIL.Image.Image=None, window_name="Image Editor", save_near_module=True):
        if image is not None:
            if not isinstance(image, PIL.Image.Image):
                raise TypeError("Image must be a PIL Image object.")
            
            self.pil_image = image
        
        elif not self.getImageFromClipboard():
            print("No image found in the clipboard. Exiting.")
            exit()
        
        self.original_image = self.pil_image.copy()
        
        self.state = 1 # {0: "Equal", 1: "pil_image is newer", 2: "cv2_image is newer"}
        
        self.pilToCv()
        
        # self.screen_size = (1920, 1080)
        self.window_name = window_name
        self.save_directory = os.path.join(os.path.dirname(os.path.abspath(__file__)) if save_near_module else os.getcwd(), "Images")
        
        os.makedirs(self.save_directory, exist_ok=True)
        
        self.cropping = False
        self.mouse_selection_start = (0, 0)
        self.mouse_selection_end = (0, 0)
        
        self.color_palette = (
            (0,   0, 0), (0,   0, 255), (0,   255, 0), (0,   255, 255),
            (255, 0, 0), (255, 0, 255), (255, 255, 0), (255, 255, 255),
            (128, 0, 128), (128, 128, 128)
        )
        self.color = self.color_palette[1] # Blue
        
        self.line_size = 4
        self.drawing = False
        self.color_picker_switch = False
    
    cdef bint getImageFromClipboard(self):
        """Fetches an image from the top of the clipboard and stores it into `self.pil_image`."""
        
        image = PIL.ImageGrab.grabclipboard()
        
        if image is None:
            return False
        
        self.pil_image = image
        
        self.state = 1
        
        return True
    
    cdef void sendToClipboard(self):
        """Copies the image from `self.pil_image` to the clipboard."""
        
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, self.pilImageToBmp())
        win32clipboard.CloseClipboard()

    cdef pilImageToBmp(self):
        """Converts the image from `self.pil_image` to a BMP format byte array and returns the byte array."""
        
        if self.state == 2:
            self.cvToPil()
        
        # Create a BytesIO object to hold the BMP image data.
        output = BytesIO()
        
        # Convert the image to RGB and save it in BMP format to the BytesIO object.
        self.pil_image.convert('RGB').save(output, 'BMP')
        
        # Get the byte data from the BytesIO object, ignoring the first 14 bytes (header data).
        image_data = output.getvalue()[14:]
        
        # Close the BytesIO object.
        output.close()
        
        return image_data

    cdef void pilToCv(self):
        """Converts `self.pil_image` to a numpy `ndarray` and stores it into `self.cv2_image`."""
        
        if self.state in [0, 2]:
            return
        
        # self.cv2_image = np.array(image)
        
        # To prevent the new image from being cropped with the dimensions of the previous image.
        self.cropping = False
        
        # Checking if image has an alpha channel.
        if self.pil_image.mode == 'RGBA':
            # You can use cv2.COLOR_RGBA2BGRA if you are sure the image has an alpha channel.
            # Image can have alpha channel only if its format supports transparency like png.
            self.cv2_image = cv2.cvtColor(np.asarray(self.pil_image), cv2.COLOR_RGBA2BGRA) # Converting RGBA to BGRA
        else:
            self.cv2_image =  cv2.cvtColor(np.asarray(self.pil_image), cv2.COLOR_RGB2BGR)  # Converting RGB to BGR
        
        self.state = 0

    cdef void cvToPil(self):
        """Converts the `self.cv2_image` to a PIL Image and stores it into `self.pil_image`."""
        
        if self.state in [0, 1]:
            return
        
        # Checking if image has an alpha channel
        if self.cv2_image.shape[2] == 4:
            # Converting BGRA to RGBA
            self.pil_image = PIL.Image.fromarray(cv2.cvtColor(self.cv2_image, cv2.COLOR_BGRA2RGBA))
        
        else:
            # Converting BGR to RGB
            self.pil_image = PIL.Image.fromarray(cv2.cvtColor(self.cv2_image, cv2.COLOR_BGR2RGB))
        
        self.state = 0

    cdef void imageInvert(self):
        """Inverts the colors of the `self.pil_image` and sends it to the clipboard."""
        
        if self.state == 2:
            self.cvToPil()
        
        # Check for alpha channel
        if self.pil_image.mode == 'RGBA':
            # Splitting the image's channels.
            r, g, b, a = self.pil_image.split()
            
            # Merging all the channels together except for the alpha channel.
            rgb_image = PIL.Image.merge('RGB', (r, g, b))
            
            # Inverting the image colors.
            inverted_image = PIL.ImageOps.invert(rgb_image)
            
            # Splitting the inverted image's channel to combine them again with the alpha channel.
            r2, g2, b2 = inverted_image.split()
            inverted_image = PIL.Image.merge('RGBA', (r2, g2, b2, a))
        
        else:
            inverted_image = PIL.ImageOps.invert(self.pil_image)
        
        self.state = 1
        
        self.sendToClipboard()

    cdef void makeTransparent(self):
        """Converts the `self.pil_image` to an RGBA image and make the background transparent."""
        
        if self.state == 2:
            self.cvToPil()
        
        # Check if the image already has an alpha channel.
        if self.pil_image.mode == 'RGB':
            # Converting the image to RGBA by adding an extra channel.
            temp_image = self.pil_image.convert("RGBA")
        
        else:
            temp_image = self.pil_image.copy()
        
        # Returns the contents of this image as a sequence object containing pixel values.
        pixels = temp_image.getdata()
        
        output_image = []
        
        # Finding the most frequently occurring color in the image (which usually represents the background color).
        # image.getcolors(image.size[0] * image.size[1]) -> returns a list of tuples, where each tuple represents the count and color of each pixel in the image.
        # Output format is like this: [(3, (0,0,0)), (4, (255,255,255))]
        bg_color = max(temp_image.getcolors(temp_image.size[0] * temp_image.size[1]))[1]  # max() returns the tuple with the highest count.
        
        # Loop through each pixel in the image.
        for pixel in pixels:
            # If the pixel is close to the background color, make it transparent.
            if all([abs(pixel[0] - bg_color[0] < 10), abs(pixel[1] - bg_color[1] < 1), abs(pixel[2] - bg_color[2] < 1)]):
                output_image.append((255, 255, 255, 0))
            
            else:  # Otherwise, keep the pixel's original color.
                output_image.append(pixel)
        
        # Update the image with the output image list
        temp_image.putdata(output_image)
        
        self.pil_image = temp_image
        
        self.state = 1
    
    cdef void toggleColorPicker(self):
        """Toggles the live color picker on/off."""
        
        self.cropping = False
        
        if not self.color_picker_switch:
            # cv2.namedWindow("Color Picker")
            # cv2.resizeWindow("Color Picker", 250, 250)
            self.cv2_image = np.concatenate((self.cv2_image[0:20], self.cv2_image), axis=0).astype(np.uint8)
        
        else:
            # cv2.destroyWindow("Color Picker")
            self.cv2_image = self.cv2_image[20:]
        
        self.color_picker_switch = not self.color_picker_switch
        
        self.state = 2
    
    cdef void updateColorBar(self):
        """Update the top strip of the `self.cv2_image` to the color of the pixel under the mouse cursor, and put text displaying the pixel color."""
        
        if self.state == 1:
            self.pilToCv()
        
        # Updating the first 20 rows of the image with the color of the pixel under the mouse cursor.
        self.cv2_image[0:20] = self.cv2_image[0:20] * (0) + self.cv2_image[self.y, self. x]
        
        # Put text displaying the pixel color.
        cv2.putText(img=self.cv2_image, text=f"({self.x}, {self.y}) | {', '.join(self.cv2_image[self.y, self.x].astype(str))}",
                    org=(0, 15), fontFace=cv2.FONT_HERSHEY_SIMPLEX,
                    color=(~self.cv2_image[self.y, self.x]).tolist(), fontScale=0.5) # fontFace, fontScale, color, thickness
        
        self.state = 2
    
    cdef void invertColors(self):
        """Inverts the colors of `self.cv2_image`."""
        
        if self.state == 1:
            self.pilToCv()
        
        self.cv2_image = ~self.cv2_image # Or cv2.bitwise_not(self.cv2_image)
        
        self.state = 2
    
    cdef void cropImage(self):
        """Crops `self.cv2_image` based on the current mouse selection."""
        
        if self.state == 1:
            self.pilToCv()
        
        cdef int x1, y1, x2, y2
        x1, y1 = self.mouse_selection_start
        x2, y2 = self.mouse_selection_end
        
        self.cv2_image = self.cv2_image[y1:y2, x1:x2]
        
        self.state = 2
    
    cdef void rotateImage(self):
        """Rotates `self.cv2_image` 90 degrees clockwise."""
        
        self.cropping = False
        
        if self.color_picker_switch:
            self.toggleColorPicker()
        
        self.cv2_image = cv2.rotate(self.cv2_image, cv2.ROTATE_90_CLOCKWISE) # cv2.ROTATE_90_COUNTERCLOCKWISE, cv2.ROTATE_180
        
        self.state = 2
    
    cdef void scaleImage(self, double fx, double fy):
        """Scales `cv2_image` based on the provided scale factor `(fx, fy)`."""
        
        self.cropping = False
        
        if self.state == 1:
            self.pilToCv()
        
        # Interpolation methods:
        # cv2.INTER_LINEAR	: The standard bilinear interpolation, ideal for enlarged images.
        # cv2.INTER_NEAREST	: The nearest neighbor interpolation, which, though fast to run, creates blocky images.
        # cv2.INTER_AREA	: The interpolation for the pixel area, which scales down images.
        # cv2.INTER_CUBIC	: The bicubic interpolation with 4×4-pixel neighborhoods, which, though slow to run, generates high-quality instances.
        # cv2.INTER_LANCZOS4: The Lanczos interpolation with an 8×8-pixel neighborhood, which generates images of the highest quality but is the slowest to run.
        self.cv2_image = cv2.resize(self.cv2_image, (int(self.cv2_image.shape[1]*fx), int(self.cv2_image.shape[0]*fy)), interpolation=cv2.INTER_CUBIC)
        
        self.state = 2
    
    cdef void saveImage(self):
        """Saves `cv2_image` to the `save_directory`."""
        
        if self.state == 1:
            self.pilToCv()
        
        cv2.imwrite(os.path.join(self.save_directory, dt.now().strftime("%Y-%m-%d, %I.%M.%S %p") + ".png"), self.cv2_image)
        winsound.PlaySound(r"C:\Windows\Media\tada.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
    
    cdef void openSaveDirectory(self):
        """Opens the specified save directory in file explorer."""
        
        list_of_files = glob(os.path.join(self.save_directory, "*.png"))
        
        if list_of_files:
            os.system(f"explorer /select, \"{max(list_of_files, key=os.path.getctime)}\"")
            winsound.PlaySound(r"C:\Windows\Media\Windows Navigation Start.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
        else:
            os.startfile(os.getcwd())
    
    cdef void openImageSelectDialog(self):
        """Opens a file dialog to select an image to open and stores it into `self.original_image` and `self.pil_image`."""
        
        image_path = openFileDialog(1, initial_dir=self.save_directory, default_extension="", title="Select an image",
                                    multiselect=False, filter="Image Files (*.png; *.jpg; *.jpeg)|*.png;*.jpg;*.jpeg|All Files (*.*)|*.*||")
        
        print(image_path, type(image_path))
        
        if image_path:
            self.original_image = PIL.Image.open(image_path[0])
            self.pil_image = self.original_image.copy()
            self.state = 1
    
    cdef void updateMouseData(self, int event, int x, int y, int flags, param) :
        """
        Description:
            Updates the state of the mouse cursor based on the given mouse event.
        ---
        Parameters:
            - `event (int)`: The type of the current mouse event.
            - `x (int)`: The current x-coordinate of the mouse.
            - `y (int):` The current y-coordinate of the mouse.
            - `flags (int)`: The flags associated with the current mouse event.
            - `param (Any)`: Any additional parameters associated with the current mouse event.
        """
        
        self.x, self.y = x, y
        
        if event == cv2.EVENT_LBUTTONDOWN:
            # Check if the mouse is within the image boundaries
            if flags & cv2.EVENT_FLAG_CTRLKEY and (x <= self.cv2_image.shape[1] and y <= self.cv2_image.shape[0]):
                self.cropping = True
                self.mouse_selection_start = x, y
            
            else:
                self.drawing = True
                
                # Draw a circle at the current mouse position.
                cv2.circle(self.cv2_image, (x, y), self.line_size//2, self.color, -1)
                
                self.px, self.py = x, y  # Store previous position for smooth drawing

        # If the left mouse button was released, stop the current action.
        elif event == cv2.EVENT_LBUTTONUP:
            if self.cropping:
                self.mouse_selection_end = x, y
                self.cropImage()
                self.cropping = False
            
            elif self.drawing:
                self.drawing = False
        
        elif event == cv2.EVENT_MOUSEMOVE:
            if self.cropping:
                self.mouse_selection_end = (x, y)
                temp_image = self.cv2_image.copy()
                
                cv2.rectangle(temp_image, self.mouse_selection_start, (x, y), (0, 255, 0), 2)
                cv2.imshow(self.window_name, temp_image)
                return
            
            if self.drawing:
                cv2.line(self.cv2_image, (self.px, self.py), (x, y), self.color, self.line_size)  # Draw line between previous and current positions
                
                self.px, self.py = x, y
                self.state = 2
            
            if self.color_picker_switch:
                self.updateColorBar()
            
            cv2.imshow(self.window_name, self.cv2_image)
    
    def runEditor(self):
        """A blocking function that runs the image editor and shows the image window."""
        
        cdef bint show_image
        
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        # screen_size = win32api.GetSystemMetrics(0), win32api.GetSystemMetrics(1)
        
        cv2.namedWindow(self.window_name)
        
        # Always on top window.
        cv2.setWindowProperty(self.window_name, cv2.WND_PROP_TOPMOST, 1)
        
        cv2.setMouseCallback(self.window_name, self.updateMouseData)
        
        cv2.imshow(self.window_name, self.cv2_image)
        
        while True:
            show_image = True
            k = cv2.waitKey(0) & 0xFF
            
            if k in (kbcon.VK_V, kbcon.AS_v, win32con.VK_SPACE) and self.getImageFromClipboard() is not None: # "V", "v" or space
                if self.color_picker_switch:
                    self.toggleColorPicker()
                
                self.state = 1
                
                self.pilToCv()
            
            elif k in (kbcon.VK_L, kbcon.AS_l): # "L" or "l"
                if self.color_picker_switch:
                    self.toggleColorPicker()
                
                self.pil_image = self.original_image.copy()
                self.state = 1
                self.pilToCv()
            
            elif k in (kbcon.VK_A, kbcon.AS_a): # 'A' or 'a'
                self.invertColors()
            
            elif k in (kbcon.VK_R, kbcon.AS_r): # "R" or "r"
                self.rotateImage()
            
            elif k in (kbcon.VK_C, kbcon.AS_c): # "C" or "c"
                self.sendToClipboard()
                show_image = False
            
            elif k in (kbcon.VK_T, kbcon.AS_t): # "T" or "t"
                self.makeTransparent()
                self.pilToCv()
            
            elif k in (kbcon.VK_W, kbcon.AS_w): # "W" or "w"
                self.toggleColorPicker()
            
            elif k in (kbcon.AS_PLUS, kbcon.AS_EQUALS): # "+" Or "="
                # if self.cv2_image.shape[0] <= screen_size[0]/1.2 and self.cv2_image.shape[1] < screen_size[1]/1.2:
                self.scaleImage(fx=1.2, fy=1.2)
            
            elif k in (kbcon.AS_MINUS, kbcon.AS_UNDERSCORE): # "-" Or "_"
                if self.cv2_image.shape[0] > 50 and self.cv2_image.shape[1] > 50:
                    self.scaleImage(fx=0.8, fy=0.8)
                else:
                    show_image = False
            
            elif k in (kbcon.VK_S, kbcon.AS_s): # "S" or "s"
                self.saveImage()
                
                show_image = False
            
            # Open file location
            elif k in (kbcon.VK_O, kbcon.AS_o): # "O" or "o"
                self.openSaveDirectory()
                
                show_image = False
            
            # elif k in (kbcon.VK_N, kbcon.AS_n): # "N" or "n"
            #     self.openImageSelectDialog()
            
            elif kbcon.VK_0 <= k <= kbcon.VK_9: # 0-9
                self.color = self.color_palette[k - kbcon.VK_0]
                
                show_image = False
            
            # cv2.getWindowProperty() used to kill the image window after clicking the exit button in the title bar.
            elif k == win32con.VK_ESCAPE or cv2.getWindowProperty(self.window_name, cv2.WND_PROP_VISIBLE) < 1: # ESC
                break
            
            if show_image:
                if self.state == 1:
                    self.pilToCv()
                
                cv2.imshow(self.window_name, self.cv2_image)
            
        
        cv2.destroyAllWindows()
