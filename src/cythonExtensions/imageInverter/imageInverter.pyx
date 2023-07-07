# cython: embedsignature = True
# cython: language_level = 3str

import PIL.ImageOps, PIL.Image, PIL.ImageGrab
from io import BytesIO
import win32clipboard
from glob import glob
import os
import numpy as np
import cv2
import winsound
from datetime import datetime as dt


def pil_image_to_bmp(image: PIL.Image.Image) -> bytes:
    """
    Description:
        Converts a PIL image object to a BMP format byte array.
    ---
    Parameters:
        image: A PIL image object.
    ---
    Returns:
        A byte array of the BMP image data.
    ---
    Raises:
        TypeError: If `image` is not a valid PIL image object.
    """
    
    # Check that `image` is a valid PIL image object
    if not isinstance(image, PIL.Image.Image):
        raise TypeError("Image must be a PIL Image object.")
    
    # Create a BytesIO object to hold the BMP image data.
    output = BytesIO()
    
    # Convert the image to RGB and save it in BMP format to the BytesIO object.
    image.convert('RGB').save(output, 'BMP')
    
    # Get the byte data from the BytesIO object, ignoring the first 14 bytes (header data).
    data = output.getvalue()[14:]
    
    # Close the BytesIO object.
    output.close()
    
    return data


def send_to_clipboard(image: np.ndarray | PIL.Image.Image) -> None:
    """
    Description:
        Copies the given image to the clipboard.
    ---
    Parameters:
        - image: The image to be copied to the clipboard. Either a numpy array or a PIL image object.
    ---
    Raises:
        TypeError: If `image` is not a valid numpy array or PIL image object.
    """
    
    # Check that `image` is a valid numpy array or PIL image object
    if not isinstance(image, (np.ndarray, PIL.Image.Image)):
        raise TypeError("image must be a valid numpy array or PIL Image object")
    
    elif isinstance(image, np.ndarray):
        # Convert BGR Channels to RGB Channels (PIL uses RGB while OpenCv uses BGR)
        # image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
        
        # Convert cv image to PIL image
        image = cv_to_pil(image)
    
    # Convert to bitmap for clipboard
    image = pil_image_to_bmp(image)
    
    # Copy to clipboard
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, image)
    win32clipboard.CloseClipboard()


def pil_to_cv(image: PIL.Image.Image) -> np.ndarray:
    """
    Description:
        Converts a PIL Image to a numpy ndarray in OpenCV format.
    ---
    Parameters:
        - image (PIL.Image.Image): the input PIL Image.
    ---
    Returns:
        np.ndarray: the converted image in OpenCV format.
    """
    
    # Converting PIL Image to numpy ndarray
    cv2_main_image = np.asarray(image) # np.array(image)
    
    # Checking if image has an alpha channel.
    if image.mode == 'RGBA':
        # You can use cv2.COLOR_RGBA2BGRA if you are sure the image has alpha channel.
        # Image can have alpha channel only its format supports transparency like png.
        return cv2.cvtColor(cv2_main_image, cv2.COLOR_RGBA2BGRA) # Converting RGBA to BGRA
    else:
        return cv2.cvtColor(cv2_main_image, cv2.COLOR_RGB2BGR)   # Converting RGB to BGR


def cv_to_pil(cv2_main_image: np.ndarray) -> PIL.Image.Image:
    """
    Description:
        Converts a numpy ndarray in OpenCV format to a PIL Image.
    ---
    Parameters:
        - cv2_main_image (np.ndarray): the input image in OpenCV format.
    ---
    Returns:
        PIL.Image.Image: the converted PIL Image.
    """
    
    # Checking if image has an alpha channel
    if cv2_main_image.shape[2] == 4:
        # Converting BGRA to RGBA
        return PIL.Image.fromarray(cv2.cvtColor(cv2_main_image, cv2.COLOR_BGRA2RGBA))
    else:
        # Converting BGR to RGB
        return PIL.Image.fromarray(cv2.cvtColor(cv2_main_image, cv2.COLOR_BGR2RGB))


def image_invert(image: PIL.Image.Image) -> None:
    """
    Description:
        Inverts the colors of the input image and sends it to clipboard.
    ---
    Parameters:
        - image (PIL.Image.Image): The image to be inverted.
    """
    
    # Check for alpha channel
    if image.mode == 'RGBA':
        # Splitting the image's channels.
        r, g, b, a = image.split()
        
        # Merging all the channels together except for the alpha channel.
        rgb_image = PIL.Image.merge('RGB', (r, g, b))
        
        # Inverting the image colors.
        inverted_image = PIL.ImageOps.invert(rgb_image)
        
        # Splitting the inverted image's channel to combine them again with the alpha channel.
        r2, g2, b2 = inverted_image.split()
        inverted_image = PIL.Image.merge('RGBA', (r2, g2, b2, a))
    else:
        inverted_image = PIL.ImageOps.invert(image)
    send_to_clipboard(inverted_image)


def make_transparent(image: PIL.Image.Image) -> PIL.Image.Image:
    """
    Description:
        Convert image to RGBA and make the background transparent.
    ---
    Parameters:
        - image (PIL.Image.Image): the input image to be made transparent.
    ---
    Returns:
        PIL.Image.Image: the modified image with a transparent background.
    """
    
    # Converting the image to RGBA by adding an extra channel.
    image = image.convert("RGBA")
    
    # Returns the contents of this image as a sequence object containing pixel values.
    data_items = image.getdata()
    
    output_image = []
    
    # Finding the most frequently occurring color in the image (Which usually represet the bg color). `image.getcolors(image.size[0] * image.size[1])` returns
    # a list of tuples, where each tuple represents the count and color of each pixel in the image. The output is like this: [(3, (0,0,0)), (4, (255,255,255))]
    bg_color = max(image.getcolors(image.size[0] * image.size[1]))[1] # max() returns the tuple with the highest count.
    
    # Loop through each pixel in the image.
    for item in data_items:
        # if item[0] == 255 and item[1] == 255 and item[2] == 255:
        # output_image.append((255, 255, 255, 0))
        
        # If the pixel is close to the background color, make it transparent.
        if all([abs(item[0]-bg_color[0] < 10), abs(item[1]-bg_color[1] < 1), abs(item[2]-bg_color[2] < 1)]):
            output_image.append((255, 255, 255, 0))
        
        else: # Otherwise, keep the pixel's original color.
            output_image.append(item)
    
    # Update the image with the output image list
    image.putdata(output_image)
    
    return image


# Avoid using global variables by storing them as class members:
class MouseHelper:
    """
    Description:
        A class that helps with mouse event handling in OpenCV.
    ---
    Attributes:
        - `x (int)`: The current x-coordinate of the mouse.
        - `y (int)`: The current y-coordinate of the mouse.
        - `px (int)`: The previous x-coordinate of the mouse.
        - `py (int)`: The previous y-coordinate of the mouse.
        - `shape (tuple[int, int])`: The shape of the image that the mouse helper is used with.
        - `mouse_selection_start (Union[None, tuple[int, int]])`: The starting point of the mouse selection for cropping.
        - `mouse_selection_end (Union[None, tuple[int, int]])`: The ending point of the mouse selection for cropping.
        - `cropping (bool)`: Whether the user is currently cropping an image or not.
        - `drawing (bool)`: Whether the user is currently drawing on an image or not.
        - `event (int)`: The type of the current mouse event.
    Methods
    -------
        - update_life_data(event, x, y, flags, param)
        
        - Updates the state of the mouse helper based on the given mouse event.
    """
    
    def __init__(self, shape: tuple):
        """
        Parameters
        ----------
        shape : tuple[int, int]
            The shape of the image that the mouse helper is used with.
        """
        
        self.x, self.y = 0, 0
        self.px, self.py = 0, 0
        self.shape = shape
        self.mouse_selection_start = None
        self.mouse_selection_end = None
        self.cropping = False
        
        self.event = None
        self.drawing = False
    
    # Mouse Callback function
    def update_life_data(self, event: int, x: int, y: int, flags: int, param) -> None:
        """
        Description:
            Updates the state of the mouse helper based on the given mouse event.
        ---
        Parameters:
            - `event (int)`: The type of the current mouse event.
            - `x (int)`: The current x-coordinate of the mouse.
            - `y (int):` The current y-coordinate of the mouse.
            - `flags (int)`: The flags associated with the current mouse event.
            - `param (Any)`: Any additional parameters associated with the current mouse event.
        """
        
        self.px, self.py = self.x, self.y
        self.x, self.y = x, y
        
        self.event = event
        
        if event == cv2.EVENT_LBUTTONDOWN:
            if flags & cv2.EVENT_FLAG_CTRLKEY:  # Check if Ctrl key is pressed
                self.cropping = True
                self.drawing = False
                self.mouse_selection_start = x, y # Cropping
            
            else:
                self.cropping = False
                self.drawing = True
        
        # If the left mouse button was released
        elif event == cv2.EVENT_LBUTTONUP:
            # Cropping
            if self.cropping and self.mouse_selection_start and (x <= self.shape[1] and y <= self.shape[0]) and (x - self.mouse_selection_start[0] > 10) and (y - self.mouse_selection_start[1] > 10):
                self.mouse_selection_end = x, y
            
            self.cropping = False
            self.drawing = False


def life_color_update(cv2_main_image: np.ndarray, x: int, y: int) -> np.ndarray:
    """
    Description:
        Update the top strip of the image to the color of the pixel at (x,y), and put text displaying the pixel color.
    ---
    Parameters:
        - cv2_main_image (np.ndarray): The OpenCV image array.
        
        - x (int): The x coordinate of the pixel to use for the top strip color.
        
        - y (int): The y coordinate of the pixel to use for the top strip color.
    ---
    Returns:
        np.ndarray: The updated OpenCV image array.
    """
    
    # Updating the first 20 rows of the image with the color of the pixel at (x,y).
    cv2_main_image[0:20] = cv2_main_image[0:20] * (0) + cv2_main_image[y, x]
    
    # Put text displaying the pixel color.
    cv2.putText(img=cv2_main_image, text=f"({x}, {y}) | {', '.join(cv2_main_image[y, x].astype(str))}",
                org=(0, 15), fontFace=cv2.FONT_HERSHEY_SIMPLEX,
                color=(~cv2_main_image[y, x]).tolist(), fontScale=0.5) # fontFace, fontScale, color, thickness
    
    return cv2_main_image


def BeginImageProcessing(show_window=False, save_near_module=True, screen_size=(1920, 1080)):
    if save_near_module:
        save_directory = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Images")
    else:
        save_directory = os.path.join(os.getcwd(), "Images")
    
    os.makedirs(save_directory, exist_ok=True)
    
    image = PIL.ImageGrab.grabclipboard()
    if show_window:
        import ctypes
        import win32api, win32con
        from cythonExtensions.commonUtils.commonUtils import KB_Con as kbcon
        
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        screen_size = win32api.GetSystemMetrics(0), win32api.GetSystemMetrics(1)
        
        # Convert PIL image to OpenCv
        cv2_main_image = pil_to_cv(image)
        
        # inverted = False # Color inversion status
        cv2.namedWindow("Image Window")
        cv2.imshow("Image Window", cv2_main_image)
        
        # These two lines will force your "Main View" window to be on top with focus.
        # cv2.setWindowProperty("Image Window", cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)
        # cv2.setWindowProperty("Image Window", cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_NORMAL)
        
        # Always on top window.
        cv2.setWindowProperty("Image Window", cv2.WND_PROP_TOPMOST, 1)
        
        color_picker_switch = False
        # color_picker_image = np.zeros((30, 250, 3), np.uint8)
        # color_picker_image[:] = (150, 0, 150)
        
        # Set mouse callback function
        mouse_helper = MouseHelper(cv2_main_image.shape)
        cv2.setMouseCallback('Image Window', mouse_helper.update_life_data)
        
        while 1:
            k = cv2.waitKey(1) & 0xFF
            
            # cv2.getWindowProperty() used to kill the image window after clicking the exit button in the title bar.
            if k == win32con.VK_ESCAPE or cv2.getWindowProperty('Image Window', cv2.WND_PROP_VISIBLE) < 1: # ESC
                break
            
            # Cropping
            if mouse_helper.cropping and mouse_helper.mouse_selection_start and mouse_helper.mouse_selection_end:
                color_picker_switch = False
                cv2_main_image = cv2_main_image[mouse_helper.mouse_selection_start[1]:mouse_helper.mouse_selection_end[1],
                                                mouse_helper.mouse_selection_start[0]:mouse_helper.mouse_selection_end[0]]
                mouse_helper.shape = cv2_main_image.shape
                mouse_helper.mouse_selection_start, mouse_helper.mouse_selection_end = None, None
                cv2.imshow('Image Window', cv2_main_image)
            
            elif mouse_helper.drawing:
                # cv2.circle(cv2_main_image, (mouse_helper.x, mouse_helper.y), 5, (0, 0, 255), -1)
                
                cv2.line(cv2_main_image, (mouse_helper.px, mouse_helper.py), (mouse_helper.x, mouse_helper.y), (0, 0, 255), 4)
                cv2.imshow('Image Window', cv2_main_image)
            
            # Invert colors
            elif k in (kbcon.VK_A, kbcon.AS_a): # 'A' or 'a'
                cv2_main_image = cv2.bitwise_not(cv2_main_image) # Or `~cv2_main_image`
                cv2.imshow("Image Window", cv2_main_image)
                # inverted = not inverted
                # send_to_clipboard(image)
                # send_to_clipboard(cv2_main_image)
            
            # Rotate image
            elif k in (kbcon.VK_R, kbcon.AS_r): # "R" or "r"
                mouse_helper.x, mouse_helper.y = 0, 0
                mouse_helper.px, mouse_helper.py = 0, 0
                cv2_main_image = cv2.rotate(cv2_main_image, cv2.ROTATE_90_CLOCKWISE) # cv2.ROTATE_90_COUNTERCLOCKWISE, cv2.ROTATE_180
                mouse_helper.shape = cv2_main_image.shape
                cv2.imshow("Image Window", cv2_main_image)
                # image = image.rotate(90, expand=1)
            
            # Copy image
            elif k in (kbcon.VK_C, kbcon.AS_c): # "C" or "c"
                send_to_clipboard(cv2_main_image)
                
            # Paste image from clipboard
            elif k in (kbcon.VK_V, kbcon.AS_v, win32con.VK_SPACE): # "V", "v" or space
                image = PIL.ImageGrab.grabclipboard()
                cv2_main_image = pil_to_cv(image)
                cv2.imshow("Image Window", cv2_main_image)
            
            # Make the image transparent
            elif k in (kbcon.VK_T, kbcon.AS_t): # "T" or "t"
                image = cv_to_pil(cv2_main_image)
                if image.mode == 'RGBA':
                    image = image.convert("RGB")
                else:
                    image = make_transparent(cv_to_pil(cv2_main_image))
                cv2_main_image = pil_to_cv(image)
                cv2.imshow("Image Window", cv2_main_image)
            
            # Display/hide live color picker
            elif k in (kbcon.VK_W, kbcon.AS_w): # "W" or "w"
                if not color_picker_switch:
                    # cv2.namedWindow("Color Picker")
                    # cv2.resizeWindow("Color Picker", 250, 250)
                    cv2_main_image = np.concatenate((cv2_main_image[0:20], cv2_main_image), axis=0).astype(np.uint8)
                    color_picker_switch = True
                else:
                    # cv2.destroyWindow("Color Picker")
                    cv2_main_image = cv2_main_image[20:]
                    color_picker_switch = False
                    cv2.imshow("Image Window", cv2_main_image)
                mouse_helper.shape = cv2_main_image.shape
            
            # Scale image up
            elif k in (kbcon.AS_PLUS, kbcon.AS_EQUALS): # "+" Or "="
                # if cv2_main_image.shape[0] <= screen_size[0]/1.2 and cv2_main_image.shape[1] < screen_size[1]/1.2:
                
                cv2_main_image = cv2.resize(cv2_main_image, (0, 0), fx=1.2, fy=1.2)
                cv2.imshow("Image Window", cv2_main_image)
                mouse_helper.shape = cv2_main_image.shape
            
            # Scale image down
            elif k in (kbcon.AS_MINUS, kbcon.AS_UNDERSCORE): # "-" Or "_"
                if cv2_main_image.shape[0] > 100 and cv2_main_image.shape[1] > 100:
                    cv2_main_image = cv2.resize(cv2_main_image, (0, 0), fx=0.8, fy=0.8)
                    cv2.imshow("Image Window", cv2_main_image)
                    mouse_helper.shape = cv2_main_image.shape
            
            # Save the image to a file
            elif k in (kbcon.VK_S, kbcon.AS_s): # "S" or "s"
                cv2.imwrite(os.path.join(save_directory, dt.now().strftime("%Y-%m-%d, %I.%M.%S %p") + ".png"), cv2_main_image)
                winsound.PlaySound(r"C:\Windows\Media\tada.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
            
            # Open file location
            elif k in (kbcon.VK_O, kbcon.AS_o): # "O" or "o"
                list_of_files = glob(os.path.join(save_directory, "*.png"))
                if list_of_files:
                    os.system(f"explorer /select, \"{max(list_of_files, key=os.path.getctime)}\"")
                    winsound.PlaySound(r"C:\Windows\Media\Windows Navigation Start.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
                else:
                    os.startfile(os.getcwd())
            
            if color_picker_switch:
                cv2_main_image = life_color_update(cv2_main_image, mouse_helper.x, mouse_helper.y)
                cv2.imshow("Image Window", cv2_main_image)
        
        cv2.destroyAllWindows()
    else:
        image_invert(image)
        winsound.PlaySound(r"SFX\coins-497.wav", winsound.SND_FILENAME)
