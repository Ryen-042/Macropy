import PIL.Image
from enum import IntEnum

class kbcon(IntEnum):
    AS_a = 97;  VK_A = 65;  SC_A = 30
    AS_c = 99;  VK_C = 67;  SC_C = 46
    AS_k = 107; VK_K = 75;  SC_K = 37
    AS_l = 108; VK_L = 76;  SC_L = 38
    AS_n = 110; VK_N = 78;  SC_N = 49
    AS_o = 111; VK_O = 79;  SC_O = 24
    AS_r = 114; VK_R = 82;  SC_R = 19
    AS_s = 115; VK_S = 83;  SC_S = 31
    AS_t = 116; VK_T = 84;  SC_T = 20
    AS_v = 118; VK_V = 86;  SC_V = 47
    AS_w = 119; VK_W = 87;  SC_W = 17
    
    VK_0 = 48;  VK_1 = 49;  VK_2 = 50
    VK_3 = 51;  VK_4 = 52;  VK_5 = 53
    VK_6 = 54;  VK_7 = 55;  VK_8 = 56
    VK_9 = 57
    
    AS_EQUALS     = 61
    AS_PLUS       = 43
    AS_MINUS      = 45
    AS_UNDERSCORE = 95


class ImageEditor:
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
    
    __slots__ = (
        "x", "y", "px", "py", "state", "line_size",                        # int
        "window_name", "save_directory",                                   # str
        "save_near_module", "cropping", "drawing", "color_picker_switch",  # bool
        "mouse_selection_start", "mouse_selection_end",                    # tuple[int, int]
        "color",                        # tuple[int, int, int]
        "color_palette",                # tuple[tuple[int, int, int]]
        "pil_image", "original_image",  # PIL.Image.Image
        "cv2_image"                     # np.ndarray
    )

    
    def __init__(self, image: PIL.Image.Image=None, window_name="Image Editor", save_near_module=True):
        ...
    
    def getImageFromClipboard(self) bool:
        """Fetches an image from the top of the clipboard and stores it into `self.pil_image`."""
        ...
    
    def sendToClipboard(self) -> None:
        """Copies the image from `self.pil_image` to the clipboard."""
        ...
    
    def pilImageToBmp(self) -> bytes:
        """Converts the image from `self.pil_image` to a BMP format byte array and returns the byte array."""
        ...

    def pilToCv(self) -> None:
        """Converts `self.pil_image` to a numpy `ndarray` and stores it into `self.cv2_image`."""
        ...

    def cvToPil(self) -> None:
        """Converts the `self.cv2_image` to a PIL Image and stores it into `self.pil_image`."""
        ...

    def imageInvert(self) -> None:
        """Inverts the colors of the `self.pil_image` and sends it to the clipboard."""
        ...

    def makeTransparent(self) -> None:
        """Converts the `self.pil_image` to an RGBA image and make the background transparent."""
        ...
    
    def toggleColorPicker(self) -> None:
        """Toggles the live color picker on/off."""
        ...
    
    def updateColorBar(self) -> None:
        """Update the top strip of the `self.cv2_image` to the color of the pixel under the mouse cursor, and put text displaying the pixel color."""
        ...
    
    def invertColors(self) -> None:
        """Inverts the colors of `self.cv2_image`."""
        ...
    
    def cropImage(self) -> None:
        """Crops `self.cv2_image` based on the current mouse selection."""
        ...
    
    def rotateImage(self) -> None:
        """Rotates `self.cv2_image` 90 degrees clockwise."""
        ...
    
    def scaleImage(self, fx: float, fy: float) -> None:
        """Scales `cv2_image` based on the provided scale factor `(fx, fy)`."""
        ...
    
    def saveImage(self) -> None:
        """Saves `cv2_image` to the `save_directory`."""
        ...
    
    def openSaveDirectory(self) -> None:
        """Opens the specified save directory in file explorer."""
        ....
    
    def openImageSelectDialog(self) -> None:
        """Opens a file dialog to select an image to open and stores it into `self.original_image` and `self.pil_image`."""
        ...
    
    def updateMouseData(self, event: int, x: int, y: int, flags: int, param)  -> None:
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
        ...
    
    def runEditor(self):
        """A blocking function that runs the image editor and shows the image window."""
        ...
