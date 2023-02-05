from __future__ import annotations
import PIL.ImageOps, PIL.Image, PIL.ImageGrab
from io import BytesIO
import win32clipboard
from glob import glob
import os
import numpy as np
import cv2
import winsound #  win32api, win32con
from datetime import datetime as dt

os.makedirs(os.path.join(os.path.dirname(os.path.abspath(__file__)), "Images"), exist_ok=True)
# os.chdir(os.path.join(os.path.dirname(os.path.abspath(__file__)), "Images"))

# Parameter must be a PIL image 
def pil_image_to_bmp(image: PIL.Image):
    output = BytesIO()
    image.convert('RGB').save(output, 'BMP')
    data = output.getvalue()[14:]
    output.close()
    return data

def send_to_clipboard(image: np.ndarray, is_cv_image = True):
    if is_cv_image:
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

def pil_to_cv(image: PIL.Image):
    cv2_main_image = np.asarray(image) # np.array(image)
    if image.mode == 'RGBA':
        return cv2.cvtColor(cv2_main_image, cv2.COLOR_RGBA2BGRA) # You can use cv2.COLOR_RGBA2BGRA if you are sure you have an alpha channel. You will only have alpha channel if your image format supports transparency.
    else:
        return cv2.cvtColor(cv2_main_image, cv2.COLOR_RGB2BGR)

def cv_to_pil(cv2_main_image: np.ndarray):
    if cv2_main_image.shape[2] == 4:
        return PIL.Image.fromarray(cv2.cvtColor(cv2_main_image, cv2.COLOR_BGRA2RGBA))
    else:
        return PIL.Image.fromarray(cv2.cvtColor(cv2_main_image, cv2.COLOR_BGR2RGB))

def image_invert(image):
    # Check for alpha channel
    if image.mode == 'RGBA':
        r,g,b,a = image.split()
        rgb_image = PIL.Image.merge('RGB', (r,g,b))
        inverted_image = PIL.ImageOps.invert(rgb_image)
        r2,g2,b2 = inverted_image.split()
        inverted_image = PIL.Image.merge('RGBA', (r2,g2,b2,a))
    else:
        inverted_image = PIL.ImageOps.invert(image)
    send_to_clipboard(inverted_image, False)

def make_transparent(image: PIL.Image):
    image = image.convert("RGBA")
    data_items = image.getdata()
    
    output_image = []
    bg_color = max(image.getcolors(image.size[0] * image.size[1]))[1] # ex: (183, (255, 79, 79))
    for item in data_items:
        # if item[0] == 255 and item[1] == 255 and item[2] == 255:
            # output_image.append((255, 255, 255, 0))
        if all([abs(item[0]-bg_color[0] < 10), abs(item[1]-bg_color[1] < 1), abs(item[2]-bg_color[2] < 1)]):
            output_image.append((255, 255, 255, 0))
        else:
            output_image.append(item)
    
    image.putdata(output_image)
    return image

# Avoid global variables by storing as class data:
class MouseHelper:
    def __init__(self, shape):
        self.x = 0
        self.y = 0
        self.xy = (0, 0)
        self.shape = shape
        self.mouse_selection_start = None
        self.mouse_selection_end = None
        self.event = None
    
    # Mouse Callback function
    def update_life_data(self, event, x, y, flags, param):
        self.x = x
        self.y = y
        self.xy = (x, y)
        self.event = event
        
        if event == cv2.EVENT_LBUTTONDOWN:
            self.mouse_selection_start = x, y
        
        # Mouse is moving
        # elif event == cv2.EVENT_MOUSEMOVE: if cropping == True: x_end, y_end = x, yif event == cv2.EVENT_MOUSEMOVE:
        # if event == cv2.EVENT_MOUSEMOVE:
        #   # Modify the mouse cursor
        #     win32api.SetCursor(win32api.LoadCursor(0, win32con.IDC_SIZEALL))
        
        # If the left mouse button was released
        elif event == cv2.EVENT_LBUTTONUP:
            if (x <= self.shape[1] and y <= self.shape[0]) and (x - self.mouse_selection_start[0] > 10) and (y - self.mouse_selection_start[1] > 10):
                self.mouse_selection_end = x, y

def life_color_update(cv2_main_image, x, y):
    cv2_main_image[0:20] = cv2_main_image[0:20] * (0) + cv2_main_image[y, x]
    cv2.putText(img=cv2_main_image, text=f"({x}, {y}) | {', '.join(cv2_main_image[y, x].astype(str))}", org=(0, 15), fontFace=cv2.FONT_HERSHEY_SIMPLEX, color=(~cv2_main_image[y, x]).tolist(), fontScale=0.5) #fontFace, fontScale, color, thickness
    return cv2_main_image



def BeginImageProcessing(show_window=False, screen_size=(1920, 1080)):
    image = PIL.ImageGrab.grabclipboard()
    if show_window:
        # Convert PIL image to OpenCv
        cv2_main_image = pil_to_cv(image)
        print(cv2_main_image.shape)
        
        
        # inverted = False # Color inversion status
        cv2.namedWindow("Image Window")
        cv2.imshow("Image Window", cv2_main_image)
        
        # Always on top window.
        cv2.setWindowProperty("Image Window", cv2.WND_PROP_TOPMOST, 1)
        
        color_picker_switch = False
        
        # Set mouse callback function
        mouse_helper = MouseHelper(cv2_main_image.shape)
        cv2.setMouseCallback('Image Window', mouse_helper.update_life_data)
        
        while 1:
            k = cv2.waitKey(1) & 0xFF
            
            # cv2.getWindowProperty() used to kill the image window after clicking the exit button in the title bar.
            if k == 27 or cv2.getWindowProperty('Image Window',cv2.WND_PROP_VISIBLE) < 1: # ESC
                break
            
            if mouse_helper.mouse_selection_start and mouse_helper.mouse_selection_end:
                cv2_main_image = cv2_main_image[mouse_helper.mouse_selection_start[1]:mouse_helper.mouse_selection_end[1],
                                                mouse_helper.mouse_selection_start[0]:mouse_helper.mouse_selection_end[0]]
                mouse_helper.shape = cv2_main_image.shape
                mouse_helper.mouse_selection_start, mouse_helper.mouse_selection_end = None, None
                cv2.imshow('Image Window', cv2_main_image)
            
            # Invert colors
            elif k in [97, 65]: # 'A' or 'a'
                cv2_main_image = cv2.bitwise_not(cv2_main_image) # Or `~cv2_main_image`
                cv2.imshow("Image Window", cv2_main_image)
            
            # Just for testing
            elif k in [98, 66]: # 'B' or 'b'
                cv2_main_image = np.concatenate((cv2_main_image[0:10, :] * (0) + (150, 0, 150), cv2_main_image), axis=0).astype(np.uint8)
                cv2.imshow("Image Window", cv2_main_image)
            
            # Rotate image
            elif k in [82, 114]: # "R" or "r"
                mouse_helper.x, mouse_helper.y = 0, 0
                mouse_helper.xy = 0, 0
                cv2_main_image = cv2.rotate(cv2_main_image, cv2.ROTATE_90_CLOCKWISE) # cv2.ROTATE_90_COUNTERCLOCKWISE, cv2.ROTATE_180
                mouse_helper.shape = cv2_main_image.shape
                cv2.imshow("Image Window", cv2_main_image)
            
            # Copy image
            elif k in [67, 99]: # "C" or "c"
                send_to_clipboard(cv2_main_image)
                
            # Paste image from clipboard
            elif k in [32, 86, 118]: # Space, "V" or "v"
                image = PIL.ImageGrab.grabclipboard()
                cv2_main_image = pil_to_cv(image)
                cv2.imshow("Image Window", cv2_main_image)
            
            # Make the image transparent
            elif k in [84, 116]: # "T" or "t"
                image = cv_to_pil(cv2_main_image)
                if image.mode == 'RGBA':
                    image = image.convert("RGB")
                else:
                    image = make_transparent(cv_to_pil(cv2_main_image))
                cv2_main_image = pil_to_cv(image)
                cv2.imshow("Image Window", cv2_main_image)
            
            # Display/hide live color picker
            elif k in [87, 119]: # "W" or "w"
                if not color_picker_switch:
                    cv2_main_image = np.concatenate((cv2_main_image[0:20], cv2_main_image), axis=0).astype(np.uint8)
                    color_picker_switch = True
                else:
                    # cv2.destroyWindow("Color Picker")
                    cv2_main_image = cv2_main_image[20:]
                    color_picker_switch = False
                    cv2.imshow("Image Window", cv2_main_image)
                mouse_helper.shape = cv2_main_image.shape
            
            # Scale image up
            elif k in [43, 61]: # "+" Or "="
                if cv2_main_image.shape[0] <= screen_size[0]/1.2 and cv2_main_image.shape[1] < screen_size[1]/1.2:
                    cv2_main_image = cv2.resize(cv2_main_image, (0, 0), fx=1.2, fy=1.2)
                    cv2.imshow("Image Window", cv2_main_image)
                    mouse_helper.shape = cv2_main_image.shape
            
            # Scale image down
            elif k in [45, 95]: # "-" Or "_"
                if cv2_main_image.shape[0] > 100 and cv2_main_image.shape[1] > 100:
                    cv2_main_image = cv2.resize(cv2_main_image, (0, 0), fx=0.8, fy=0.8)
                    cv2.imshow("Image Window", cv2_main_image)
                    mouse_helper.shape = cv2_main_image.shape
            
            # Save the image to a file
            elif k in [83, 115]: # "S" or "s"
                cv2.imwrite(os.path.join(os.path.dirname(os.path.abspath(__file__)), "Images", dt.now().strftime("%Y-%m-%d, %I.%M.%S %p") + ".png"), cv2_main_image)
                winsound.PlaySound(r"C:\Windows\Media\tada.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
            
            # Open file location
            elif k in [79, 111]: # "O" or "o"
                list_of_files = glob(os.path.join(os.path.dirname(os.path.abspath(__file__)), "Images", "*.png"))
                if list_of_files:
                    os.system(f"explorer /select, \"{max(list_of_files, key=os.path.getctime)}\"")
                    winsound.PlaySound(r"C:\Windows\Media\Windows Navigation Start.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
                else:
                    os.startfile(os.getcwd())
            
            if color_picker_switch:
                cv2_main_image = life_color_update(cv2_main_image, *mouse_helper.xy)
                cv2.imshow("Image Window", cv2_main_image)
        
        cv2.destroyAllWindows()
    else:
        image_invert(image)

if __name__ == '__main__':
    from systemHelper import EnableDPI_Awareness
    from win32api import GetSystemMetrics
    x, y = GetSystemMetrics(0), GetSystemMetrics(1)
    EnableDPI_Awareness()
    BeginImageProcessing(show_window=True, screen_size=(x, y))
