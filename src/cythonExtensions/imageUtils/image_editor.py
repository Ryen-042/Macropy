import numpy as np
import cv2, PIL, win32clipboard, os, win32con, winsound, win32gui
from PIL import ImageGrab
from io import BytesIO
from datetime import datetime as dt
from glob import glob

class ImageEditor:
    def __init__(self, image=None, window_name = "Image Editor"):
        if image is None:
            # Try to get image from clipboard
            image = self.get_image_from_clipboard()

        if image is not None:
            self.original_image = image.copy()
            self.image = image.copy()
        else:
            print("No image found in the clipboard. Exiting.")
            exit()
        
        self.window_name = window_name
        self.cropping = False
        self.mouse_start = (0, 0)
        self.mouse_end = (0, 0)
        self.drawing = False
        self.color_picker_switch = False

    def get_image_from_clipboard(self):
        try:
            clipboard_image = ImageGrab.grabclipboard()
            if clipboard_image is not None:
                return cv2.cvtColor(np.array(clipboard_image), cv2.COLOR_RGB2BGR)
        except Exception as e:
            print(f"Error getting image from clipboard: {e}")
        return None

    def update_mouse_data(self, event, x, y, flags, param):
        if event == cv2.EVENT_LBUTTONDOWN:
            if flags & cv2.EVENT_FLAG_CTRLKEY:
                self.cropping = True
                self.mouse_start = (x, y)
            else:
                self.drawing = True
                self.prev_x, self.prev_y = x, y  # Store previous position for smooth drawing

        elif event == cv2.EVENT_LBUTTONUP:
            if self.cropping:
                self.mouse_end = (x, y)
                self.crop_image()
                self.cropping = False
            elif self.drawing:
                self.drawing = False
            
            cv2.imshow(self.window_name, self.image)  # Refresh the window immediately after modifying the image

        elif event == cv2.EVENT_MOUSEMOVE:
            if self.cropping:
                self.mouse_end = (x, y)
                temp_image = self.image.copy()
                cv2.rectangle(temp_image, self.mouse_start, (x, y), (0, 255, 0), 2)
                cv2.imshow(self.window_name, temp_image)
            elif self.drawing:
                cv2.line(self.image, (self.prev_x, self.prev_y), (x, y), (0, 0, 255), 5)  # Draw line between previous and current positions
                self.prev_x, self.prev_y = x, y  # Update previous position
                cv2.imshow(self.window_name, self.image)  # Refresh the window immediately after modifying the image

    def crop_image(self):
        x1, y1 = self.mouse_start
        x2, y2 = self.mouse_end
        self.image = self.image[y1:y2, x1:x2]

    def invert_colors(self):
        if self.image is not None:
            self.image = cv2.bitwise_not(self.image)

    def rotate_image(self):
        if self.image is not None:
            self.image = cv2.rotate(self.image, cv2.ROTATE_90_CLOCKWISE)

    def copy_to_clipboard(self):
        if self.image is not None:
            # Convert image to RGB format
            img_rgb = cv2.cvtColor(self.image, cv2.COLOR_BGR2RGB)

            # Convert to a PIL image
            img_pil = PIL.Image.fromarray(img_rgb)

            # Convert to DIB format
            output = BytesIO()
            img_pil.convert("RGB").save(output, format="BMP")
            data = output.getvalue()[14:]  # BMP file header is 14 bytes, so exclude it

            # Open the clipboard
            win32clipboard.OpenClipboard()

            # Empty the clipboard
            win32clipboard.EmptyClipboard()

            # Set clipboard data with the DIB format
            win32clipboard.SetClipboardData(win32con.CF_DIB, data)

            # Close the clipboard
            win32clipboard.CloseClipboard()

    def save_to_file(self):
        if self.image is not None:
            save_directory = os.path.join(os.getcwd(), "Images")
            os.makedirs(save_directory, exist_ok=True)
            filename = dt.now().strftime("%Y-%m-%d, %I.%M.%S %p") + ".png"
            cv2.imwrite(os.path.join(save_directory, filename), self.image)
            winsound.PlaySound(r"C:\Windows\Media\tada.wav", winsound.SND_FILENAME | winsound.SND_ASYNC)

    def open_file_location(self):
        list_of_files = glob(os.path.join(os.getcwd(), "Images", "*.png"))
        if list_of_files:
            os.system(f"explorer /select, \"{max(list_of_files, key=os.path.getctime)}\"")
            winsound.PlaySound(r"C:\Windows\Media\Windows Navigation Start.wav", winsound.SND_FILENAME | winsound.SND_ASYNC)
        else:
            os.startfile(os.getcwd())

    def display_hide_color_picker(self):
        if self.image is not None:
            if not self.color_picker_switch:
                self.image = np.concatenate((self.image[0:20], self.image), axis=0).astype(np.uint8)
                self.color_picker_switch = True
            else:
                self.image = self.image[20:]
                self.color_picker_switch = False

            cv2.imshow(self.window_name, self.image)  # Refresh the window immediately after modifying the image

    def open_image_dialog(self):
        image_path = cv2.File_Dialog(["*.jpg", "*.png"]).show()
        if image_path:
            self.image = cv2.imread(image_path)
            self.original_image = self.image.copy()

    def run_editor(self):
        cv2.namedWindow(self.window_name)
        cv2.imshow(self.window_name, self.image)
        cv2.setMouseCallback(self.window_name, self.update_mouse_data)

        while True:
            key = cv2.waitKey() & 0xFF

            if key == ord('c'):
                self.copy_to_clipboard()
            
            elif key == ord('v'):
                self.image = self.get_image_from_clipboard()
            
            elif key == ord('l'): # original_image
                self.image = self.original_image.copy()

            elif key == ord('i'):
                self.invert_colors()

            elif key == ord('r'):
                self.rotate_image()

            elif key == ord('s'):
                self.save_to_file()

            elif key == ord('o'):
                self.open_file_location()

            elif key == ord('w'):
                self.display_hide_color_picker()

            elif key == ord('n'):
                self.open_image_dialog()

            elif key == 27:  # ESC key to exit
                break

            cv2.imshow(self.window_name, self.image)

        cv2.destroyAllWindows()


if __name__ == "__main__":
    editor = ImageEditor()
    editor.run_editor()
