# cython: language_level = 3str

import win32gui, os, winsound, pythoncom
import PIL.Image
from win32com.client import Dispatch


cdef getUniqueName(directory, filename="New File", sequence_pattern=" (%s)", extension=".txt"):
    """
    Description:
        Finds the next unused incremental filename in the specified directory.
    ---
    Usage:
        >>> getUniqueName(directory, "New File", " (%s)", ".txt")
    ---
    Returns:
        `str`: A string in this format `f"{directory}\\{filename}{pattern}{extension}"`.
    ---
    Examples: (directory is omitted)
        `"New File.txt" | "New File (1).txt" | "New File (2).txt" | ...`
    
    ---
    Complexity:
        Runs in `log(n)` time where `n` is the number of existing files in sequence.
    """
    
    filename = os.path.join(directory, filename)
    if not os.path.exists(filename + extension):
        return filename + extension
    
    filename += sequence_pattern + extension
    cdef int file_counter = 1
    
    # First do an exponential search
    while os.path.exists(filename % file_counter):
        file_counter = file_counter * 2
    
    # Result lies somewhere in the interval (i/2..i]
    # We call this interval (a..b] and narrow it down until a + 1 = b
    cdef int a, b, c
    a, b = (file_counter // 2, file_counter)
    
    while a + 1 < b:
        c = (a + b) // 2 # interval midpoint
        a, b = (c, b) if os.path.exists(filename % c) else (a, c)
    
    return filename % b


def iconize():
    """Converts the selected image files from the active explorer window into icons."""
    
    pythoncom.CoInitialize()
    
    explorer = Dispatch("Shell.Application")
    
    explorer_windows = explorer.Windows()
    
    cdef int fg_hwnd = win32gui.GetForegroundWindow()
    
    if not fg_hwnd:
        print("Error: Unable to retrieve the foreground window.\n")
        winsound.PlaySound(r"SFX\wrong.swf.wav", winsound.SND_FILENAME)
        
        return
    
    
    # Check other explorer windows if any.
    cdef active_explorer
    for explorer_window in explorer_windows:
        if explorer_window.HWND == fg_hwnd:
            active_explorer = explorer_window
            break
    
    if not active_explorer:
        print("Error: Did not find any opened windows explorer.\n")
        winsound.PlaySound(r"SFX\wrong.swf.wav", winsound.SND_FILENAME)
        
        return
    
    winsound.PlaySound(r"SFX\connection-sound.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
    
    cdef list image_locations = []
    for selected_item in active_explorer.Document.SelectedItems():
        if selected_item.Path.endswith(('.png', '.jpg', '.jpeg')):
            image_locations.append(selected_item.Path)
    
    os.makedirs("Images/Icons", exist_ok=True)
    
    for image_loc in image_locations:
        image       = PIL.Image.open(image_loc)
        new_image = image.resize((512, 512))
        new_image.save(os.path.join("Images", "Icons", os.path.splitext(os.path.basename(image_loc))[0] + " - (512x512).ico"))
    
    winsound.PlaySound(r"SFX\coins-497.wav", winsound.SND_FILENAME)
    os.startfile(os.path.join(os.getcwd(), "Images", "Icons"))
    
    pythoncom.CoUninitialize()


def imagesToPDF():
    """Combines the selected images from the active explorer window into a PDF file with an incremental name then select it.
    Please note that the function sorts the file names alphabetically before merging."""
    
    import img2pdf
    
    pythoncom.CoInitialize()
    
    explorer = Dispatch("Shell.Application")
    
    explorer_windows = explorer.Windows()
    
    cdef int fg_hwnd = win32gui.GetForegroundWindow()
    
    if not fg_hwnd:
        print("Error: Unable to retrieve the foreground window.\n")
        winsound.PlaySound(r"SFX\wrong.swf.wav", winsound.SND_FILENAME)
        
        return
    
    # Check other explorer windows if any.
    cdef active_explorer = None
    for explorer_window in explorer_windows:
        if explorer_window.HWND == fg_hwnd:
            active_explorer = explorer_window
            break
    
    if not active_explorer:
        print("Error: Did not find any opened windows explorer.\n")
        winsound.PlaySound(r"SFX\wrong.swf.wav", winsound.SND_FILENAME)
        
        return
    
    winsound.PlaySound(r"SFX\connection-sound.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
    
    cdef list image_locations = []
    for selected_item in active_explorer.Document.SelectedItems():
        if selected_item.Path.endswith(('.png', '.jpg', '.jpeg')):
            image_locations.append(selected_item.Path)
    
    minWidth = 100
    minHeight = 100
    
    filtered_images = []
    for path in image_locations:
        img = PIL.Image.open(path)
        width, height = img.size
        
        if width >= minWidth and height >= minHeight:
            filtered_images.append(path)
    
    filtered_images.sort()
    
    directory = os.path.dirname(filtered_images[0])
    
    file_fullpath = getUniqueName(directory, "New PDF", " (%s)", ".pdf")
    
    with open(file_fullpath, "wb") as pdf_output_file:
        pdf_output_file.write(img2pdf.convert(filtered_images))
        
        if len(filtered_images) <= 20:
            active_explorer.Document.SelectItem(file_fullpath, 1|4|8|16)
        
    winsound.PlaySound(r"SFX\coins-497.wav", winsound.SND_FILENAME)
    
    pythoncom.CoUninitialize()


cdef void invertClipboardImage():
    """Copies the image from the top of the clipboard, inverts it, then copies it back to the clipboard."""
    
    import PIL.ImageOps, PIL.ImageGrab, win32clipboard
    from io import BytesIO
    
    image = PIL.ImageGrab.grabclipboard()

    if image is None:
        print("No image found in the clipboard. Try again after taking a screenshot.")
        return

    image = PIL.ImageOps.invert(image)

    output = BytesIO()

    image.convert('RGB').save(output, 'BMP')

    image_data = output.getvalue()[14:]

    # Close the BytesIO object.
    output.close()

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, image_data)
    win32clipboard.CloseClipboard()

    winsound.PlaySound(r"SFX\coins-497.wav", winsound.SND_FILENAME)