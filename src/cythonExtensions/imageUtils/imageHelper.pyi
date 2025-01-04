def iconize() -> None:
    """Converts the selected image files from the active explorer window into icons."""
    ...


def imagesToPDF(mode=1, targetWidth=690, widthThreshold=740, minWidth=100, minHeight=100) -> None:
    """
    Combines the selected images into a PDF file.
    
    Args:
        `imageFiles -> list[str]`: List of file paths to the images to be combined into a PDF.
        `mode -> str`: Mode of operation. Use `1` for normal mode or `2` for image-resize mode to resize images to a specific width, maintaining aspect ratio.
        `targetWidth -> int`: Desired width for images in image-resize mode if they are smaller than widthThreshold.
        `widthThreshold -> int`: Width threshold to determine if resizing is necessary in image-resize mode.
        `minWidth -> int`: Minimum width of images to be included in the PDF.
        `minHeight -> int`: Minimum height of images to be included in the PDF.
    
    Notes:
        - The function filters out images that are smaller than a minimum width and height (100x100 pixels).
        - In image-resize mode, images with width less than widthThreshold are resized to targetWidth while maintaining aspect ratio.
        - Images are sorted alphabetically by their file names before being combined into a PDF.
        - The resulting PDF is saved in the directory of the first image file with a unique name.
        - If the number of images is 20 or fewer, the resulting PDF is selected in the active explorer window.
        - A sound is played upon successful creation of the PDF.
        - Temporary resized images in image-resize mode are cleaned up after the PDF is created.
    """
    ...


def invertClipboardImage() -> None:
    """Copies the image from the top of the clipboard, inverts it, then copies it back to the clipboard."""
    