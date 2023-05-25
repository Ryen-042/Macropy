"""This module provides functions for manipulating Windows Explorer and Desktop."""

from win32com.client import CDispatch
from typing import Callable, Optional

def _filter(target: str, patterns: tuple[str]) -> bool:
    """Returns `True` if the target string ends with any of the patterns, `False` otherwise."""
    ...

def GetUniqueName(directory: str, filename="New File", sequence_pattern=" (%s)", extension=".txt") -> str:
    """
    Description:
        Finds the next unused incremental filename in the specified directory.
    ---
    Usage:
        >>> GetUniqueName(directory, "New File", " (%s)", ".txt")
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
    ...

def GetActiveExplorer(explorer_windows: Optional[CDispatch], check_desktop=True) -> CDispatch:
    """Returns the active (focused) explorer/desktop window object."""
    ...

def GetExplorerAddress(active_explorer: Optional[CDispatch]) -> str:
    """Returns the address of the active explorer window."""
    ...

def GetSelectedItemsFromActiveExplorer(active_explorer: Optional[CDispatch], patterns: Optional[tuple[str]]) -> list[str]:
    """
    Description:
        Returns the absolute paths of the selected items in the active explorer window.
    ---
    Parameters:
        `active_explorer -> CDispatch`:
            The active explorer window object.
        
        `patterns -> tuple[str]`:
            A tuple containing the file extensions to filter the selected items by.
    ---
    Returns:
        `list[str]`: A list containing the paths to the selected items in the active explorer window.
    """
    ...

def CopySelectedFileNames(active_explorer: Optional[CDispatch], check_desktop=True) -> list[str]:
    """
    Description:
        Copies the absolute paths of the selected files from the active explorer/desktop window.
    ---
    Returns:
        `list[str]`: A list containing the paths to the selected items in the active explorer/desktop window.
    """
    ...

# https://learn.microsoft.com/en-us/windows/win32/dlgbox/open-and-save-as-dialog-boxes
def OpenFileDialog(dialog_type: int, default_extension="", default_filename="", extra_flags: int=0,
                   filter="", multiselect=False, title="File Dialog", initial_dir="") -> list[str]:
    """
    Description:
        Opens a file selection dialog window and returns a list of paths to the selected files/folders.
    ---
    Parameters:
        `dialog_type -> int`:
            Specify the dialog type. `0` for a file save dialog, `1` for a file select dialog.
        
        `default_filename -> str`:
            A name that will be automatically typed in the file name box.
        
        `filter -> srt`:
            A string containing a filter that specifies acceptable files.
            
            Ex: `"Text Files (*.txt)|*.txt|All Files (*.*)|*.*|"`.
        
        `multiselect -> bool`:
            Allow selection of multiple files.
    """
    ...

# https://mail.python.org/pipermail/python-win32/2012-September/012533.html
def SelectFilesFromDirectory(directory: str, file_names: list[str]) -> None:
    """Given an absolute directory path and the names of its items (names relative to the path), if an explorer window with the specified directory is present, use it, otherwise open a new one, then select all the items specified."""
    ...

def CreateFile(active_explorer: Optional[CDispatch]) -> int:
    """
    Description:
        Creates a new file with an incremental name in the active explorer/desktop window then select it.
        A file is created only if an explorer/desktop window is active (focused).
    ---
    Returns:
        - `int`: A status code that represents the success/failure of the operation.
            - `0`: `No explorer window was focused`
            - `1`: `A file was created successfully`
    """
    ...

def ImagesToPDF(active_explorer: Optional[CDispatch]) -> None:
    """Combines the selected images from the active explorer window into a PDF file with an incremental name then select it.
    Please note that the function sorts the file names alphabetically before merging."""
    ...

def OfficeFileToPDF(active_explorer: Optional[CDispatch], office_application="Powerpoint") -> None:
    """
    Description:
        Converts the selected files from the active explorer window that are associated with the specified office application into a PDF format.
    ---
    Parameters:
        `active_explorer -> CDispatch`:
            The active explorer window object.
        
        `office_application -> str`:
            Only the associated files types with this application will be converted, any other types will be ignored.
            
            Ex => `"Powerpoint": (".pptx", ".ppt")` | `"Word": (".docx", ".doc")`
    """
    ...

def GenericFileConverter(active_explorer: Optional[CDispatch], patterns: Optional[tuple[str]], convert_func: Optional[Callable[[str, str], None]], new_extension="") -> None:
    """
    Description:
        Converts the selected files from the active explorer window using the specified filter and convert functions.
    ---
    Parameters:
        `active_explorer -> CDispatch`:
            The active explorer window object.
        
        `patterns -> tuple[str]`:
            A tuple of file extensions to filter the selected files by.
        
        `convert_func -> Callable[[str, str], None]`:
            A function that takes the input file path and the output file path and performs the conversion.
        
        `new_extension -> str`:
            The new file extension for the converted files (you can treat it as a suffix added at the end of filenames).
    ---
    Examples:
    >>> # To convert image files to .ico files
    >>> GenericFileConverter(None, (".png", ".jpg"), lambda f1, f2: PIL.Image.open(f1).resize((512, 512)).save(f2), " - (512x512).ico")
    
    >>> # To convert audio files to .wav files
    >>> GenericFileConverter(None, (".mp3"), lambda f1, f2: subprocess.call(["ffmpeg", "-loglevel", "error", "-hide_banner", "-nostats",'-i', f1, f2]), ".wav")
    """
    ...

def FlattenDirectories(active_explorer: Optional[CDispatch], files_only=False) -> None:
    """Flattens the selected folders from the active explorer window to the explorer current location."""
    ...
