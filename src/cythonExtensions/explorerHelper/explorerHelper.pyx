# cython: language_level = 3str

"""This extension module provides functions for manipulating the `Windows Explorer` and the `Desktop`."""

cimport cython

import win32gui, win32ui, win32con, os, win32clipboard, winsound
from win32com.client import Dispatch
from win32com.shell import shell

from cythonExtensions.commonUtils.commonUtils import ShellAutomationObjectWrapper as ShellWrapper, PThread, sendToClipboard
from cythonExtensions.windowHelper import windowHelper as winHelper


# Source: https://stackoverflow.com/questions/17984809/how-do-i-create-an-incrementing-filename-in-python
cpdef getUniqueName(directory, filename="New File", sequence_pattern=" (%s)", extension=".txt"):
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


cdef getActiveExplorer(explorer_windows=None, bint check_desktop=True):
    """Returns the active (focused) explorer/desktop window object."""
    
    cdef bint initializer_called = PThread.coInitialize()
    
    # No automation object was passed; create one.
    if not explorer_windows:
        explorer_windows = ShellWrapper.explorer.Windows()
    
    output = None
    
    cdef int fg_hwnd = win32gui.GetForegroundWindow()
    
    if not fg_hwnd:
        return None
    
    try:
        curr_className = win32gui.GetClassName(fg_hwnd)
    
    except Exception as e:
        print(f"Error: {e}\nA problem occurred while retrieving the className of the foreground window.\n")
        
        return None
    
    # Check if the active window has one of the Desktop window class names. This check is necessary because
    # `GetForegroundWindow()` and `explorer_windows.Item().HWND` might not be the same even when the Desktop is the active window.
    if check_desktop and curr_className in ("WorkerW", "Progman"):
        output = explorer_windows.Item() # Not passing a number to `Item()` returns the desktop window object.
    
    else:
        # Check other explorer windows if any.
        for explorer_window in explorer_windows:
            if explorer_window.HWND == fg_hwnd:
                output = explorer_window
                break
    
    if initializer_called:
        PThread.coUninitialize()
    
    return output


cdef getExplorerAddress(active_explorer=None):
    """Returns the address of the active explorer window."""
    
    if not active_explorer:
        active_explorer = getActiveExplorer(explorer_windows=None, check_desktop=False)
    
    if not active_explorer:
        return ""
    
    # return win32gui.GetWindowText(active_explorer.HWND) # About 10 times faster but sometimes fails to return the full path.
    return active_explorer.Document.Folder.Self.Path


cdef list getSelectedItemsFromActiveExplorer(active_explorer=None, tuple patterns=None):
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
    
    cdef bint initializer_called = PThread.coInitialize()
    
    if not active_explorer:
        active_explorer = getActiveExplorer(explorer_windows=None, check_desktop=True)
    
    cdef list output = []
    
    if active_explorer:
        for selected_item in active_explorer.Document.SelectedItems():
            if not patterns or selected_item.Path.endswith(patterns):
                output.append(selected_item.Path)
    
    if initializer_called:
        PThread.coUninitialize()
    
    return output


def copySelectedFileNames(active_explorer=None, check_desktop=True) -> list[str]:
    """
    Description:
        Copies the absolute paths of the selected files from the active explorer/desktop window.
    ---
    Parameters:
        `active_explorer -> CDispatch`:
            The active explorer window object.
        
        `check_desktop=True`:
            Whether to allow checking desktop items or not.
    ---
    Returns:
        `list[str]`: A list containing the paths to the selected items in the active explorer/desktop window.
    """
    
    cdef bint initializer_called = PThread.coInitialize()
    
    # If no automation object was passed (i.e., `None` was passed), create one.
    if not active_explorer:
        active_explorer = getActiveExplorer(explorer_windows=None, check_desktop=check_desktop)
    
    cdef list selected_files_paths = []
    
    
    if active_explorer:
        selected_files_paths = getSelectedItemsFromActiveExplorer(active_explorer)
        
        if selected_files_paths:
            concatenated_file_paths = '"' + '" "'.join(selected_files_paths) + '"'
            
            sendToClipboard(concatenated_file_paths, win32clipboard.CF_UNICODETEXT)
        
        winsound.PlaySound(r"SFX\coins-497.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
    
    if initializer_called:
        PThread.coUninitialize()
    
    return selected_files_paths


# https://learn.microsoft.com/en-us/windows/win32/dlgbox/open-and-save-as-dialog-boxes
def openFileDialog(dialog_type: int, default_extension="", default_filename="", extra_flags=0,
                   filter="", multiselect=False, title="File Dialog", initial_dir="") -> list[str]:
    """
    Description:
        Opens a file selection dialog window and returns a list of paths to the selected files/folders.
    ---
    Parameters:
        `dialog_type -> int`:
            Specify the dialog type. `0` for a file save dialog, `1` for a file select dialog.
        
        `default_extension -> str`:
            An extension that will be auto selected for filtering.
        
        `default_filename -> str`:
            A name that will be automatically typed in the file name box.
        
        `extra_flags -> int`:
            An int that will be passed to the file dialog api as a flag.
        
        `filter -> srt`:
            A string containing a filter that specifies acceptable files.
            
            Ex: `"Text Files (*.txt)|*.txt|All Files (*.*)|*.*|"`.
        
        `multiselect -> bool`:
            Allow selection of multiple files.
        
        `title -> str`:
            The title of the file dialog window.
        
        `Initial_dir -> str`:
            A path to a directory the file dialog will open in.
    ---
    Usage:
        >>> # Api -> CreateFileDialog(FileSave_0/FileOpen_1, DefaultExtension, InitialFilename, Flags, Filter)
        >>> o = win32ui.CreateFileDialog(1, ".txt", "default.txt", 0, "Text Files (*.txt)|*.txt|All Files (*.*)|*.*|")
    """
    
    cdef int dialog_flags = extra_flags|win32con.OFN_OVERWRITEPROMPT|win32con.OFN_FILEMUSTEXIST # |win32con.OFN_EXPLORER
    
    if multiselect and dialog_type:
        dialog_flags|=win32con.OFN_ALLOWMULTISELECT
    
    dialog_window = win32ui.CreateFileDialog(dialog_type, default_extension, default_filename, dialog_flags, filter)
    dialog_window.SetOFNTitle(title)
    
    if initial_dir:
        dialog_window.SetOFNInitialDir(initial_dir)
    
    if dialog_window.DoModal()!=win32con.IDOK:
        return []
    
    return dialog_window.GetPathNames()


# https://mail.python.org/pipermail/python-win32/2012-September/012533.html
@cython.wraparound(False)
def selectFilesFromDirectory(directory, file_names: list[str]) -> None:
    """Given an absolute directory path and the names of its items (names relative to the path), if an explorer window with the specified directory is present, use it, otherwise open a new one, then select all the items specified."""
    
    folder_pidl = shell.SHILCreateFromPath(directory, 0)[0]
    
    desktop = shell.SHGetDesktopFolder()
    shell_folder = desktop.BindToObject(folder_pidl, None, shell.IID_IShellFolder)
    
    cdef dict name_to_item_mapping = dict([(desktop.GetDisplayNameOf(item, 0), item) for item in shell_folder])
    
    cdef list to_show = []
    
    for file in file_names:
        if not name_to_item_mapping.get(file):
            raise Exception('File: "%s" not found in "%s"' % (file, directory))
        
        to_show.append(name_to_item_mapping[file])
    
    shell.SHOpenFolderAndSelectItems(folder_pidl, to_show, 0)


def createNewFile(active_explorer=None) -> int:
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
    
    cdef bint initializer_called = PThread.coInitialize()
    
    if not active_explorer:
        active_explorer = getActiveExplorer(explorer_windows=None, check_desktop=True)
    
    cdef int output = 0
    
    if active_explorer:
        file_fullpath = getUniqueName(directory=getExplorerAddress(active_explorer))
        
        with open(file_fullpath, 'w') as newfile:
            newfile.writelines('# Created using "File Factory"')
        
        # Selects the file and put it in edit mode: https://learn.microsoft.com/en-us/windows/win32/shell/shellfolderview-selectitem
        active_explorer.Document.SelectItem(file_fullpath, 0x1F) # 0x1F = 31 # 1|4|8|16 = 29
        
        winsound.PlaySound(r"SFX\coins-497.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
        
        output = 1
    
    if initializer_called:
        PThread.coUninitialize()
    
    return output


@cython.wraparound(False)
def officeFileToPDF(active_explorer=None, office_application="Powerpoint"):
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
    
    cdef bint initializer_called = PThread.coInitialize()
    
    if not active_explorer:
        active_explorer = getActiveExplorer(explorer_windows=None, check_desktop=False)
    
    if not active_explorer:
        if initializer_called:
            PThread.coUninitialize()
        
        return
    
    office_application_char0 = office_application[0].lower()
    cdef list selected_files_paths = getSelectedItemsFromActiveExplorer(active_explorer,
                patterns={"p": (".pptx", ".ppt"), "w": (".docx", ".doc")}.get(office_application_char0))
    
    # Check if any file exists to pop it from the selected files list before starting any office application.
    cdef int file_path_counter = 0
    
    cdef int len_selected_files = len(selected_files_paths)
    
    for file_path in selected_files_paths[:]:
        new_filepath = os.path.splitext(file_path)[0] + ".pdf"
        
        if os.path.exists(new_filepath):
            print("Next file already exists: %s" % new_filepath)
            continue
        
        selected_files_paths[file_path_counter] = file_path
        
        file_path_counter += 1
    
    else:
        selected_files_paths = selected_files_paths[:file_path_counter]
        
        if file_path_counter == 0 or file_path_counter < len_selected_files:
            winsound.PlaySound(r"SFX\wrong.swf.wav", winsound.SND_FILENAME)
    
    if not selected_files_paths:
        if initializer_called:
            PThread.coUninitialize()
        
        return
    
    winsound.PlaySound(r"SFX\connection-sound.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
    
    cdef office_dispatch = Dispatch(f"{office_application}.Application")
    
    # office_dispatch.Visible = 1 # Uncomment this if an error happened
    
    # office_dispatch.ActiveWindow.WindowState = 2 # (ppWindowNormal, ppWindowMinimized, ppWindowMaximized) = 1, 2, 3
    
    for file_path in selected_files_paths:
        new_filepath = os.path.splitext(file_path)[0] + ".pdf"
        
        # office_window = office_dispatch.Presentations.Open(file_path) if office_application_char0 == "p" else \
        #                 office_dispatch.Documents.Open(file_path)     if office_application_char0 == "w" else None
        
        office_window = (office_application_char0 == "p" and office_dispatch.Presentations.Open(file_path)) or \
                        (office_application_char0 == "w"and office_dispatch.Documents.Open(file_path))
        
        # WdSaveFormat enumeration (Word): https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveformat
        office_window.SaveAs(new_filepath, {"p": 32, "w": 17}.get(office_application_char0))
        
        print("Success: %s" % new_filepath)
        
        office_window.Close()
    
    office_dispatch.Quit()
    
    active_explorer.Document.SelectItem(new_filepath, 0x1F)
    
    if initializer_called:
        PThread.coUninitialize()
    
    winsound.PlaySound(r"SFX\coins-497.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)


def genericFileConverter(active_explorer=None, tuple patterns=None, convert_func=None, new_loc="", str new_extension="") -> None:
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
        
        `new_loc -> str`:
            The new location for the converted files. Defaults to the same location as the selected files.
        
        `new_extension -> str`:
            The new file extension for the converted files (you can treat it as a suffix added at the end of filenames).
    ---
    Examples:
    >>> # To convert image files to .ico files
    >>> genericFileConverter(None, (".png", ".jpg"), lambda f1, f2: PIL.Image.open(f1).resize((512, 512)).save(f2), new_extension=" - (512x512).ico")
    
    >>> # To convert audio files to .wav files
    >>> genericFileConverter(None, (".mp3", ), lambda f1, f2: subprocess.call(["ffmpeg", "-loglevel", "error", "-hide_banner", "-nostats",'-i', f1, f2]), new_extension=".wav")
    """
    
    if winHelper.showMessageBox("Are you sure you want to convert the selected files?", "Confirmation", 2, win32con.MB_ICONQUESTION) == 7:
        return
    
    cdef bint initializer_called = PThread.coInitialize()
    
    if not active_explorer:
        active_explorer = getActiveExplorer(explorer_windows=None, check_desktop=False)
    
    if not active_explorer:
        if initializer_called:
            PThread.coUninitialize()
        
        return
    
    cdef list selected_files_paths = getSelectedItemsFromActiveExplorer(active_explorer, patterns=patterns)
    
    if not selected_files_paths:
        if initializer_called:
            PThread.coUninitialize()
        
        return
    
    winsound.PlaySound(r"SFX\connection-sound.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
    
    for file_path in selected_files_paths:
        new_filepath = os.path.splitext(file_path)[0] + new_extension
        
        if new_loc:
            new_filepath = os.path.join(new_loc, os.path.basename(new_filepath))
        
        if os.path.exists(new_filepath):
            print("Warning, file already exists: %s" % new_filepath)
            continue
        
        convert_func(file_path, new_filepath)
        print(f"File converted: {new_filepath}")
    
    winsound.PlaySound(r"SFX\coins-497.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
    
    active_explorer.Document.SelectItem(new_filepath, 0x1F)
    
    if initializer_called:
        PThread.coUninitialize()


def flattenDirectories(active_explorer=None) -> None:
    """Flattens the selected folders from the active explorer window to the explorer current location."""
    
    cdef bint initializer_called = PThread.coInitialize()
    
    # If no automation object was passed (i.e., `None` was passed), create one.
    if not active_explorer:
        active_explorer = getActiveExplorer(explorer_windows=None, check_desktop=False)
    
    if not active_explorer:
        if initializer_called:
            PThread.coUninitialize()
        
        return
    
    cdef list selected_files_paths = getSelectedItemsFromActiveExplorer(active_explorer)
    
    if not selected_files_paths:
        if initializer_called:
            PThread.coUninitialize()
        
        return
    
    src = getExplorerAddress(active_explorer)
    
    dst = os.path.join(src, getUniqueName(src, "Flattened", extension=""))
    
    os.makedirs(dst, exist_ok=True)
    
    # for root_dir, cur_dir, files in os.walk(target_path, topdown=True):
    for folder in os.listdir(src):
        for file_name in os.listdir(os.path.join(src, folder)):
            target_src = os.path.join(src, folder, file_name)
            target_dst = os.path.join(dst, file_name)
            
            if os.path.isfile(target_src):
                os.rename(target_src, target_dst)
    
    if initializer_called:
        PThread.coUninitialize()
