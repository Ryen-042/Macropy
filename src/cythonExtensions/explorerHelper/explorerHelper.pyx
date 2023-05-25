# cython: embedsignature = True
# cython: language_level = 3str

"""This extension module provides functions for manipulating Windows Explorer and Desktop."""

from win32com.client import Dispatch
from win32com.shell import shell
import win32gui, win32ui, win32con, win32clipboard, winsound
from cythonExtensions.commonUtils.commonUtils import ShellAutomationObjectWrapper as ShellWrapper, PThread, SendToClipboard
import os
import img2pdf

cdef bint _filter(str target, tuple patterns):
    """Returns `True` if the target string ends with any of the patterns, `False` otherwise."""
    
    return target.endswith(patterns)

# Source: https://stackoverflow.com/questions/17984809/how-do-i-create-an-incrementing-filename-in-python
cpdef str GetUniqueName(str directory, str filename="New File", str sequence_pattern=" (%s)", extension=".txt"):
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

cpdef GetActiveExplorer(explorer_windows=None, bint check_desktop=True):
    """Returns the active (focused) explorer/desktop window object."""
    
    cdef bint initializer_called = PThread.CoInitialize()
    
    cdef int fg_hwnd = win32gui.GetForegroundWindow()
    
    # No automation object was passed; create one.
    if not explorer_windows:
        explorer_windows = ShellWrapper.explorer.Windows()
    
    # print(explorer_windows.Item().HWND, fg_hwnd, GetClassName(fg_hwnd), sep=" | ")
    
    # Check if the active window has one of the Desktop window class names. This check is necessary because
    # `GetForegroundWindow()` and `explorer_windows.Item().HWND` might not be the same even when the Desktop is the active window.
    output = None
    
    cdef str curr_className = win32gui.GetClassName(fg_hwnd)
    
    if check_desktop and curr_className in ("WorkerW", "Progman"):
        output = explorer_windows.Item() # Not passing a number to `Item()` returns the desktop window object.
    
    else:
        # Check other explorer windows if any.
        for explorer_window in explorer_windows:
            if explorer_window.HWND == fg_hwnd:
                output = explorer_window
                break
    
    if initializer_called:
        PThread.CoUninitialize()
    
    return output

cpdef str GetExplorerAddress(active_explorer=None):
    """Returns the address of the active explorer window."""
    
    if not active_explorer:
        active_explorer = GetActiveExplorer(ShellWrapper.explorer.Windows(), check_desktop=False)
    
    if not active_explorer:
        return ""
    
    # return win32gui.GetWindowText(active_explorer.HWND) # About 10 times faster but sometimes fails to return the full path.
    return active_explorer.Document.Folder.Self.Path

cpdef list GetSelectedItemsFromActiveExplorer(active_explorer=None, tuple patterns=None):
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
    
    cdef bint initializer_called = PThread.CoInitialize()
    
    if not active_explorer:
        active_explorer = GetActiveExplorer(ShellWrapper.explorer.Windows(), check_desktop=True)
    
    cdef list output = []
    
    if active_explorer:
        for selected_item in active_explorer.Document.SelectedItems():
            if not patterns or _filter(selected_item.Path, patterns):
                output.append(selected_item.Path)
    
    if initializer_called:
        PThread.CoUninitialize()
    
    return output

cpdef list CopySelectedFileNames(active_explorer=None, bint check_desktop=True):
    """
    Description:
        Copies the absolute paths of the selected files from the active explorer/desktop window.
    ---
    Returns:
        `list[str]`: A list containing the paths to the selected items in the active explorer/desktop window.
    """
    
    cdef bint initializer_called = PThread.CoInitialize()
    
    # If no automation object was passed (i.e., `None` was passed), create one.
    if not active_explorer:
        active_explorer = GetActiveExplorer(ShellWrapper.explorer.Windows(), check_desktop)
    
    cdef list selected_files_paths = []
    
    cdef str concatenated_file_paths
    
    if active_explorer:
        selected_files_paths = GetSelectedItemsFromActiveExplorer(active_explorer)
        
        if selected_files_paths:
            concatenated_file_paths = '"' + '" "'.join(selected_files_paths) + '"'
            
            SendToClipboard(concatenated_file_paths, win32clipboard.CF_UNICODETEXT)
        
        winsound.PlaySound(r"SFX\coins-497.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
    
    if initializer_called:
        PThread.CoUninitialize()
    
    return selected_files_paths

# https://learn.microsoft.com/en-us/windows/win32/dlgbox/open-and-save-as-dialog-boxes
cdef list OpenFileDialog(int dialog_type, str default_extension="", str default_filename="", int extra_flags=0,
                         str filter="", bint multiselect=False, str title="File Dialog", str initial_dir=""):
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
    
    cdef int dialog_flags = extra_flags|win32con.OFN_OVERWRITEPROMPT|win32con.OFN_FILEMUSTEXIST # |win32con.OFN_EXPLORER
    
    if multiselect and dialog_type:
        dialog_flags|=win32con.OFN_ALLOWMULTISELECT
    
    ## API: CreateFileDialog(FileSave_0/FileOpen_1, DefaultExtension, InitialFilename, Flags, Filter)
    # o = win32ui.CreateFileDialog(1, ".txt", "default.txt", 0, "Text Files (*.txt)|*.txt|All Files (*.*)|*.*|")
    
    dialog_window = win32ui.CreateFileDialog(dialog_type, default_extension, default_filename, dialog_flags, filter)
    dialog_window.SetOFNTitle(title)
    
    if initial_dir:
        dialog_window.SetOFNInitialDir(initial_dir)
    
    if dialog_window.DoModal()!=win32con.IDOK:
        return []
    
    return dialog_window.GetPathNames()

# https://mail.python.org/pipermail/python-win32/2012-September/012533.html
cpdef void SelectFilesFromDirectory(str directory, list file_names):
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

cpdef int CreateFile(active_explorer=None):
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
    
    cdef bint initializer_called = PThread.CoInitialize()
    
    if not active_explorer:
        active_explorer = GetActiveExplorer(ShellWrapper.explorer.Windows(), check_desktop=True)
    
    cdef int output = 0
    cdef str file_fullpath
    
    if active_explorer:
        file_fullpath = GetUniqueName(directory=GetExplorerAddress(active_explorer))
        
        with open(file_fullpath, 'w') as newfile:
            newfile.writelines('# Created using "File Factory"')
        
        # Selects the file and put it in edit mode: https://learn.microsoft.com/en-us/windows/win32/shell/shellfolderview-selectitem
        active_explorer.Document.SelectItem(file_fullpath, 0x1F) # 0x1F = 31 # 1|4|8|16 = 29
        
        winsound.PlaySound(r"SFX\coins-497.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
        
        output = 1
    
    if initializer_called:
        PThread.CoUninitialize()
    
    return output

cpdef ImagesToPDF(active_explorer=None):
    """Combines the selected images from the active explorer window into a PDF file with an incremental name then select it.
    Please note that the function sorts the file names alphabetically before merging."""
    
    cdef bint initializer_called = PThread.CoInitialize()
    
    if not active_explorer:
        active_explorer = GetActiveExplorer(ShellWrapper.explorer.Windows(), check_desktop=False)
    
    cdef list selected_files_paths
    cdef str directory, file_fullpath
    
    if active_explorer:
        selected_files_paths = sorted(GetSelectedItemsFromActiveExplorer(active_explorer, patterns=('.png', '.jpg', '.jpeg')))
        
        if selected_files_paths:
            winsound.PlaySound(r"SFX\connection-sound.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
            
            directory = os.path.dirname(selected_files_paths[0])
            
            file_fullpath = GetUniqueName(directory, "New PDF", extension=".pdf")
            
            with open(file_fullpath, "wb") as pdf_output_file:
                pdf_output_file.write(img2pdf.convert(selected_files_paths))
                
                if len(selected_files_paths) <= 20:
                    active_explorer.Document.SelectItem(file_fullpath, 1|4|8|16)
                
                winsound.PlaySound(r"SFX\coins-497.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
    
    if initializer_called:
        PThread.CoUninitialize()

cpdef void OfficeFileToPDF(active_explorer=None, str office_application="Powerpoint"):
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
    
    cdef bint initializer_called = PThread.CoInitialize()
    
    if not active_explorer:
        active_explorer = GetActiveExplorer(ShellWrapper.explorer.Windows(), check_desktop=False)
    
    if not active_explorer:
        if initializer_called:
            PThread.CoUninitialize()
        
        return
    
    cdef str office_application_char0 = office_application[0].lower()
    cdef list selected_files_paths = GetSelectedItemsFromActiveExplorer(active_explorer,
                patterns={"p": (".pptx", ".ppt"), "w": (".docx", ".doc")}.get(office_application_char0))
    
    # Check if any file exists to pop it from the selected files list before starting any office application.
    cdef int file_path_counter = 0
    
    cdef int len_selected_files = len(selected_files_paths)
    
    cdef str file_path, new_filepath
    
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
            PThread.CoUninitialize()
        
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
    
    if initializer_called:
        PThread.CoUninitialize()
    
    winsound.PlaySound(r"SFX\coins-497.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)

cpdef void GenericFileConverter(active_explorer=None, tuple patterns=None, convert_func=None, str new_extension=""):
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
    
    cdef bint initializer_called = PThread.CoInitialize()
    
    if not active_explorer:
        active_explorer = GetActiveExplorer(ShellWrapper.explorer.Windows(), check_desktop=False)
    
    if not active_explorer:
        if initializer_called:
            PThread.CoUninitialize()
        
        return
    
    cdef list selected_files_paths = GetSelectedItemsFromActiveExplorer(active_explorer, patterns=patterns)
    cdef str new_filepath
    
    if selected_files_paths:
        winsound.PlaySound(r"SFX\connection-sound.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
        
        for file_path in selected_files_paths:
            new_filepath = os.path.splitext(file_path)[0] + new_extension
            
            if os.path.exists(new_filepath):
                print("Failure, file already exists: %s" % new_filepath)
                continue
            
            convert_func(file_path, new_filepath)
            print(new_filepath)
        
        winsound.PlaySound(r"SFX\coins-497.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
    
    if initializer_called:
        PThread.CoUninitialize()

cpdef void FlattenDirectories(active_explorer=None):
    """Flattens the selected folders from the active explorer window to the explorer current location."""
    
    
    cdef bint initializer_called = PThread.CoInitialize()
    
    # If no automation object was passed (i.e., `None` was passed), create one.
    if not active_explorer:
        active_explorer = GetActiveExplorer(ShellWrapper.explorer.Windows(), False)
    
    if not active_explorer:
        if initializer_called:
            PThread.CoUninitialize()
        
        return
    
    cdef list selected_files_paths = GetSelectedItemsFromActiveExplorer(active_explorer)
    
    if not selected_files_paths:
        if initializer_called:
            PThread.CoUninitialize()
        
        return
    
    cdef str src = GetExplorerAddress(active_explorer)
    
    cdef str dst = os.path.join(src, GetUniqueName(src, "Flattened", extension=""))
    
    os.makedirs(dst, exist_ok=True)
    
    cdef str folder, file_name, target_src, target_dst
    
    # for root_dir, cur_dir, files in os.walk(target_path, topdown=True):
    for folder in os.listdir(src):
        for file_name in os.listdir(os.path.join(src, folder)):
            target_src = os.path.join(src, folder, file_name)
            target_dst = os.path.join(dst, file_name)
            
            if os.path.isfile(target_src):
                os.rename(target_src, target_dst)
    
    if initializer_called:
        PThread.CoUninitialize()
