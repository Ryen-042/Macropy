"""This module provides functions for manipulating Windows Explorer and Desktop."""

# Postponed evaluation of annotations, for more info: https://docs.python.org/3/whatsnew/3.7.html#pep-563-postponed-evaluation-of-annotations
from __future__ import annotations
from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from typing import Callable

from win32com.client import Dispatch, CDispatch
from win32com.shell import shell
import win32gui, win32ui, win32con, win32clipboard, winsound
from common import ShellAutomationObjectWrapper as ShellWrapper, PThread, SendToClipboard
import os, subprocess


# Source: https://stackoverflow.com/questions/17984809/how-do-i-create-an-incrementing-filename-in-python
def GetUniqueName(directory: str, filename="New File", sequence_pattern=" (%s)", extension=".txt") -> str:
    """
    Description:
        Finds the next unused incremental filename in the specified directory.
    ---
    Usage:
        >>> GetUniqueName(directory, "New File", " (%s)", ".txt")
    ---
    Returns:
        `str`: A string in this format `f"{directory}\{filename}{pattern}{extension}"`.
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
    file_counter = 1
    
    # First do an exponential search
    while os.path.exists(filename % file_counter):
        file_counter = file_counter * 2
    
    # Result lies somewhere in the interval (i/2..i]
    # We call this interval (a..b] and narrow it down until a + 1 = b
    a, b = (file_counter // 2, file_counter)
    while a + 1 < b:
        c = (a + b) // 2 # interval midpoint
        a, b = (c, b) if os.path.exists(filename % c) else (a, c)
    
    return filename % b

def GetExplorerAddress(active_explorer: CDispatch) -> str:
    """Returns the address of the active explorer window."""
    
    if not active_explorer:
        active_explorer = GetActiveExplorer(ShellWrapper.explorer.Windows(), check_desktop=False)
    
    return win32gui.GetWindowText(active_explorer.HWND) # About 10 times faster.
    # return active_explorer.Document.Folder.Self.Path


def GetActiveExplorer(explorer_windows: CDispatch, check_desktop=True) -> CDispatch | str:
    """
    Description:
        Returns the active (focused) explorer/desktop window object.
    ---
    Returns:
        - `CDispatch`: The active explorer/desktop window object.
        - `str`: The class name of the active window if it was not an explorer window.
    """
    
    initializer_called = PThread.CoInitialize()
    
    fg_hwnd = win32gui.GetForegroundWindow()
    
    # No automation object was passed; create one.
    if not explorer_windows:
        explorer_windows = ShellWrapper.explorer.Windows()
    
    # Check if the active window has one of the Desktop window class names. This check is necessary because
    # `GetForegroundWindow()` and `explorer_windows.Item().HWND` might not be the same even when the Desktop is the active window.
    output = None
    curr_className = win32gui.GetClassName(fg_hwnd)
    if check_desktop and curr_className in ("WorkerW", "Progman"):
        output = explorer_windows.Item() # Not passing a number to `Item()` returns the desktop window object.
    
    else:
        # Check other explorer windows if any.
        for explorer_window in explorer_windows:
            if explorer_window.HWND == fg_hwnd:
                output = explorer_window
    
    PThread.CoUninitialize(initializer_called)
    return output or curr_className

def GetSelectedItemsFromActiveExplorer(active_explorer: CDispatch, paths_only=False, filter: Callable[[str], bool] = None) -> list[str]:
    """Returns a list containing the paths to the selected items in the active explorer/desktop window."""
    
    initializer_called = PThread.CoInitialize()
    if not active_explorer:
        active_explorer = GetActiveExplorer(ShellWrapper.explorer.Windows(), check_desktop=True)
    
    selected_items = active_explorer.Document.SelectedItems()
    output = []
    if paths_only:
        for selected_item in selected_items:
            if not callable(filter) or filter(selected_item.Path):
                output.append(selected_item.Path)
    
    PThread.CoUninitialize(initializer_called)
    return (not paths_only and selected_items) or output

def CopySelectedFileNames(check_desktop=True) -> list[str]:
    """
    Description:
        Copies the absolute paths of the selected files from the active explorer/desktop window.
    ---
    Returns:
        `list[str]`: A list containing the paths to the selected items in the active explorer/desktop window.
    """
    
    initializer_called = PThread.CoInitialize()
    
    # If no automation object was passed (i.e., `None` was passed), create one.
    active_explorer = GetActiveExplorer(ShellWrapper.explorer.Windows(), check_desktop)
    
    selected_files_paths = []
    if isinstance(active_explorer, CDispatch):
        selected_files_paths = GetSelectedItemsFromActiveExplorer(active_explorer, paths_only=True)
        if selected_files_paths:
            concatenated_file_paths = '"' + '" "'.join(selected_files_paths) + '"'
            SendToClipboard(concatenated_file_paths, win32clipboard.CF_UNICODETEXT)
        winsound.PlaySound(r"SFX\coins-497.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
    
    PThread.CoUninitialize(initializer_called)
    return selected_files_paths

# https://learn.microsoft.com/en-us/windows/win32/dlgbox/open-and-save-as-dialog-boxes
def OpenFileDialog(dialog_type, default_extension="", default_filename="", extra_flags: int=0,
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
    
    dialog_flags = extra_flags|win32con.OFN_OVERWRITEPROMPT|win32con.OFN_FILEMUSTEXIST #|win32con.OFN_EXPLORER
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
def SelectFilesFromDirectory(directory, file_names):
    """Given an absolute directory path and the names of its items (names relative to the path), if an explorer window with the specified directory is present, use it, otherwise open a new one, then select all the items specified."""
    
    folder_pidl = shell.SHILCreateFromPath(directory, 0)[0]
    
    desktop = shell.SHGetDesktopFolder()
    shell_folder = desktop.BindToObject(folder_pidl, None, shell.IID_IShellFolder)
    name_to_item_mapping = dict([(desktop.GetDisplayNameOf(item, 0), item) for item in shell_folder])
    
    to_show = []
    for file in file_names:
        if not name_to_item_mapping.get(file):
            raise Exception('File: "%s" not found in "%s"' % (file, directory))
        to_show.append(name_to_item_mapping[file])
    shell.SHOpenFolderAndSelectItems(folder_pidl, to_show, 0)

def CreateFile(active_explorer: CDispatch) -> int:
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
    
    initializer_called = PThread.CoInitialize()
    
    if not isinstance(active_explorer, CDispatch):
        active_explorer = GetActiveExplorer(ShellWrapper.explorer.Windows(), check_desktop=True)
    
    output = 0
    if isinstance(active_explorer, CDispatch):
        file_fullpath = GetUniqueName(directory=GetExplorerAddress(active_explorer))
        
        ## Creates a file and returns its file object, which you can use functions like `WriteLine()` on.
        # filesystem = win32com.client.Dispatch("Scripting.FileSystemObject")
        # newfile = filesystem.CreateTextFile(os.path.join(filepath, filename), False) # <class 'win32com.client.CDispatch'>
        # newfile.WriteLine('Created using "File Factory"') # Inserts a newline automatically
        # newfile.Write(os.path.join(filepath, filename)) # No newline inserted
        # newfile.Close()
        
        with open(file_fullpath, 'w') as newfile:
            newfile.writelines('# Created using "File Factory"')
        
        # Selects the file and put it in edit mode: https://learn.microsoft.com/en-us/windows/win32/shell/shellfolderview-selectitem
        active_explorer.Document.SelectItem(file_fullpath, 0x1F) # 0x1F = 31 # 1|4|8|16 = 29
        winsound.PlaySound(r"SFX\coins-497.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
        output = 1
    
    PThread.CoUninitialize(initializer_called)
    return output

def ImagesToPDF(active_explorer: CDispatch) -> None:
    """Combines the selected images from the active explorer window into a PDF file with an incremental name then select it."""
    
    import img2pdf
    
    initializer_called = PThread.CoInitialize()
    if not isinstance(active_explorer, CDispatch):
        active_explorer = GetActiveExplorer(ShellWrapper.explorer.Windows(), check_desktop=False)
    
    if isinstance(active_explorer, CDispatch):
        selected_files_paths = sorted(GetSelectedItemsFromActiveExplorer(active_explorer, paths_only=True, filter=lambda f: f.endswith(('.png', '.jpg', '.jpeg'))))
        if selected_files_paths:
            winsound.PlaySound(r"SFX\connection-sound.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
            directory = os.path.split(selected_files_paths[0])[0]
            file_fullpath = GetUniqueName(directory, "New PDF", extension=".pdf")
            with open(file_fullpath,"wb") as pdf_output_file:
                pdf_output_file.write(img2pdf.convert(selected_files_paths))
                if len(selected_files_paths) <= 20:
                    active_explorer.Document.SelectItem(file_fullpath, 1|4|8|16)
                winsound.PlaySound(r"SFX\coins-497.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
    PThread.CoUninitialize(initializer_called)

def CtrlShift_M():
    initializer_called = PThread.CoInitialize()
    
    active_explorer = GetActiveExplorer(ShellWrapper.explorer.Windows(), check_desktop=True)
    if isinstance(active_explorer, CDispatch):
        CreateFile(active_explorer)
    
    # else:
        # Define any other action here
    
    PThread.CoUninitialize(initializer_called)

def OfficeFileToPDF(office_application="Powerpoint") -> None:
    """
    Description:
        Converts the selected files from the active explorer window that are associated with the specified office application into a PDF format.
    ---
    Parameters:
        `office_application` (`str`)
            Only the associated files types with this application will be converted, any other types will be ignored.
            
            Ex => `"Powerpoint": (".pptx", ".ppt")` | `"Word": (".docx", ".doc")`
    """
    
    initializer_called = PThread.CoInitialize()
    
    active_explorer = GetActiveExplorer(ShellWrapper.explorer.Windows(), check_desktop=False)
    
    if not isinstance(active_explorer, CDispatch):
        PThread.CoUninitialize(initializer_called)
        return
    
    office_application_char0 = office_application[0].lower()
    selected_files_paths = GetSelectedItemsFromActiveExplorer(active_explorer, paths_only=True,
                filter=lambda f: f.endswith({"p": (".pptx", ".ppt"), "w": (".docx", ".doc")}.get(office_application_char0)))
    
    # Check if any file exists to pop it from the selected files list before starting any office application.
    file_path_counter = 0
    len_selected_files = len(selected_files_paths)
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
        PThread.CoUninitialize(initializer_called)
        return
    
    winsound.PlaySound(r"SFX\connection-sound.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
    office_dispatch = Dispatch(f"{office_application}.Application")
    office_dispatch.Visible = 1
    
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
    PThread.CoUninitialize(initializer_called)
    winsound.PlaySound(r"SFX\coins-497.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)

def CtrlShift_P(explorer: CDispatch=None):
    initializer_called = PThread.CoInitialize()
    active_explorer = GetActiveExplorer((explorer or ShellWrapper.explorer).Windows(), check_desktop=False)
    
    if isinstance(active_explorer, CDispatch):
        ImagesToPDF(active_explorer)
    
    # else:
        # Define any other action here.
    
    PThread.CoUninitialize(initializer_called)

def MP3ToWAV():
    """Converts the selected audio files from the active explorer window into `.wav` files."""
    
    initializer_called = PThread.CoInitialize()
    active_explorer = GetActiveExplorer(ShellWrapper.explorer.Windows(), check_desktop=False)
    
    if not isinstance(active_explorer, CDispatch):
        PThread.CoUninitialize(initializer_called)
        return
    
    selected_files_paths = GetSelectedItemsFromActiveExplorer(active_explorer, paths_only=True,
                                                                        filter=lambda f: f.endswith(".mp3"))
    
    if selected_files_paths:
        winsound.PlaySound(r"SFX\connection-sound.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
        for file_path in selected_files_paths:
            new_filepath = os.path.splitext(file_path)[0] + ".wav"
            if os.path.exists(new_filepath):
                print("Failure, file already exists: %s" % new_filepath)
                
                continue
            
            # Run a command and wait for it to complete.
            subprocess.call(["ffmpeg", "-loglevel", "error", "-hide_banner", "-nostats",'-i', file_path, new_filepath])
            
            print("Success: %s" % new_filepath)
        
        winsound.PlaySound(r"SFX\coins-497.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
    
    PThread.CoUninitialize(initializer_called)


def FlattenDirectories(files_only=False):
    """Flattens the selected folders from the active explorer window to the explorer current location."""
    
    initializer_called = PThread.CoInitialize()
    
    # If no automation object was passed (i.e., `None` was passed), create one.
    active_explorer = GetActiveExplorer(ShellWrapper.explorer.Windows(), False)
    
    if not isinstance(active_explorer, CDispatch):
        PThread.CoUninitialize(initializer_called)
        return
    
    selected_files_paths = GetSelectedItemsFromActiveExplorer(active_explorer, paths_only=True)
    if not selected_files_paths:
        PThread.CoUninitialize(initializer_called)
        return
    
    src = GetExplorerAddress(active_explorer)
    dst = os.path.join(src, GetUniqueName(src, "Flattened", extension=""))
    os.makedirs(dst, exist_ok=True)
    
    # for root_dir, cur_dir, files in os.walk(target_path, topdown=True):
    for folder in os.listdir(src):
        for file_name in os.listdir(os.path.join(src, folder)):
            target_src = os.path.join(src, folder, file_name)
            target_dst = os.path.join(dst, file_name)
            
            if os.path.isfile(target_src):
                os.rename(target_src, target_dst)
    
    PThread.CoUninitialize(initializer_called)

if __name__ == "__main__":
    # MP3ToWAV(None)
    # FlattenDirectories(None)
    pass
