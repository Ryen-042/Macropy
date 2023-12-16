# cython: language_level = 3str

"""This extension module provides system/script-specific functions."""

import wmi, ctypes, os, sys, psutil, subprocess, importlib
import win32gui, win32api, win32process, win32con, winsound, win32security
from time import sleep
from win11toast import toast

import scriptConfigs as configs
from cythonExtensions.commonUtils.commonUtils import ControllerHouse as ctrlHouse, PThread, Management as mgmt
from cythonExtensions.windowHelper import windowHelper as winHelper

cpdef void reloadConfigs():
    """Re-imports the `scriptConfigs` module and reloads the defined configurations."""
    
    importlib.reload(configs)
    
    ctrlHouse.abbreviations = configs.ABBREVIATIONS
    ctrlHouse.non_prefixed_abbreviations = configs.NON_PERFIXED_ABBREVIATIONS
    ctrlHouse.locations = configs.LOCATIONS
    mgmt.silent = configs.SUPPRESS_TERMINAL_OUTPUT

@PThread.throttle(10)
def terminateScript(graceful=False) -> None:
    """
    Description:
        Terminates the running main script process.
    ---
    Parameters:
        `graceful -> bool`:
            `True` : Does not terminate the script. Only sets the global Event variable `Management.terminateEvent`.
            `False`: Forcefully terminating the script.
    """
    
    if graceful:
        print('Exitting...')
        winsound.PlaySound(r"SFX\crack_the_whip.wav", winsound.SND_FILENAME)
        
        # Set the global Event variable `Management.terminateEvent` to signal the other threads to terminate.
        mgmt.terminateEvent.set()
        
        # Post the quit message to the thread's message queue so that the `GetMessage` function returns `False` and the thread terminates.
        ctypes.windll.user32.PostThreadMessageW(PThread.mainThreadId, win32con.WM_QUIT, 0, 0)
    
    elif not len(sys.argv) > 1 or sys.argv[1] not in ("-p", "--profile", "--prof"): # else:
        print("Forcefully terminating the script...")
        
        # os._exit() vs sys.exit():
        # - `os._exit()` exits script immediately, without calling cleanup handlers, flushing stdio buffers, etc.
        # - `sys.exit()` raises `SystemExit` exception, this causes the interpreter to exit.
        os._exit(1)


cpdef bint isProcessElevated(int hwnd=0):
    """
    Description:
        - Checks if the window with the specified handle has elevated privileges.
            - If a handle is not specified, the foreground window is checked.
            - If the specified handle is `-1`, the current python process is checked.
    
    ---
    Return:
        - `True`:  `The specified process has elevated privileges`.
        - `False`: `The specified process dose not have elevated privileges`.
    """
    
    if hwnd == -1:
        return ctypes.windll.shell32.IsUserAnAdmin()
    
    hwnd = hwnd or win32gui.GetForegroundWindow()
    
    try:
        if psutil.Process(win32process.GetWindowThreadProcessId(hwnd)[-1]).cwd() is None:
            return True
    
    except psutil.AccessDenied:
        return True
    
    except (psutil.NoSuchProcess, ValueError):
        if hwnd:
            print(f"No such process with the specified handle: {hwnd}")
    
    return False


def startWithElevatedPrivileges(terminate=True, cmder=False, cmdShow=win32con.SW_SHOWNORMAL) -> int: # win32con.SW_FORCEMINIMIZE
    """Starts another instance of the main python process with elevated privileges."""
    
    if isProcessElevated(-1):
        print("Hotkey ignored. The script already has elevated privileges.")
        
        return 0
    
    if terminate:
        terminateScript(graceful=True)
    
    if cmder:
        # Docs: Run with elevated privileges and disable ‘Press Enter or Esc to close console’: https://conemu.github.io/en/NewConsole.html
        # Docs: Run a command with elevation in cmder & configure where the command should started (in a split or new tab): https://conemu.github.io/en/csudo.html
        return os.system(f'''start "" "{os.path.join(os.getenv("CmderDir"), "Cmder.exe")}" /x "-run python \\"{os.path.join(configs.MAIN_MODULE_LOCATION, "__main__.py")}\\" -cur_console:a:n"''')
        
        # Useful: https://stackoverflow.com/questions/130763/request-uac-elevation-from-within-a-python-script
        # Useful: https://stackoverflow.com/questions/19672352/how-to-run-script-with-elevated-privilege-on-windows
        # MsDocs: https://learn.microsoft.com/en-us/windows/win32/api/shellapi/nf-shellapi-shellexecutew
        # return ctypes.windll.shell32.ShellExecuteW(None, None, os.path.join(os.getenv("CmderDir"), "Cmder.exe"), f'/x "-run python \\"{os.path.join(configs.MAIN_MODULE_LOCATION, "__main__.py")}\\" -cur_console:a:n"', None, win32con.SW_SHOWNORMAL)
        
        # This does not use UAC so you would have to enter the password for the admin account.
        # return os.system(fr'''runas /user:Administrator "{os.path.join(os.getenv("CmderDir"), "Cmder.exe")} /x \"-run csudo python \\\"{__file__}\\\"\""''')
    
    return ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, subprocess.list2cmdline(sys.argv), None, cmdShow)


def scheduleElevatedProcessChecker(delay=10.0) -> None:
    """Reprots each `delay` time if the active process window is elevated while the current python process is not elevated."""
    
    while not mgmt.terminateEvent.wait(delay):
        if isProcessElevated():
            if isProcessElevated(-1):
                print("The script has elevated privileges. No need for further checks.")
                return
            
            winsound.PlaySound(r"C:\Windows\Media\Windows Exclamation.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
            
            windowTitle = win32gui.GetWindowText(win32gui.GetForegroundWindow())
            if windowTitle.endswith("(Not Responding)"):
                print(f"Attention: the active process '{windowTitle}' is not responding and no keyboard events can be received.")
            
            else:
                print(f"Attention: the active process '{windowTitle}' has elevated privileges and no keyboard events can be received.")


def displayCPU_Usage() -> None:
    """Prints the current CPU and Memort usage to the console."""
    cdef double cpu_usage, scale, mem_usage
    cdef int columns, max_width, cpu_filled, cpu_remaining, mem_filled, mem_remaining
    
    # Define visual characters for progress bars.
    char_empty    = "░" # ▄ ░ ▒ ▓
    char_fill     = "█"
    
    # Get the size of the terminal to scale the progress bars.
    scale         = 0.30
    columns       = os.get_terminal_size().columns
    max_width     = int(columns * scale)
    
    # Creating a CPU bar.
    cpu_usage     = psutil.cpu_percent()
    cpu_filled    = int(cpu_usage * max_width / 100)
    cpu_remaining = max_width - cpu_filled
    cpu_bar       = char_fill * cpu_filled + char_empty * cpu_remaining
    
    # Creating a memory bar.
    mem_usage     = psutil.virtual_memory().percent
    mem_filled    = int(mem_usage * max_width / 100)
    mem_remaining = max_width - mem_filled
    mem_bar       = char_fill * mem_filled + char_empty * mem_remaining
    
    # Printing the bars.
    print(f"\033[F\033[KCPU Usage: | {cpu_bar} | ({cpu_usage})\nMem Usage: | {mem_bar} | ({mem_usage})", end="")


def enableDPI_Awareness() -> int:
    """Enables `DPI Awareness` for the current thread to allow for accurate dimensions reporting."""
    
    # Creator note: behavior on later OSes is undefined, although when I run it on my Windows 10 machine, it seems to work with effects identical to SetProcessDpiAwareness(1)
    # Source: https://stackoverflow.com/questions/44398075/can-dpi-scaling-be-enabled-disabled-programmatically-on-a-per-session-basis/44422362#44422362
    # Source: https://stackoverflow.com/questions/32541475/win32api-is-not-giving-the-correct-coordinates-with-getcursorpos-in-python
    
    # Query DPI Awareness (Windows 10 and 8)
    awareness = ctypes.c_int()
    cdef int errorCode = ctypes.windll.shcore.GetProcessDpiAwareness(0, ctypes.byref(awareness))
    print("Current awareness value:", awareness.value, end=" -> ")
    
    if awareness.value != 0:
        return errorCode
    
    # Set DPI Awareness  (Windows 10 and 8). The argument is the awareness level. Check the link below for all valid values.
    # https://learn.microsoft.com/en-us/windows/win32/api/shellscalingapi/ne-shellscalingapi-process_dpi_awareness
    # Further reading: https://learn.microsoft.com/en-us/windows/win32/hidpi/high-dpi-desktop-application-development-on-windows
    errorCode = ctypes.windll.shcore.SetProcessDpiAwareness(2) # 1 seems to work fine
    
    # Printing the updated value.
    ctypes.windll.shcore.GetProcessDpiAwareness(0, ctypes.byref(awareness))
    print(awareness.value)
    
    # Set DPI Awareness (Windows 7 and Vista). Also seems to work fine for windows 10.
    # success = ctypes.windll.user32.SetProcessDPIAware()
    return errorCode


def sendScriptWorkingNotification(near_module=True) -> None:
    """Sends a notification to the user that the script is working."""
    if near_module:
        directory = os.path.dirname(os.path.abspath(__file__))
    else:
        directory = os.getcwd()
    
    cdef tuple buttons = (
        {'activationType': 'protocol', 'arguments': "0:", 'content': 'Exit Script', "hint-buttonStyle": "Critical"},
        {'activationType': 'protocol', 'arguments': directory, 'content': 'Open Script Folder'},
        # {'activationType': 'protocol', 'arguments': "1:", 'content': 'Nice Work', "hint-buttonStyle": "Success"},
        {'activationType': 'protocol', 'arguments': "1:", 'content': 'Reload Configs', "hint-buttonStyle": "Success"},
    )
    
    # notify() doesn't work properly here. Use toast() inside a thread instead.
    toast(
        'Script is Running.', 'The script is running in the background.', buttons=buttons,
        on_click = lambda args: (args["arguments"][0] == "0" and terminateScript(True)) or (args["arguments"][0] == "1" and reloadConfigs()),
        icon  = {"src": os.path.join(directory, "Images", "static", "keyboard.png"), 'placement': 'appLogoOverride'},
        image = {'src': os.path.join(directory, "Images", "static", "keyboard (0.5).png"), 'placement': 'hero'},
        # audio = os.path.join(configs.MAIN_MODULE_LOCATION, "SFX", "bonk sound.mp3"), # {'silent': 'true'}
    )


@PThread.throttle(0.05)
def changeBrightness(opcode=1, increment=5) -> None:
    """Increments (`opcode=any non-zero value`) or decrements (`opcode=0`) the screen brightness by an (`increment`) percent."""
    
    PThread.coInitialize()
    
    # Connectting to WMI.
    c = wmi.WMI(namespace="wmi")
    
    # Getting the current brightness value.
    current_brightness = c.WmiMonitorBrightness()[0].CurrentBrightness
    
    # Adding the increment value to the current brightness value.
    brightness = current_brightness + (-increment if not opcode else increment)
    
    # Clipping the brightness level to a valid range from 0 to 100.
    brightness = min(max(brightness, 0), 100)
    
    # Setting the screen brightness to the new value.
    c.WmiMonitorBrightnessMethods()[0].WmiSetBrightness(Brightness=brightness, Timeout=0)
    
    print(f"Current & New Brightness: {current_brightness} -> {brightness}")
    
    PThread.coUninitialize()


def screenOff() -> None:
    """Turns off the screen."""
    win32gui.SendMessage(win32con.HWND_BROADCAST, win32con.WM_SYSCOMMAND, win32con.SC_MONITORPOWER, 2)


def flashScreen(delay=0.15) -> None:
    """Inverts the color of the screen for the specified number of seconds."""
    cdef int x, y
    
    hdc = win32gui.GetDC(0) # Get the screen as a Device Context object
    x, y = win32api.GetSystemMetrics(0), win32api.GetSystemMetrics(1) # Retrieve monitor size, e.g., (1920, 1080).
    win32gui.PatBlt(hdc, 0, 0, x, y, win32con.PATINVERT) # Invert the device context.
    sleep(delay)
    win32gui.PatBlt(hdc, 0, 0, x, y, win32con.PATINVERT) # Invert back to normal.
    win32gui.DeleteDC(hdc) # Clean up memory.


@PThread.throttle(15)
def goToSleep() -> None:
    """Puts the device to sleep."""
    
    # Source: https://learn.microsoft.com/en-us/sysinternals/downloads/psshutdown
    # For a list of available options, run (make sure to write correct path to the executable): c:\Utilities\PSTools\psshutdown
    
    winsound.PlaySound(r"C:\Windows\Media\Windows Logoff Sound.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
    print("Putting the device to sleep...")
    
    os.system(r'"C:\Program Files\Utilities\PSTools\psshutdown.exe" -d -t 0 -nobanner')
    print("Device is now active.")


@PThread.throttle(0.1)
def shutdown(request_confirmation=False) -> None:
    """Shuts down the computer."""
    
    if request_confirmation:
        if winHelper.showMessageBox("Do you really want to shutdown?", "Confirm Shutdown", 2, win32con.MB_ICONQUESTION) != 6:
            return
    
    # Source: https://stackoverflow.com/questions/34039845/how-to-shutdown-a-computer-using-python/62595482#62595482
    win32security.AdjustTokenPrivileges(win32security.OpenProcessToken(win32api.GetCurrentProcess(),
                                                                       win32security.TOKEN_ADJUST_PRIVILEGES | win32security.TOKEN_QUERY),
                                        False, [(win32security.LookupPrivilegeValue(None, win32security.SE_SHUTDOWN_NAME),
                                                 win32security.SE_PRIVILEGE_ENABLED)])
    
    # API: https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-exitwindowsex
    # win32api.ExitWindowsEx(win32con.EWX_POWEROFF)
    
    cdef int EWX_HYBRID_SHUTDOWN = 0x00400000
    win32api.ExitWindowsEx(win32con.EWX_SHUTDOWN | EWX_HYBRID_SHUTDOWN)
    
    # Make sure that the script is terminating while the system is shutting down.
    terminateScript(graceful=False)


def getProcessFileAddress(hwnd: int) -> str:
    """Given a window handle, returns its process file address."""
    
    process_handle = 0
    try:
        process_handle = win32api.OpenProcess(win32con.PROCESS_ALL_ACCESS, False, win32process.GetWindowThreadProcessId(hwnd)[1])
        
        process_location = win32process.GetModuleFileNameEx(process_handle, 0)
    
    except Exception as e:
        print(f"Error accessing process with hwnd={hwnd}: {e}")
        
        process_location = ""
    
    if process_handle:
        win32api.CloseHandle(process_handle)
    
    return process_location


# Source: https://stackoverflow.com/questions/38628332/how-to-get-the-process-id-of-not-responding-foreground-app
# Useful: https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-ishungappwindow
#       : https://www.pinvoke.net/default.aspx/ntdll.NtSuspendProcess
#       : https://www.pinvoke.net/default.aspx/ntdll/NtResumeProcess.html
#       : https://undoc.airesoft.co.uk/user32.dll/HungWindowFromGhostWindow.php
#       : https://undoc.airesoft.co.uk/user32.dll/GhostWindowFromHungWindow.php
def suspendProcess(hwnd=0) -> int:
    """Suspends a process given its window handle. Uses the handle of the active window if no handle is passed."""
    
    if not hwnd:
        hwnd = win32gui.GetForegroundWindow()
    
    _thread_id, process_id = win32process.GetWindowThreadProcessId(hwnd)
    if isProcessSuspended(process_id):
        try:
            print(f"The '{psutil.Process(process_id).name()}' process with hwnd={hwnd} and pid={process_id} is already suspended.")
        except Exception as ex:
            print(f"The process with hwnd={hwnd} and pid={process_id} is already suspended.")
        
        return 0
    
    process_handle = 0
    try:
        process_handle = ctypes.windll.kernel32.OpenProcess(win32con.PROCESS_SUSPEND_RESUME, False, process_id)
        
        ctypes.windll.ntdll.NtSuspendProcess(process_handle)
        
        try:
            process_name = psutil.Process(process_id).name()
            
            winsound.PlaySound(r"SFX\no-trespassing-368.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
            
            print(f"Successfully suspended the '{process_name}' process with hwnd={hwnd} and pid={process_id}.")
        
        except Exception as ex:
            print(f"Successfully suspended the process with hwnd={hwnd}, pid={process_id}")
        
        output = 1
    
    except Exception as ex:
        print(f"{ex}: Could not suspend the process with hwnd={hwnd} and pid={process_id}. Make sure you have the necessary permissions.")
        
        output = 0
    
    if process_handle:
        ctypes.windll.kernel32.CloseHandle(process_handle)
    
    return output


def resumeProcess(hwnd=0) -> int:
    """Resumes a suspended process given its window handle. Uses the handle of the active window if no handle is passed."""
    
    if not hwnd:
        hwnd = win32gui.GetForegroundWindow()
    
    if ctypes.windll.user32.IsHungAppWindow(hwnd):
        hwnd = ctypes.windll.user32.HungWindowFromGhostWindow(hwnd)
    
    _thread_id, process_id = win32process.GetWindowThreadProcessId(hwnd)
    
    if not isProcessSuspended(process_id):
        try:
            print(f"The '{psutil.Process(process_id).name()}' process with hwnd={hwnd} and pid={process_id} is not suspended.")
        except Exception as ex:
            print(f"The process with hwnd={hwnd} and pid={process_id} is not suspended.")
        
        return 0
    
    process_handle = 0
    try:
        process_handle = ctypes.windll.kernel32.OpenProcess(win32con.PROCESS_SUSPEND_RESUME, False, process_id)
        
        ctypes.windll.ntdll.NtResumeProcess(process_handle)
        
        try:
            process_name = psutil.Process(process_id).name()
            
            winsound.PlaySound(r"SFX\pedantic-490.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
            
            print(f"Successfully resumed the '{process_name}' process with hwnd={hwnd} and pid={process_id}.")
        
        except Exception as ex:
            print(f"Successfully resumed the process with hwnd={hwnd}, pid={process_id}")
        
        output = 1
    
    except Exception as ex:
        print(f"{ex}: Could not resume the process with hwnd={hwnd} and pid={process_id}. Make sure you have the necessary permissions.")
        
        output = 0
    
    if process_handle:
        ctypes.windll.kernel32.CloseHandle(process_handle)
    
    return output


def getHungwindowHandle(hwnd=0) -> int:
    """Returns the actual hwnd of a hung window given its ghost window handle. Uses the handle of the active window if no handle is passed."""
    
    if not hwnd:
        hwnd = win32gui.GetForegroundWindow()
    
    if ctypes.windll.user32.IsHungAppWindow(hwnd):
        real_hwnd = ctypes.windll.user32.HungWindowFromGhostWindow(hwnd)
        
        return real_hwnd
        
        # To get the reverse mapping:
        # ghost_hwnd = ctypes.windll.user32.GhostWindowFromHungWindow(real_hwnd)
    
    return 0


cdef bint isProcessSuspended(int pid):
    """Returns whether a process is suspended or not."""
    
    try:
        process = psutil.Process(pid)
        return process.status() == psutil.STATUS_STOPPED
    
    except psutil.NoSuchProcess:
        return False
