# cython: embedsignature = True
# cython: language_level = 3str

"""This extension module provides system/script-specific functions."""

import win32gui, win32api, win32process, win32con, winsound, win32security
import wmi, ctypes, os, sys, psutil
from win11toast import toast
from time import sleep
from cythonExtensions.commonUtils.commonUtils import PThread, Management as mgmt
from cythonExtensions.windowHelper import windowHelper as winHelper

@PThread.Throttle(10)
def TerminateScript(graceful=False) -> None:
    """
    Description:
        Terminates the running main script process.
    ---
    Parameters:
        `graceful -> bool`:
            `True` : Does not terminate the script. Only sets the global Event variable `DebuggingHouse.terminateEvent`.
            `False`: Forcefully terminating the script.
    """
    
    if graceful:
        print('Exitting...')
        winsound.PlaySound(r"SFX\crack_the_whip.wav", winsound.SND_FILENAME)
        
        # Set the global Event variable `DebuggingHouse.terminateEvent` to signal the other threads to terminate.
        mgmt.terminateEvent.set()
        
        # Post the quit message to the thread's message queue so that the `GetMessage` function returns `False` and the thread terminates.
        ctypes.windll.user32.PostThreadMessageW(PThread.mainThreadId, win32con.WM_QUIT, 0, 0)
    
    elif not len(sys.argv) > 1 or sys.argv[1] not in ("-p", "--profile", "--prof"): # else:
        print("Forcefully terminating the script...")
        
        # os._exit() vs sys.exit():
        # - `os._exit()` exits script immediately, without calling cleanup handlers, flushing stdio buffers, etc.
        # - `sys.exit()` raises `SystemExit` exception, this causes the interpreter to exit.
        os._exit(1)


cpdef bint IsProcessElevated(int hwnd=0):
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
        print(f"No such process with the specified handle: {hwnd}")
    
    return False

cpdef void RequestElevation():
    """Restarts the current python process with elevated privileges."""
    
    # Source  : https://stackoverflow.com/questions/130763/request-uac-elevation-from-within-a-python-script
    # See also: https://stackoverflow.com/questions/19672352/how-to-run-script-with-elevated-privilege-on-windows
    
    # Re-run the program with admin rights
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)

cpdef void ScheduleElevatedProcessChecker(float delay=10):
    """Reprots each `delay` time if the active process window is elevated while the current python process is not elevated."""
    
    while not mgmt.terminateEvent.wait(delay):
        if IsProcessElevated():
            if IsProcessElevated(-1):
                print("The script has elevated privileges. No need for further checks.")
                return
            
            winsound.PlaySound(r"C:\Windows\Media\Windows Exclamation.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
            print(f"Attention: the active process '{win32gui.GetWindowText(win32gui.GetForegroundWindow())}' has elevated privileges and no keyboard events can be received.")
        
        

cpdef void DisplayCPUsage():
    """Prints the current CPU and Memort usage to the console."""
    cdef str char_empty, char_fill, cpu_bar, mem_bar
    cdef float cpu_usage, scale, mem_usage
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

cpdef int EnableDPI_Awareness():
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


def SendScriptWorkingNotification(near_module=True) -> None:
    """Sends a notification to the user that the script is working."""
    if near_module:
        directory = os.path.dirname(os.path.abspath(__file__))
    else:
        directory = os.getcwd()
    
    cdef tuple buttons = ({'activationType': 'protocol', 'arguments': "0:", 'content': 'Exit Script', "hint-buttonStyle": "Critical"},
                          {'activationType': 'protocol', 'arguments': directory, 'content': 'Open Script Folder'},
                          {'activationType': 'protocol', 'arguments': "1:", 'content': 'Nice Work', "hint-buttonStyle": "Success"})
    
    # notify() doesn't work properly here. Use toast() inside a thread instead.
    toast('Script is Running.', 'The script is running in the background.', buttons=buttons,
          on_click = lambda args: args["arguments"][0] == "0" and TerminateScript(True),
          icon  = {"src": os.path.join(directory, "Images", "static", "keyboard.png"), 'placement': 'appLogoOverride'},
          image = {'src': os.path.join(directory, "Images", "static", "keyboard (0.5).png"), 'placement': 'hero'},
          audio = {'silent': 'true'})


@PThread.Throttle(0.05)
def ChangeBrightness(opcode=1, increment=5) -> None:
    """Increments (`opcode=any non-zero value`) or decrements (`opcode=0`) the screen brightness by an (`increment`) percent."""
    
    PThread.CoInitialize()
    
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
    
    PThread.CoUninitialize()


cpdef void ScreenOff():
    """Turns off the screen."""
    win32gui.SendMessage(win32con.HWND_BROADCAST, win32con.WM_SYSCOMMAND, win32con.SC_MONITORPOWER, 2)


cpdef void FlashScreen(float delay=0.15):
    """Inverts the color of the screen for the specified number of seconds."""
    cdef int x, y
    
    hdc = win32gui.GetDC(0) # Get the screen as a Device Context object
    x, y = win32api.GetSystemMetrics(0), win32api.GetSystemMetrics(1) # Retrieve monitor size, e.g., (1920, 1080).
    win32gui.PatBlt(hdc, 0, 0, x, y, win32con.PATINVERT) # Invert the device context.
    sleep(delay)
    win32gui.PatBlt(hdc, 0, 0, x, y, win32con.PATINVERT) # Invert back to normal.
    win32gui.DeleteDC(hdc) # Clean up memory.


@PThread.Throttle(15)
def GoToSleep() -> None:
    """Puts the device to sleep."""
    
    # Source: https://learn.microsoft.com/en-us/sysinternals/downloads/psshutdown
    # For a list of available options, run (make sure to write correct path to the executable): c:\Utilities\PSTools\psshutdown
    
    winsound.PlaySound(r"C:\Windows\Media\Windows Logoff Sound.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
    print("Putting the device to sleep...")
    
    os.system(r'"C:\Utilities\PSTools\psshutdown.exe" -d -t 0 -nobanner')
    print("Device is now active.")


@PThread.Throttle(15)
def Shutdown(request_confirmation=False) -> None:
    """Shuts down the computer."""
    
    if request_confirmation:
        if winHelper.ShowMessageBox("Do you really want to shutdown?", "Confirm Shutdown", 2, win32con.MB_ICONQUESTION) != 6:
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
    TerminateScript(graceful=False)


cpdef str GetProcessExe(int hwnd):
    """Given a window handle, returns the process executable path."""
    
    return win32process.GetModuleFileNameEx(win32api.OpenProcess(win32con.PROCESS_ALL_ACCESS, 0,
                                                                 win32process.GetWindowThreadProcessId(hwnd)[1]), 0)
