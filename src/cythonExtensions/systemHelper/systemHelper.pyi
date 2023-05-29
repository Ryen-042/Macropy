"""This module provides system/script-specific functions."""


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
    ...


def IsProcessElevated(hwnd=0) -> bool:
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
    ...


def RequestElevation() -> None:
    """Restarts the current python process with elevated privileges."""
    ...


def ScheduleElevatedProcessChecker(delay=10.0) -> None:
    """Reprots each `delay` time if the active process window is elevated while the current python process is not elevated."""
    ...


def DisplayCPUsage() -> None:
    """Prints the current CPU and Memort usage to the console."""
    ...


def EnableDPI_Awareness() -> int:
    """Enables `DPI Awareness` for the current thread to allow for accurate dimensions reporting."""
    ...


def SendScriptWorkingNotification(near_module=True) -> None:
    """Sends a notification to the user that the script is working."""
    ...


def ChangeBrightness(opcode=1, increment=5) -> None:
    """Increments (`opcode=any non-zero value`) or decrements (`opcode=0`) the screen brightness by an (`increment`) percent."""
    ...


def ScreenOff():
    """Turns off the screen."""
    ...


def FlashScreen(delay=0.15) -> None:
    """Inverts the color of the screen for the specified number of seconds."""
    ...


def GoToSleep() -> None:
    """Puts the device to sleep."""
    ...


def Shutdown(request_confirmation=False) -> None:
    """Shuts down the computer."""
    ...


def GetProcessExe(hwnd: int) -> str:
    """Given a window handle, returns the process executable path."""
    ...
