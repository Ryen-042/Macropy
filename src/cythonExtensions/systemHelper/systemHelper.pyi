"""This module provides system/script-specific functions."""

import win32con


def reloadConfigs() -> None:
    """Re-imports the `scriptConfigs` module and reloads the defined configurations."""
    
    ...


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
    ...


def isProcessElevated(hwnd=0) -> bool:
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


def startWithElevatedPrivileges(terminate=True, cmder=False, cmdShow=win32con.SW_SHOWNORMAL) -> int:
    """Starts another instance of the main python process with elevated privileges."""
    ...


def scheduleElevatedProcessChecker(delay=10.0) -> None:
    """Reprots each `delay` time if the active process window is elevated while the current python process is not elevated."""
    ...


def displayCPU_Usage() -> None:
    """Prints the current CPU and Memort usage to the console."""
    ...


def enableDPI_Awareness() -> int:
    """Enables `DPI Awareness` for the current thread to allow for accurate dimensions reporting."""
    ...


def sendScriptWorkingNotification(near_module=True) -> None:
    """Sends a notification to the user that the script is working."""
    ...


def changeBrightness(opcode=1, increment=5) -> None:
    """Increments (`opcode=any non-zero value`) or decrements (`opcode=0`) the screen brightness by an (`increment`) percent."""
    ...


def screenOff() -> None:
    """Turns off the screen."""
    ...

def toggleMonitorMode(mode: int):
    """Toggles the monitor mode between internal, external, extended, and clone."""
    ...

def flashScreen(delay=0.15) -> None:
    """Inverts the color of the screen for the specified number of seconds."""
    ...


def goToSleep() -> None:
    """Puts the device to sleep."""
    ...


def shutdown(request_confirmation=False) -> None:
    """Shuts down the computer."""
    ...


def suspendProcess(hwnd=0) -> int:
    """Suspends a process given its window handle. Uses the handle of the active window if no handle is passed."""
    ...


def resumeProcess(hwnd=0) -> int:
    """Resumes a suspended process given its window handle. Uses the handle of the active window if no handle is passed."""
    ...


def getHungwindowHandle(hwnd=0) -> int:
    """Returns the actual hwnd of a hung window given its ghost window handle. Uses the handle of the active window if no handle is passed."""
    ...


def isProcessSuspended(pid: int) -> bool:
    """Returns whether a process is suspended or not."""
    ...
