"""This module contains the main entry for the entire script. The `main.py` module calls this extension to start the script."""

def AcquireScriptLock() -> int:
    """Acquires the script lock. This is used to prevent multiple instances of the script from running at the same time.
    If another instance of the script is already running, this instance will be terminated."""
    ...


def begin_script() -> None:
    """The main entry for the entire script. Acquires the script lock then configures and starts the keyboard listeners and other components."""
    ...

def begin_script_with_profiling(save_near_module=True) -> None:
    """Starts the main script with profiling."""
    ...
