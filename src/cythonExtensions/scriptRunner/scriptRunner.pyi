"""This module contains the main entry for the entire script. The `main.py` module calls this extension to start the script."""

from contextlib import contextmanager


def AcquireScriptLock() -> int:
    """Acquires the script lock. This is used to prevent multiple instances of the script from running at the same time.
    If another instance of the script is already running, this instance will be terminated."""
    ...


def begin_script() -> None:
    """The main entry for the entire script. Acquires the script lock then configures and starts the keyboard listeners and other components."""
    ...


@contextmanager
def profilerManager(filename="", engine="yappi", clock="wall", output_type="pstat", profile_builtins=True, profile_threads=True, save_near_module=False):
    """
    Description:
        A context manager that can be used to profile a block of code.
    ---
    Parameters:
        `filename -> str`
            The output file name. Defaults to the date of running this context manager `"Y-m-d (Ip-M-S).prof"`.
        
        `engine -> str`
            Selects one of the next two profilers: `yappi`, `cprofiler`.
        
        `clock -> str`
            Sets the underlying clock type (`wall` or `cpu`).
        
        `output_type -> str`
            The target type that the profile stats will be saved in. Can be either "pstat" or "callgrind".
        
        `profile_builtins -> bool`
            Enable profiling for built-in functions.
        
        `profile_threads -> bool`
            Enable profiling for all threads or just the main thread.
        
        `save_near_module -> bool`
            Selects where to save the output file. `True` will save the file relative to this module's location,
            and `False` will save it relative to the `__main__` module.
    ---
    Usage:
    >>> with profile():
            # Some code.
    """
    ...


def begin_script_with_cProfile(save_near_module=False):
    """Starts the main script with profiling."""
    ...
