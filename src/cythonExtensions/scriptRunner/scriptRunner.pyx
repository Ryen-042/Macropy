# cython: embedsignature = True
# cython: language_level = 3str

"""This extension module contains the main entry for the entire script. The `main.py` module calls this extension to start the script."""

import os


def AcquireScriptLock() -> int:
    """Acquires the script lock. This is used to prevent multiple instances of the script from running at the same time.
    If another instance of the script is already running, this instance will be terminated."""
    
    # Source: https://www.oreilly.com/library/view/python-cookbook/0596001673/ch06s09.html
    from win32event import CreateMutex
    from win32api import GetLastError
    from winerror import ERROR_ALREADY_EXISTS
    
    #+ Making sure this is the only running instance of the script by creating a mutex.
    #? If another process tried to create a new mutex with the same name, the handle to the existing
    # mutex object is returned, and an error is reported, which can be examined using `GetLastError()`.
    handle = CreateMutex(None, 1, 'Macropy')
    
    if GetLastError() == ERROR_ALREADY_EXISTS:
        import winsound
        
        print("Warning! Another instance of the script is already running.")
        print("Please close the other instance before running a new one.")
        print("Terminating this instance...")
        winsound.PlaySound(r"SFX\denied.wav", winsound.SND_FILENAME)
        
        os._exit(1)
    
    return handle


cpdef void begin_script():
    """The main entry for the entire script. Acquires the script lock then configures and starts the keyboard listeners and other components."""
    
    # Making sure that this is the only running instance of the main function. If not, then terminate this one.
    mutexHandle = AcquireScriptLock()
    print("Script lock acquired.")
    
    import threading
    import winsound, pythoncom
    from cythonExtensions.systemHelper import systemHelper as sysHelper
    from cythonExtensions.commonUtils.commonUtils import Management as mgmt, PThread
    from cythonExtensions.eventListeners.eventListeners import KeyPress, KeyRelease
    from cythonExtensions.hookManager.hookManager import HookManager, KeyboardHookManager
    from time import sleep, time
    import scriptConfigs as configs
    
    winsound.PlaySound(r"SFX\achievement-message-tone.wav", winsound.SND_FILENAME|winsound.SND_ASYNC)
    print("Initializing script...")
    
    #+ Initializing the uncaught exception logger.
    if configs.ENABLE_LOGGING:
        import sys
        sys.excepthook = mgmt.LogUncaughtExceptions
    
    #+ Initializing the COM library for the main thread. Sets the current thread to be a COM apartment thread: https://stackoverflow.com/questions/21141217/how-to-launch-win32-applications-in-separate-threads-in-python
    #? As long as the Automation object is used in the same thread in which it was created, the COM library is already initialized and you do not need to call the CoInitialize function so this line is redundant.
    pythoncom.CoInitialize()
    
    #+ For a valid window size reporting.
    if configs.ENABLE_DPI_AWARENESS:
        sysHelper.EnableDPI_Awareness()
    
    #+ Scheduling a checker to notify if a process with elevated privileges is active when the script does not have elevated privileges.
    #? This is necessary because no keyboard events are reported when this scenario happens.
    if configs.ENABLE_ELEVATED_PRIVILEGES_CHECKER and not sysHelper.IsProcessElevated(-1):
        print("Starting the elevated processes checker...")
        PThread(target=sysHelper.ScheduleElevatedProcessChecker).start()
    
    #+ Initializing the keyboard hook manager.
    # kbHook = pyWinhook.HookManager()
    hookManager = HookManager()
    kbHook = KeyboardHookManager()
    
    #+ Initializing all keyboard event callbakcs.
    print("Initializing keyboard listeners...")
    
    kbHook.addKeyDownListener(KeyPress)
    kbHook.addKeyUpListener(KeyRelease)
    
    #+ Starting the program main loop.
    print("Activating keyboard listeners...\n")
    
    ## Installing the low level hook.
    if not hookManager.InstallHook(kbHook.KeyboardHook):
        print("Failed to install hook!")
        os._exit(1)
    
    # Begin listening for windows events. This function will not return until the hook stops.
    hookManager.BeginListening()
    
    ##! Reaching this point means that the script is being terminated.
    
    # Uninstall the hook.
    print("Uninstalling the hook...")
    hookManager.UninstallHook()
    
    # Wait for a certain number of seconds before forcefully stopping the running threads.
    cdef float countdown_start = time()
    
    # Get a list of all running threads
    cdef list alive_threads = threading.enumerate()
    
    # Count the number of threads in the list
    cdef int still_alive = len(alive_threads)
    
    # A flag to break the outer loop.
    cdef bint break_outer = False
    
    for thread in alive_threads:
        if thread != threading.main_thread():
            print(f"{still_alive} thread{'s are' if still_alive > 1 else ' is'} still active.", end="\r")
            while True:
                if time() - countdown_start >= 10:
                    print(f"\n{still_alive} threads are still running after 10s wait. The threads will be forcefully terminated...")
                    break_outer = True
                    break
                
                if thread.is_alive():
                    sleep(0.2)
                else:
                    print(f"{still_alive} threads are still active.", end="\r")
                    break
        
        still_alive -= 1
        
        if break_outer:
            break
    
    else:
        #! Un-initializing the COM library if the main loop is terminated.
        print("Un-initializing COM library in the main thread...")
        pythoncom.CoUninitialize()
    
    # Releasing the acquired script lock.
    from win32event import ReleaseMutex
    ReleaseMutex(mutexHandle)
    print("Script lock released.")
    
    #! Terminate the script.
    sysHelper.TerminateScript(graceful=False)


from contextlib import contextmanager


# Source: https://dev.to/rydra/getting-started-on-profiling-with-python-3a4
# Useful: https://coderzcolumn.com/tutorials/python/yappi-yet-another-python-profiler, https://github.com/sumerc/yappi/blob/master/doc/api.md
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
    
    from datetime import datetime as dt
    
    # Making a directory to store the profiling results.
    cdef str output_location
    if save_near_module:
        output_location = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dumpfiles")
    else:
        output_location = os.path.join(os.getcwd(), "dumpfiles")
    
    os.makedirs(output_location, exist_ok=True)
    
    if not filename:
        output_location = os.path.join(output_location, f"{dt.now().strftime('%Y-%m-%d (%I%p-%M-%S)')}.prof")
    else:
        output_location = os.path.join(output_location, filename)
    
    print("PROFILING ENABLED.")
    
    if engine == 'yappi':
        import yappi
        
        try:
            yappi.set_clock_type(clock)
            yappi.start(builtins=profile_builtins, profile_threads=profile_threads)
            
            # The yield statement is used to temporarily suspend the execution of the context manager and return control to the caller.
            # When the context manager is exited (either normally or due to an exception), the code after the yield statement is
            # executed to clean up any resources used by the context manager.
            yield
        
        finally:
            yappi.stop()
            
            print(f"Dumping profile to: {output_location}\n")
            
            yappi.get_func_stats().save(output_location, type=output_type)
            
            yappi.get_thread_stats().print_all()
    
    else:
        import cProfile
        
        profiler = cProfile.Profile()
        try:
            profiler.enable()
            yield
        
        finally:
            profiler.disable()
            profiler.print_stats()
            profiler.dump_stats(output_location)
            
            from pyprof2calltree import convert, visualize
            
            print("Saving the profiling results as `kgrind`...")
            convert(profiler.getstats(), os.path.join(os.path.splitext(output_location)[0], 'profiling_results.kgrind'))
            
            # `visualize` requires you have a separate program to work. You can download this and add it to the
            # system's path environment vairable: https://sourceforge.net/projects/qcachegrindwin/files/0.7.4/
            print("visualize the profiling results...")
            visualize(profiler.getstats())


cpdef void begin_script_with_cProfile(save_near_module=False):
    """Starts the main script with profiling."""
    
    import cProfile, pstats
    
    print("PROFILING ENABLED.")
    
    with cProfile.Profile() as profile:
        begin_script()
    
    from datetime import datetime as dt
    
    # Printing the profiling results.
    profiling_results = pstats.Stats(profile)
    profiling_results.sort_stats(pstats.SortKey.TIME)
    profiling_results.print_stats()
    
    # Making a directory to store the profiling results.
    cdef str dump_loc
    if save_near_module:
        dump_loc = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dumpfiles")
    else:
        dump_loc = os.path.join(os.getcwd(), "dumpfiles")
    
    os.makedirs(dump_loc, exist_ok=True)
    
    # Dumping the profiling results to a file. The file name is the current date and time.
    print("Dumping profile to:", end=" ")
    dump_loc = os.path.join(dump_loc, f"{dt.now().strftime('%Y-%m-%d (%I%p-%M-%S)')}.prof")
    print(dump_loc)
    
    # Dumping the profiling results.
    profiling_results.dump_stats(dump_loc)


cpdef void begin_script_with_profiling(filename="", engine="yappi", clock="wall", output_type="pstat", profile_builtins=True, profile_threads=True, save_near_module=False):
    """Starts the main script with profiling."""
    
    with profilerManager(filename=filename, engine=engine, clock=clock, output_type=output_type, profile_builtins=profile_builtins, profile_threads=profile_threads, save_near_module=save_near_module):
        begin_script()
