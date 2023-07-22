"""Main module for the whole script. Starts the entire script by calling the `main()` function from the `scriptRunner` extension module."""


def main():
    """Starts the script by calling the `beginScript()` function from the `scriptRunner` extension module."""
    import sys, os
    
    # Changing the working directory to where this script is.
    os.chdir(os.path.dirname(__file__))
    
    # This line adds the directory path of this module to the sys.path list.
    # sys.path is a list of strings that specifies the search path for Python modules.
    # By adding the directory path of the script file to this list, it allows Python to
    # locate and import any modules in that directory as well as any subdirectories within it.
    sys.path.append(os.path.dirname(__file__))
    
    from cythonExtensions.systemHelper import systemHelper as sysHelper
    
    
    if len(sys.argv) > 1 and any(arg in ("-e", "--elevated") for arg in sys.argv) and not sysHelper.isProcessElevated(-1):
        sysHelper.startWithElevatedPrivileges(terminate=False, cmder=True)
    
    elif len(sys.argv) > 1 and sys.argv[1] in ("-p", "--profile", "--prof"):
        from cythonExtensions.scriptRunner.scriptRunner import beginScript, profilerManager
        
        with profilerManager(engine="yappi"):
            beginScript()
    
    else:
        from cythonExtensions.scriptRunner.scriptRunner import beginScript
        
        beginScript()


if __name__ == '__main__':
    main()
