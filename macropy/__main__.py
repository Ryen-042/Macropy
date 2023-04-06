"""Main module for the whole script. Starts the entire script by calling the `main()` function from the `scriptController` extension module."""

def main():
    """Starts the script by calling the `begin_script()` function from the `scriptController` extension module."""
    import sys, os
    
    #+ Changing the working directory to where this script is.
    os.chdir(os.path.dirname(__file__))
    
    # This line adds the directory path of this module to the sys.path list.
    # sys.path is a list of strings that specifies the search path for Python modules.
    # By adding the directory path of the script file to this list, it allows Python to
    # locate and import any modules in that directory as well as any subdirectories within it.
    sys.path.append(os.path.dirname(__file__))
    
    from scriptController import begin_script_with_profiling, begin_script
    
    if len(sys.argv) > 1 and sys.argv[1] in ("-p", "--profile", "--prof"):
        begin_script_with_profiling()
    else:
        begin_script()

if __name__ == '__main__':
    main()
