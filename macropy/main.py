"""Main module for the whole script. Starts the entire script by calling the `main()` function from the `scriptController` extension module."""

if __name__ == '__main__':
    import sys, os
    from scriptController import main_with_profiling, main
    
    #+ Changing the working directory to where this script is.
    os.chdir(os.path.dirname(__file__))
    
    if len(sys.argv) > 1 and sys.argv[1] in ("-p", "--profile", "--prof"):
        main_with_profiling()
    else:
        main()
