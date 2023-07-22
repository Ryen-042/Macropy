"""This module contains the configuration variables used by the script."""
import os

# If you want to use a prefix other than '!' and ':', to be able to use key expansion when 'suppressKbInputs' is set,
# you need to add an extra check in the 'eventHandlers.textExpansion' function.
LOCATIONS = {
            "!cmd":    r"C:\Windows\System32\cmd.exe",
            "!paint":  r"C:\Windows\System32\mspaint.exe",
            "!prog":   r"C:\Program Files"
            }
"""A dictionary of abbreviations and their corresponding path address. Abbreviations must be lowercase only."""

ABBREVIATIONS = {
            ":py"    : "python",
            ":name" : "name place holder",
            ":gmail" : "example37@gmail.com"
}
"""A dictionary of abbreviations and their corresponding expansion. Abbreviations must be lowercase only."""

SUPPRESS_TERMINAL_OUTPUT = True
"""A boolean value that determines whether the script should suppress all terminal output (except error messages) or not."""

ENABLE_DPI_AWARENESS = True
"""A boolean value that determines whether the script should be granted DPI awareness or not."""

ENABLE_LOGGING = True
"""A boolean value that determines whether the script should log and show error messages or not."""

ENABLE_ELEVATED_PRIVILEGES_CHECKER = True
"""A boolean value that determines whether the script should check for if the foreground window has elevated privileges or not."""

MAIN_MODULE_LOCATION = os.path.join(os.path.dirname(__file__), "__main__.py")
"""Stores the full path address of the `__main__` module."""
