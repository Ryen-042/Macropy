"""This module contains the configuration variables used by the script."""
import os

MAX_ALIAS_LENGTH = 40
"""The maximum length of an alias."""

# If you want to use a prefix other than '!' and ':', to be able to use key expansion when 'suppressKbInputs' is set,
# you need to add an extra check in the 'eventHandlers.textExpansion' function.
LOCATIONS = {
            "!cmd":    r"C:\Windows\System32\cmd.exe",
            "!paint":  r"C:\Windows\System32\mspaint.exe",
            "!prog":   r"C:\Program Files"
            }
"""A dictionary of abbreviations and their corresponding path address. Abbreviationss must be lowercase only."""

ABBREVIATIONS = {
            ":py"    : "python",
            ":name" : "Ahmed Tarek",
            ":gmail" : "AhmedTarek37@gmail.com"
}
"""A dictionary of aliases and their corresponding expansion. The aliases must be in lowercase characters."""

NON_PERFIXED_ABBREVIATIONS= {
    "404err" : "We are doomed!",
    "143": "I love you",
    "nt": "Nice try",
}

SUPPRESS_TERMINAL_OUTPUT = True
"""A boolean value that determines whether the script should suppress all terminal output (except error messages) or not."""

ENABLE_DPI_AWARENESS = True
"""A boolean value that determines whether the script should be granted DPI awareness or not."""

ENABLE_LOGGING = True
"""A boolean value that determines whether the script should log and show error messages or not."""

ENABLE_ELEVATED_PRIVILEGES_CHECKER = True
"""A boolean value that determines whether the script should check for if the foreground window has elevated privileges or not."""

MAIN_MODULE_LOCATION = os.path.dirname(__file__)
"""Holds the full path address to the `__main__` module."""

ENABLE_SYSTEM_TRAY_ICON = True
"""A boolean value that determines whether the script should show a system tray icon or not."""
