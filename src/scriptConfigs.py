"""This module contains the configuration variables used by the script."""

# If you want to use a prefix other than '!' and ':', to be able to use key expansion when 'suppress_all_keys' is set,
# you need to add an extra check in the 'eventListeners.ExpanderEvent' function.
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

SHOW_PRESSED_KEYS = False
"""A boolean value that determines whether the script should show the pressed keys by default or not."""

ENABLE_DPI_AWARENESS = True
"""A boolean value that determines whether the script should be granted DPI awareness or not."""

ENABLE_LOGGING = True
"""A boolean value that determines whether the script should log and show error messages or not."""

ENABLE_ELEVATED_PRIVILEGES_CHECKER = True
"""A boolean value that determines whether the script should check for if the foreground window has elevated privileges or not."""
