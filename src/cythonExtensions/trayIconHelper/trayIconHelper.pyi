from typing import Callable

class TrayIcon:
    """
    Description:
        A class that implements a non-blocking system tray icon with a popup menu.
        
        - Please note that both this class and the `createTrayIcon` function must be called from the main thread. Otherwise, the system tray icon will not be displayed.
        
        - You are also required use listen for Windows messages using `GetMessageW` or `PumpMessages` in order to interact with the system tray icon.
        
        - For a usage example, see the `createTrayIcon` function at the end of this module (it doesn't listen for Windows messages).
        
        - This class is based on the example by [Simon Brunning](http://www.brunningonline.net/simon/blog/archives/SysTrayIcon.py.html) and the python 3 version by [Stephan Yazvinski](https://stackoverflow.com/a/65878630).
    
    ---
    Parameters:
        - `icons -> itertools.cycle`: An iterable of icon files to use for the system tray icon.
        - `hover_text -> str`: Text to display when hovering over the system tray icon.
        - `menu_options -> list[tuple[str, str, tuple[...] | str]]`: List of tuples representing the context menu options.
            - Each tuple contains (option_text, option_icon, option_action).
            - A `Quit` option is automatically added to the end of the menu options so you don't need to specify it.
        - `on_quit -> Callable`: Callback function to execute when the user quits the application.
        - `window_class_name -> str`: Name of the window class for the system tray icon window.
        - `default_sc_menu_action -> int`: Index of the default menu option (used for single-click).
        - `default_dc_menu_action -> int`: Index of the default menu option (used for double-click).    ---
    ---
    Methods:
        - `buildTrayIcon`: Create a system tray icon.
        - `refreshIcon`: Refresh the system tray icon.
        - `restartHandler`: Handle the restart message (TaskbarCreated).
        - `destroyHandler`: Handle the destroy message (WM_DESTROY).
        - `notifyHandler`: Handle notification messages.
        - `showMenu`: Display the context menu.
        - `createMenu`: Create the context menu.
        - `prepMenuIcon`: Prepare the menu icon.
        - `commandHandler`: Handle command messages (WM_COMMAND).
        - `executeMenuOption`: Execute the selected menu option.
    """
    
    __slots__ = (
        "default_sc_menu_action", "default_dc_menu_action", "FIRST_ID", "_next_action_id", "hwnd", # int
        "icon", "hover_text", "window_class_name", "QUIT", # str
        "destroy_called", # bool
        "on_quit", "menu_options", "menu_actions_set", "menu_actions_dict", "SPECIAL_ACTIONS", "notify_id", "icons", # object
    )
    
    
    def __init__(self, icons, hover_text: str, menu_options: tuple[tuple[str, str, object]], on_quit: Callable=None, window_class_name="", default_sc_menu_action: int=-1, default_dc_menu_action: int=-1):
        ...

    def _addIdsToMenuOptions(self, menu_options: tuple[tuple]) -> list[tuple[str, str, str | tuple, int]]:
        """
        Description:
            Adds unique IDs to menu options for identifying actions.
        ---
        Parameters:
            `menu_options -> tuple[tuple[str, str, str | tuple]]`: List of tuples representing the context menu options.
        
        ---
        Returns:
            `list[tuple[str, str, str | tuple, int]]`: List of tuples with added IDs for menu options.
        """
        ...
    
    def buildTrayIcon(self) -> None:
        """Creates a system tray icon."""
        ...
    
    def refreshIcon(self) -> None:
        """Refreshes the system tray icon by attempting to find an icon file and using the default icon if no icon file is found."""
        ...

    def restartHandler(self, hwnd: int, msg: int, wparam: int, lparam: int) -> bool:
        """
        Handles the restart message (TaskbarCreated).
        
        This function refreshes the system tray icon.
        """
        ...

    def destroyHandler(self, hwnd: int, msg: int, wparam: int, lparam: int) -> bool:
        """
        Handles the destroy message (WM_DESTROY).
        
        This function executes the on_quit callback (if provided) and terminates the application.
        """
        ...

    def notifyHandler(self, hwnd: int, msg: int, wparam: int, lparam: int) -> bool:
        """
        Handles notification messages.
        
        This function responds to different mouse click events on the system tray icon.
        """
        ...

    def showMenu(self) -> None:
        """
        Displays the context menu.
        
        This function creates and displays the context menu at the current cursor position.
        """
        ...

    def createMenu(self, menu, menu_options: list[tuple[str, str, str | tuple, int]]) -> None:
        """
        Description:
            Creates a context menu and populates it with options and submenus.
        ---
        Parameters:
            - `menu`: Handle to the menu.
            - `menu_options -> list[tuple[str, str, str | tuple, int]]`: List of tuples representing the context menu options.
        """
        ...

    def prepMenuIcon(self, icon: str):
        """
        Description:
            Prepares the menu icon. This function loads the menu icon, creates a compatible bitmap, fills the background,
            and draws the icon on the bitmap.
        ---
        Parameters:
            `icon -> str`: Path to the icon file.
        ---
        Returns:
            A Handle to the compatible bitmap.
        """
        ...

    def commandHandler(self,hwnd: int, msg: int, wparam: int, lparam: int) -> bool:
        """
        Description:
            Handles command messages (WM_COMMAND) and executes the selected menu option.
        ---
        Parameters:
            - `hwnd -> int`: Handle to the window.
            - `msg -> int`: Message identifier.
            - `wparam -> int`: First message parameter.
            - `lparam -> int`: Second message parameter.
        """
        ...

    def executeMenuOption(self, id: int) -> bool:
        """
        Description:
            Executes the action associated with the selected menu option.
        ---
        Parameters:
            id: ID of the selected menu option.
        """
        ...

def nonStringIterable(obj) -> bool:
    """
    Description:
        Checks if an object is iterable and not a string.
    ---
    Parameters:
        `obj`: Object to check for iterability.
    ---
    Returns:
        `bool`: `True` if the object is iterable and not a string, `False` otherwise.
    """
    ...


def switchIcon(trayIcon: TrayIcon) -> int:
    ...


def showNotification(trayIcon: TrayIcon) -> int:
    ...


def reloadHotkeys(trayIcon: TrayIcon) -> int:
    ...


def reloadConfigs(trayIcon: TrayIcon) -> int:
    ...


def openScriptFolder(trayIcon: TrayIcon) -> int:
    ...


def createTrayIcon(default_sc_menu_action=2, default_dc_menu_action=0, hover_text="Macropy", on_quit=None) -> None:
    """
    Description:
        A non-blocking function that creates a system tray icon with a popup menu.
        
        You are required to listen for Windows messages using `GetMessageW` or `PumpMessages` in order to interact with the system tray icon.
    
    ---
    Parameters:
        - `default_sc_menu_action -> int`: Index of the default menu option (used for single-click on the tray icon).
        - `default_dc_menu_action -> int`: Index of the default menu option (used for double-click on the tray icon).
        - `hover_text -> str`: Text to display when hovering over the system tray icon.
        - `on_quit -> Callable`: Callback function to execute when the user quits the application.
    """
    ...
