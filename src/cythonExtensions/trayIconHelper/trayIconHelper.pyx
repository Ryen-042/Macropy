# cython: language_level = 3str

from cythonExtensions.trayIconHelper cimport trayIconHelper

import win32api, win32con, win32gui_struct,  win32gui, os, importlib, glob, itertools, time, atexit, threading
from traceback import print_exc
from typing import Callable

import scriptConfigs as configs
from cythonExtensions.commonUtils.commonUtils import PThread, Management as mgmt
from cythonExtensions.systemHelper import systemHelper as sysHelper
from cythonExtensions.eventHandlers import eventHandlers

cdef class TrayIcon:
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
    
    cdef int default_sc_menu_action, default_dc_menu_action, FIRST_ID, _next_action_id, hwnd
    cdef float _counter
    cdef bint destroy_called
    cdef icon, hover_text, window_class_name, QUIT, on_quit, menu_options, menu_actions_set, menu_actions_dict, SPECIAL_ACTIONS, notify_id, icons, _timer
    
    def __init__(self, icons, hover_text: str, menu_options: tuple[tuple[str, str, object]], on_quit: Callable=None, window_class_name="Macropy", default_sc_menu_action: int=-1, default_dc_menu_action: int=-1):
        self.destroy_called = False
        
        self.FIRST_ID = 1023  # Starting ID for menu actions to avoid conflicts with system commands
        
        self.QUIT = 'Quit'
        self.SPECIAL_ACTIONS = [self.QUIT]
        
        self.icons = icons
        self.icon = next(self.icons)
        self.hover_text = hover_text
        
        # Adding 'Quit' option to menu_options
        menu_options = menu_options + ((self.QUIT, "", self.QUIT),)
        
        # Initialize action IDs and menu options
        self._next_action_id = self.FIRST_ID
        self.menu_actions_set = set()
        self.menu_options = self._addIdsToMenuOptions(menu_options)
        self.menu_actions_dict = dict(self.menu_actions_set)
        self._next_action_id = 0 # Reset the value of the temporary action ID
        
        self.on_quit = on_quit
        self.window_class_name = window_class_name
        
        self.default_sc_menu_action = (default_sc_menu_action if not None else -1)
        self.default_dc_menu_action = (default_dc_menu_action if not None else -1)
        
        atexit.register(self.destroyHandler, 0, 0, 0, 0)
        
        self.buildTrayIcon()
        
        self._timer = None  # A timer for delaying the execution of the single-click action to prevent its execution when double-clicking on the tray icon.
        self._counter = 0   # A delay counter that prevents the single-click action from being executed for a short period of time after double-clicking on the tray icon.

    cdef list[tuple] _addIdsToMenuOptions(self, tuple[tuple] menu_options):
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
        
        cdef tuple menu_option
        
        cdef list[tuple] result = []  # Initialize an empty list to store processed menu options.
        for menu_option in menu_options:
            option_text, option_icon, option_action = menu_option
            
            if callable(option_action) or option_action in self.SPECIAL_ACTIONS:
                # If option_action is callable (function or method) or a special action, add it to menu_actions_set.
                self.menu_actions_set.add((self._next_action_id, option_action))
                
                # Append the modified menu_option with an added ID to the result list.
                result.append(menu_option + (self._next_action_id,))
            
            elif nonStringIterable(option_action):
                # If option_action is an iterable (e.g., a sub-menu), recursively add IDs to its menu options.
                result.append((option_text,
                            option_icon,
                            self._addIdsToMenuOptions(option_action),
                            self._next_action_id))
            
            else:
                # Print a message for unknown items (unsupported option type).
                print(f"Unknown item specified: ({option_text}, {option_icon}, {option_action})")
            
            self._next_action_id += 1  # Increment the action ID for the next menu option.
        
        return result  # Return the list of processed menu options.

    cdef void buildTrayIcon(self):
        """Creates a system tray icon."""
        
        # Define message_map for handling Windows messages
        message_map = {
            # TaskbarCreated message is sent when the taskbar is created (e.g., during startup)
            win32gui.RegisterWindowMessage("TaskbarCreated"): self.restartHandler,
            # WM_DESTROY message is sent when the window is being destroyed
            win32con.WM_DESTROY: self.destroyHandler,
            # WM_COMMAND message is sent when the user selects a command item from a menu
            win32con.WM_COMMAND: self.commandHandler,
            # Custom user-defined message used for notifications
            win32con.WM_USER + 20: self.notifyHandler,
        }
        
        # Register the Window class
        window_class = win32gui.WNDCLASS()                                   # Create an instance of the WNDCLASS structure.
        hinst = window_class.hInstance = win32gui.GetModuleHandle(None)      # Set the instance for the window class.
        window_class.lpszClassName = self.window_class_name                  # Set the class name for the window class.
        window_class.style = win32con.CS_VREDRAW | win32con.CS_HREDRAW       # Set the window class style.
        
        window_class.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)    # Load the arrow cursor for the window class.
        # window_class.hCursor = win32gui.LoadImage(None, win32con.OCR_NORMAL, win32con.IMAGE_CURSOR, 0, 0, win32con.LR_SHARED)
        
        window_class.hbrBackground = win32con.COLOR_WINDOW                   # Set the background color for the window class.
        window_class.lpfnWndProc = message_map                               # Set the message map for handling window messages.
        classAtom = win32gui.RegisterClass(window_class)                     # Register the window class and obtain a class atom.
        
        # Create a window using the registered window class.
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = win32gui.CreateWindow(
            classAtom,
            self.window_class_name,
            style,
            0,
            0,
            win32con.CW_USEDEFAULT,
            win32con.CW_USEDEFAULT,
            0,
            0,
            hinst,
            None
        )
        
        win32gui.UpdateWindow(self.hwnd)  # Update the window to ensure it's displayed.
        self.notify_id = None  # Initialize the notification ID.
        self.refreshIcon()

    cdef void refreshIcon(self):
        """Refreshes the system tray icon by attempting to find an icon file and using the default icon if no icon file is found."""
        
        cdef int icon_flags, hicon, message
        
        hinst = win32gui.GetModuleHandle(None)  # Get the module handle for the current process.
        
        if os.path.isfile(self.icon):  # Check if the specified icon file exists.
            icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE  # Set flags for loading the icon from a file.
            hicon = win32gui.LoadImage(hinst,
                                    self.icon,
                                    win32con.IMAGE_ICON,
                                    0,
                                    0,
                                    icon_flags)  # Load the icon from the specified file.
        
        else:
            print("Can't find icon file - using default.")
            hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)  # Load the default application icon.
        
        if self.notify_id:  # Check if the notification ID already exists.
            message = win32gui.NIM_MODIFY  # Use the modify message if the notification ID exists.
        else:
            message = win32gui.NIM_ADD  # Use the add message if the notification ID doesn't exist.
        
        self.notify_id = (self.hwnd,
                        0,
                        win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP,
                        win32con.WM_USER+20,
                        hicon,
                        self.hover_text)  # (Re)Define the notification ID parameters.
        
        try:
            win32gui.Shell_NotifyIcon(message, self.notify_id)  # Send a system notification message.
        except:
            print("Failed to refresh the taskbar icon - is explorer running?")
            
            time.sleep(2)
            
            print("Rebuilding taskbar icon...")
            
            self.notify_id = None # Reset the notification ID to force the creation of a new try icon.
            
            self.refreshIcon()

    cdef bint restartHandler(self, int hwnd, int msg, int wparam, int lparam):
        """
        Handles the restart message (TaskbarCreated).
        
        This function refreshes the system tray icon.
        """
        
        self.refreshIcon()
        
        return True

    cdef bint destroyHandler(self, int hwnd, int msg, int wparam, int lparam):
        """
        Handles the destroy message (WM_DESTROY).
        
        This function executes the on_quit callback (if provided) and terminates the application.
        """
        
        # Check if the 'destroy' function has already been called to prevent infinite recursion when 'DestroyWindow' is called below.
        if self.destroy_called:
            return True
        
        self.destroy_called = True
        
        if self.on_quit is not None:
            self.on_quit(self)
        
        # Notify the shell that the icon is being removed
        cdef tuple[int, int] nid = (self.hwnd, 0)
        
        print("Destroying taskbar icon")
        try:
            win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
        except:
            print("Failed to destroy the taskbar icon - is explorer running?")
            print_exc()
        
        # Post the quit message to terminate the message loop and end the app
        win32gui.PostQuitMessage(0)
        
        # Destroy the window
        win32gui.DestroyWindow(self.hwnd)
        
        # Explicitly unregister the window class
        win32gui.UnregisterClass(self.window_class_name, None)
        
        atexit.unregister(self.destroyHandler)
        
        return True

    cdef bint notifyHandler(self, int hwnd, int msg, int wparam, int lparam):
        """
        Handles notification messages.
        
        This function responds to different mouse click events on the system tray icon.
        """
        
        if lparam == win32con.WM_LBUTTONDBLCLK:  # Handle double-click with left mouse button in the system tray icon
            self._timer.cancel()
            self.executeMenuOption(self.default_dc_menu_action + self.FIRST_ID)
            
            self._counter = time.time() + 0.5
        
        elif lparam == win32con.WM_LBUTTONUP:  # Handle left-click in the system tray icon (no action for now)
            if self._counter == 0 or time.time() - self._counter > 0:
                self._timer = threading.Timer(0.5, self.executeMenuOption, args=(self.default_sc_menu_action + self.FIRST_ID,))
                self._timer.start()
            
            self._counter = 0
        
        elif lparam == win32con.WM_RBUTTONUP:  # Handle right-click in the system tray icon
            self.showMenu()
        
        # elif lparam == win32con.WM_MBUTTONUP:  # Handle middle-click in the system tray icon
        #     pass
        
        return True  # Indicate that the message has been processed

    cdef void showMenu(self):
        """
        Displays the context menu.
        
        This function creates and displays the context menu at the current cursor position.
        """
        
        # Create a new popup menu
        menu = win32gui.CreatePopupMenu()
        
        # Populate the menu with items based on self.menu_options
        self.createMenu(menu, self.menu_options)
        
        pos = win32gui.GetCursorPos()
        
        # Set the foreground window to ensure the menu is displayed correctly
        win32gui.SetForegroundWindow(self.hwnd)
        
        # Display the popup menu at the cursor position
        win32gui.TrackPopupMenu(
            menu, win32con.TPM_LEFTALIGN, pos[0], pos[1], 0, self.hwnd, None
        )
        
        # Post a WM_NULL message to the window to ensure the menu is dismissed after use
        win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)

    cdef void createMenu(self, menu, list[tuple] menu_options):
        """
        Description:
            Creates a context menu and populates it with options and submenus.
        ---
        Parameters:
            - `menu`: Handle to the menu.
            - `menu_options -> list[tuple[str, str, str | tuple, int]]`: List of tuples representing the context menu options.
        """

        cdef int option_id
        cdef option_action, option_icon
        
        option_icon = None
        
        # Iterate through menu options in reverse order
        for option_text, option_icon_location, option_action, option_id in menu_options[::-1]:
            # If an icon is specified, prepare the menu icon
            if option_icon_location:
                option_icon = self.prepMenuIcon(option_icon_location)
            
            # Check if the option has a valid ID in the menu_actions_dict dictionary
            if option_id in self.menu_actions_dict:
                # Pack MENUITEMINFO structure for the menu item
                item, extras = win32gui_struct.PackMENUITEMINFO(text=option_text,
                                                                hbmpItem=option_icon,
                                                                wID=option_id)
                
                # Insert the menu item into the menu
                win32gui.InsertMenuItem(menu, 0, 1, item)
            
            else:
                # If the option has a submenu, create a new submenu
                submenu = win32gui.CreatePopupMenu()
                
                # Recursively create the submenu
                self.createMenu(submenu, option_action)
                
                # Pack MENUITEMINFO structure for the menu item with submenu
                item, extras = win32gui_struct.PackMENUITEMINFO(text=option_text,
                                                                hbmpItem=option_icon,
                                                                hSubMenu=submenu)
                
                # Insert the menu item with submenu into the menu
                win32gui.InsertMenuItem(menu, 0, 1, item)

    cdef prepMenuIcon(self, icon):
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
        
        # Check if the icon file exists
        if not os.path.isfile(icon):
            print('Menu icon file %s not found' % icon)
            return 0
        
        # Get the system metrics for the width and height of small icons
        ico_x = win32api.GetSystemMetrics(win32con.SM_CXSMICON)
        ico_y = win32api.GetSystemMetrics(win32con.SM_CYSMICON)

        # Load the icon from a file
        hicon = win32gui.LoadImage(0, icon, win32con.IMAGE_ICON, ico_x, ico_y, win32con.LR_LOADFROMFILE)

        # Create a compatible device context
        hdcBitmap = win32gui.CreateCompatibleDC(0)
        
        # Get the device context for the entire screen
        hdcScreen = win32gui.GetDC(0)
        
        # Create a compatible bitmap
        hbm = win32gui.CreateCompatibleBitmap(hdcScreen, ico_x, ico_y)
        
        # Select the compatible bitmap into the device context
        hbmOld = win32gui.SelectObject(hdcBitmap, hbm)

        # Get a system color brush for the menu background
        brush = win32gui.GetSysColorBrush(win32con.COLOR_MENU)
        
        # Fill the rectangle (16x16) with the menu background color
        win32gui.FillRect(hdcBitmap, (0, 0, 16, 16), brush)

        # Draw the icon onto the compatible bitmap
        win32gui.DrawIconEx(hdcBitmap, 0, 0, hicon, ico_x, ico_y, 0, 0, win32con.DI_NORMAL)
        
        # Restore the original bitmap into the device context
        win32gui.SelectObject(hdcBitmap, hbmOld)
        
        # Delete the compatible device context
        win32gui.DeleteDC(hdcBitmap)

        # Return the handle to the compatible bitmap
        return hbm

    cdef bint commandHandler(self, int hwnd, int msg, int wparam, int lparam):
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
        
        self.executeMenuOption(win32gui.LOWORD(wparam))
        
        return True

    cdef bint executeMenuOption(self, int id):
        """
        Description:
            Executes the action associated with the selected menu option.
        ---
        Parameters:
            id: ID of the selected menu option.
        """
        
        cdef object menu_action = self.menu_actions_dict.get(id)
        if menu_action is not None:
            if menu_action == self.QUIT:  # If the menu action is QUIT, destroy the window.
                win32gui.DestroyWindow(self.hwnd)
            
            elif callable(menu_action):  # If the menu action is callable, execute it with the current instance as an argument.
                PThread(target=menu_action, args=(self,)).start()
            
            elif nonStringIterable(menu_action):  # If it's iterable (like a submenu), show the submenu
                PThread(target=self.showMenu).start()
            else:
                print(f"Unexpected value for menu action with ID {id}: {menu_action}")
        
        # Ignore unknown menu action id
        # else:
        #     print(f"No action found for ID {id}.")
        
        return True

cdef bint nonStringIterable(obj):
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
    
    try:
        iter(obj)
    
    except TypeError:
        return False
    
    else:
        return not isinstance(obj, str)


cdef int switchIcon(TrayIcon trayIcon):
    trayIcon.icon = next(trayIcon.icons)
    
    trayIcon.refreshIcon()
    
    return 0


cdef int showNotification(TrayIcon trayIcon):
    sysHelper.sendScriptWorkingNotification(False)
    
    return 0


cdef int reloadHotkeys(TrayIcon trayIcon):
    importlib.reload(eventHandlers.cbs)
    
    return 0


cdef int reloadConfigs(TrayIcon trayIcon):
    sysHelper.reloadConfigs()
    
    return 0


cdef int openScriptFolder(TrayIcon trayIcon):
    os.startfile(configs.MAIN_MODULE_LOCATION)
    
    return 0


cdef int clearConsoleLogs(TrayIcon trayIcon):
    os.system("cls")
    
    return 0


cdef int toggleSilentMode(TrayIcon trayIcon):
    mgmt.toggleSilentMode()
    
    return 0


# cdef void createTrayIcon(int default_sc_menu_action=2, int default_dc_menu_action=0, hover_text="Macropy", on_quit=None):
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
    
    # Find available icon files or use default icons near this module
    icon_files = glob.glob(os.path.join(configs.MAIN_MODULE_LOCATION, "Images", "Icons", '*.ico'))
    
    if icon_files:
        icons = itertools.cycle(icon_files)
    else:
        icons = itertools.cycle(glob.glob(os.path.join(os.path.dirname(__file__), "icons", "*.ico")))
    
    # Define the menu options, including labels, icons, and associated actions
    menu_options = (
        ('Open Script Folder', "", openScriptFolder),
        ('Show Notification', next(icons), showNotification),
        ('Switch Icon', "", switchIcon),
        ('Reload', next(icons), (
            ('Hotkeys', next(icons), reloadHotkeys),
            ('Configs', next(icons), reloadConfigs),
        )),
        ('Clear Console Logs', "", clearConsoleLogs),
        ('Toggle Silent Mode', "", toggleSilentMode),
    )
    
    TrayIcon(icons, hover_text, menu_options, on_quit=on_quit, default_sc_menu_action=default_sc_menu_action, default_dc_menu_action=default_dc_menu_action)
