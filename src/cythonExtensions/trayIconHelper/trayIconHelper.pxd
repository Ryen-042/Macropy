cdef class TrayIcon:
    
    cdef int default_sc_menu_action, default_dc_menu_action, FIRST_ID, _next_action_id, hwnd
    cdef float _counter
    cdef bint destroy_called
    cdef icon, hover_text, window_class_name, QUIT, on_quit, menu_options, menu_actions_set, menu_actions_dict, SPECIAL_ACTIONS, notify_id, icons, _timer
    
    cdef list[tuple] _addIdsToMenuOptions(self, tuple[tuple] menu_options)
    
    cdef void buildTrayIcon(self)
    
    cdef void refreshIcon(self)

    cdef bint restartHandler(self, int hwnd, int msg, int wparam, int lparam)

    cdef bint destroyHandler(self, int hwnd, int msg, int wparam, int lparam)

    cdef bint notifyHandler(self, int hwnd, int msg, int wparam, int lparam)

    cdef void showMenu(self)

    cdef void createMenu(self, menu, list[tuple] menu_options)

    cdef prepMenuIcon(self, icon)

    cdef bint commandHandler(self, int hwnd, int msg, int wparam, int lparam)

    cdef bint executeMenuOption(self, int id)

cdef int switchIcon(TrayIcon trayIcon)

cdef int showNotification(TrayIcon trayIcon)

cdef int reloadHotkeys(TrayIcon trayIcon)

cdef int reloadConfigs(TrayIcon trayIcon)

cdef int openScriptFolder(TrayIcon trayIcon)
