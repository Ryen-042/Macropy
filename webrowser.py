import os
import sys

from PyQt5.QtCore import QUrl
from PyQt5.QtWidgets import QApplication, QMainWindow, QHBoxLayout, QVBoxLayout, QWidget, QPushButton, QLineEdit
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtGui import QColor
from PyQt5.QtWebEngineWidgets import QWebEngineProfile

class Browser(QMainWindow):
    def __init__(self):
        super().__init__()

        # Create a QWebEngineView widget
        self.view = QWebEngineView()

        # Create the back, forward and refresh buttons
        self.back_button = QPushButton("Back")
        self.forward_button = QPushButton("Forward")
        self.refresh_button = QPushButton("Refresh")

        # Connect the back, forward and refresh buttons to the appropriate slots
        self.back_button.clicked.connect(self.view.back)
        self.forward_button.clicked.connect(self.view.forward)
        self.refresh_button.clicked.connect(self.view.reload)

        # Create an address bar
        self.address_bar = QLineEdit()

        # Connect the address bar to the appropriate slot
        self.address_bar.returnPressed.connect(self.search)

        # Create a layout to hold the buttons and address bar
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.back_button)
        button_layout.addWidget(self.forward_button)
        button_layout.addWidget(self.refresh_button)
        button_layout.addWidget(self.address_bar)

        # Create a widget to hold the layout, and set it as the central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Create a layout to hold the view and button layout
        main_layout = QVBoxLayout()
        main_layout.addWidget(self.view)
        main_layout.addLayout(button_layout)

        # Set the main layout for the central widget
        central_widget.setLayout(main_layout)

        # Set the default URL
        self.view.setUrl(QUrl("https://www.google.com"))

        # Update the address bar when the URL changes
        self.view.urlChanged.connect(self.update_address_bar)

        # Apply the dark mode theme to the web page
        # self.view.setStyleSheet("""
            # * {
                # background-color: #303030 !important;
                # color: white !important;
            # }
        # """)
        self.view.page().setBackgroundColor(QColor("#606060"))

        # Apply the dark mode theme to the main window, buttons and address bar
        # self.setStyleSheet("""
            # QMainWindow {
                # background-color: #303030 !important;
                # color: black !important;
            # }
            # QPushButton {
                # background-color: #606060 !important;
                # color: black !important;
            # }
            # QLineEdit {
                # background-color: #606060 !important;
                # color: black !important;
                # border: #303030 !important;
            # }
            # QMenuBar {
                # background-color: #303030 !important;
                # color: black !important;
            # }
        # """)
        # self.setStyleSheet("""
            # ::-webkit-scrollbar {
                # width: 12px;
                # background-color: #303030;
            # }
            # ::-webkit-scrollbar-thumb {
                # background-color: #606060;
            # }
        # """)

    def search(self):
        query = self.address_bar.text()
        if "." not in query:
            # Assume the user is searching for something
            url = "https://www.google.com/search?q={}".format(query)
        else:
            # Assume the user is entering a URL
            url = QUrl.fromUserInput(query)
        self.view.setUrl(QUrl(url))

    def update_address_bar(self, url):
        self.address_bar.setText(url.toString())

if __name__ == '__main__':
    # Source: https://pypi.org/project/qt-material/
    from qt_material import apply_stylesheet

    # create the application and the main window
    app = QApplication(sys.argv)
    browser = Browser()
    
    # setup stylesheet
    apply_stylesheet(app, theme='dark_purple.xml')
    
    
    browser.show()
    
    app.exec_()
    
    # Clear the cookies, cache, and browsing history
    QWebEngineProfile.defaultProfile().clearAllVisitedLinks()
    browser.view.page().profile().cookieStore().deleteAllCookies()
    
    from PyQt5.QtWebEngineWidgets import QWebEngineProfile
    # # Clear the cache and browsing history
    QWebEngineProfile.defaultProfile().clearAllVisitedLinks()
    
    # method to clear the browsing history
    browser.view.page().history().clear()
    
    sys.exit()
