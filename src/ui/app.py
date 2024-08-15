import sys
from PyQt6.QtWidgets import QApplication, QMainWindow

from .main_window import MainWindow
from .list_file import ListFile


class App():
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.window = MainWindow()

        self.list_file = ListFile()
        self.window.show()

        self.window.centralWidget()
        
        

    def run(self):
        sys.exit(self.app.exec())

