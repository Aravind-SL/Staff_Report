import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QVBoxLayout

from .main_window import MainWindow
from .list_file import ListFile
from .variables_editor import VariableEditor


class App():
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.window = MainWindow()

        self.list_file = ListFile()
        self.window.show()

        vlayout = QVBoxLayout(self.window.centralWidget())
        vlayout.addWidget(self.list_file)

        self.editor = VariableEditor()
        vlayout.addWidget(self.editor)
        

        
        self.list_file.on_list_update(lambda li: self.editor.update_tabs(li.files))




    def run(self):
        sys.exit(self.app.exec())

