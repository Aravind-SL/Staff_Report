import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QVBoxLayout

from .main_window import MainWindow
from .list_file import ListClass
from .variables_editor import VariableEditor


class App():
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.window = MainWindow()

        self.list_class = ListClass()
        self.window.show()

        vlayout = QVBoxLayout(self.window.centralWidget())
        vlayout.addWidget(self.list_class)

        self.editor = VariableEditor()
        vlayout.addWidget(self.editor)
        

        
        self.list_class.class_added.connect(self.editor.add_class)




    def run(self):
        sys.exit(self.app.exec())

