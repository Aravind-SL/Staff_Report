import os
from typing import *
from PyQt6.QtWidgets import QWidget, QPushButton, QVBoxLayout, QListWidget, QListWidgetItem

class ListFile(QWidget):
    def __init__(self):
        super().__init__()
        self.add_btn = QPushButton("Some Button")
        self.add_btn.setObjectName("addButton")
        self.add_btn.clicked.connect(self.__add_file)
        
        self.layout.addWidget(self.add_btn)

        self.__files =  set({})
        self.file_list_widget = QListWidget()
        self.layout.addWidget(self.file_list_widget)

        self.update_files()


    def __add_file(self) -> None:
        diag = QFileDialog()
        path = diag.getOpenFileNames(None, 'Open a File', '.',  "Excel Files (*.xls *.xlsx *.xlsm)")
        if path != ([], ''):
            self.__files |=  set(path[0])
            self.update_files()

    def get_files(self) -> List[Union[str, os.PathLike]]:
        return self.__files

    def update_files(self) -> None:
        self.file_list_widget.clear()
        self.file_list_widget.addItems(self.__files)

