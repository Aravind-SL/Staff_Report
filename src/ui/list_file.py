import os
from typing import *
from model import WorkBookFile
from PyQt6.QtWidgets import (
    QWidget, 
    QLabel,
    QPushButton, 
    QVBoxLayout,
    QHBoxLayout,
    QListWidget, 
    QListWidgetItem,
    QFileDialog
)


class ListFile(QWidget):
    def __init__(self):
        super().__init__()
        self.add_btn = QPushButton("Add Workbook")
        self.add_btn.setObjectName("addButton")
        self.add_btn.clicked.connect(self.__add_file)
        
        layout = QVBoxLayout()
        layout.addWidget(self.add_btn)

        self.__files =  set({})
        self.file_list_widget = QListWidget()
        layout.addWidget(self.file_list_widget)


        layout.addStretch()
        self.rmv_btn =QPushButton("Remove") 
        self.rmv_btn.setObjectName("rmvButton")
        layout.addWidget(self.rmv_btn)
        self.rmv_btn.clicked.connect(self.__remove_file)

        self.update_files()
        self.setLayout(layout)

        self.__on_update = lambda self : None 

    @property
    def files(self):
        return self.__files


    def __remove_file(self):
    
        if self.file_list_widget.selectedIndexes():
            item = self.file_list_widget.currentIndex()

            self.__files.remove(item.data())
            self.file_list_widget.takeItem(item.row())


    def __add_file(self) -> None:
        diag = QFileDialog()
        path = diag.getOpenFileNames(None, 'Open a File', '.',  "Excel Files (*.xls *.xlsx *.xlsm)")
        if path != ([], ''):
            self.__files |=  set([
                WorkBookFile(x, None, None)
                for x in path[0]
            ])
            self.update_files()
            self.__on_update(self)

    def get_files(self) -> List[Union[str, os.PathLike]]:
        return self.__files

    def update_files(self) -> None:
        self.file_list_widget.clear()
        for it in map(lambda x: QListWidgetItem(str(x)), self.__files):
            self.file_list_widget.addItem(it)

    def on_list_update(self, cb):
        self.__on_update = cb

