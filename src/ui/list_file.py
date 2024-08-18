import os
import glob
from typing import *
from model import CourseFile, ClassReportInput
from  PyQt6 import QtCore
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

class ListCourseFiles(QListWidget):
    def __init__(self, class_report: ClassReportInput):
        super().__init__()
        self.addItems(list(map(str, class_report.course_reports_input)))
        self.setStyleSheet("border: 1px solid;")



class ListClass(QWidget):
    class_added = QtCore.pyqtSignal(ClassReportInput)
    class_removed = QtCore.pyqtSignal(ClassReportInput)
    def __init__(self):
        super().__init__()
        self.add_btn = QPushButton("Add Class")
        self.add_btn.setObjectName("addButton")
        self.add_btn.clicked.connect(self.__add_class)
        
        layout = QVBoxLayout()
        layout.addWidget(self.add_btn)

        self.__classes =  {}
        sub_layout = QHBoxLayout()
        
        self.class_list = QListWidget()

        self.class_list.itemSelectionChanged.connect(self.__select_class)
        self.class_list.setMaximumWidth(self.width()//2)
        sub_layout.addWidget(self.class_list)

        self.file_list_holder = QHBoxLayout()
        self.file_list = QLabel("Select Class")
        self.file_list_holder.addWidget(self.file_list)

        sub_layout.addLayout(self.file_list_holder)

        layout.addLayout(sub_layout)

        layout.addStretch()
        self.rmv_btn =QPushButton("Remove") 
        self.rmv_btn.setObjectName("rmvButton")
        layout.addWidget(self.rmv_btn)
        self.rmv_btn.clicked.connect(self.__remove_class)

        self.update_classes()
        self.setLayout(layout)

    @property
    def classes(self):
        return self.__classes

    def __select_class(self, unset=False):
        if self.class_list.selectedIndexes() and not unset:
            item = self.class_list.currentIndex()
            cls = self.__classes[item.data()]
            w = ListCourseFiles(cls)
        else:
            w = QLabel("Select Class")
        self.file_list_holder.replaceWidget(self.file_list , w)
        self.file_list.deleteLater()
        self.file_list = w


    def __remove_class(self):
        if self.class_list.selectedIndexes():
            item = self.class_list.currentIndex()
            self.class_list.selectionModel().clear()
            self.class_removed.emit(self.__classes[item.data()])
            del self.__classes[item.data()]
            self.class_list.takeItem(item.row())


    def __add_class(self) -> None:
        diag = QFileDialog()
        path = diag.getExistingDirectory(None, 'Select Class', '.')
        if path is not None and path != '':

            files =  set([
                CourseFile(x, None, None)
                for x in glob.glob(f"{path}/*.xlsx")
            ])
            name = path.split('/')[-1]
            self.__classes[name] = ClassReportInput(name, files)
            self.update_classes()

            self.class_added.emit(self.__classes[name])

    def update_classes(self) -> None:
        self.class_list.clear()
        for it in map(QListWidgetItem, self.__classes.keys()):
            self.class_list.addItem(it)
