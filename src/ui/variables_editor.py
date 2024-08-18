from typing import Dict
from .utils import set_status_message_for
from .main_window import MainWindow
from model import CourseFile, ClassReportInput
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QWidget, QPushButton, 
    QVBoxLayout, QHBoxLayout,
    QListWidget, QLabel,
    QFileDialog, QTabWidget,
    QFormLayout, QLineEdit,
    QScrollArea
)

class VariableEditor(QTabWidget):
    def __init__(self):
        super().__init__()

        self.setMaximumHeight(300)


    def add_class(self, cls: ClassReportInput):
        self.addTab(self.var_tab(cls), str(cls))

    def var_tab(self, cr: ClassReportInput) -> QWidget:
        tab_wid = QWidget()
        vlayout = QVBoxLayout()
        hlayout = QHBoxLayout()
        vlayout.addLayout(hlayout)

        class_vars_layout = QFormLayout()
        class_vars_layout.addWidget(QLabel(cr.name))

        top_n_editor = QLineEdit("")
        top_n_editor.setInputMask("00")
        top_n_editor.setMaximumWidth(self.width()//4)
        class_vars_layout.addRow("Top N", top_n_editor)

        hlayout.addLayout(class_vars_layout)

        hlayout.addWidget(CourseVarsList(cr.course_reports_input))

        process = QPushButton("Process")
        process.clicked.connect( lambda : set_status_message_for("Processing") )
        vlayout.addWidget(process)

        tab_wid.setLayout(vlayout)
        return tab_wid

    def build_default_tab(self):
        nothing_layout = QVBoxLayout()
        nothing_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        nothing_layout.addWidget(
            QLabel("Nothing to show")
        )

        self.nothing_tab.setLayout(nothing_layout)


class CourseVarsList(QScrollArea):
    input_height = 40
    input_mask = "00"
    def __init__(self, courses):
        super().__init__()

        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.setWidgetResizable(True)



        container = QWidget()
        self.setWidget(container)
        layout = QFormLayout()
        container.setLayout(layout)
        container.setStyleSheet("background: #202020;")


        for course in courses:
            slow_learner_input = QLineEdit("")
            slow_learner_input.setMaximumHeight(self.input_height)
            slow_learner_input.setInputMask(self.input_mask)
            slow_learner_input.setCursorPosition(0)
            advanced_learner_input = QLineEdit("")
            advanced_learner_input.setMaximumHeight(self.input_height)
            advanced_learner_input.setInputMask(self.input_mask)
            advanced_learner_input.setCursorPosition(0)
            label = QLabel(str(course), alignment=Qt.AlignmentFlag.AlignLeft)
            layout.addWidget(label)
            layout.addRow("Slow Learners Threshold", slow_learner_input)
            layout.addRow("Advanced Learners Threshold", advanced_learner_input)

            slow_learner_input.textChanged.connect(lambda x: x.isnumeric() and course.set_slow_learners_threshold(int(x)))
            advanced_learner_input.textChanged.connect(lambda x: x.isnumeric() and course.set_advanced_learners_threshold(int(x)))

