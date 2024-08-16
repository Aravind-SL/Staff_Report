from .utils import set_status_message_for
from .main_window import MainWindow
from model import WorkBookFile
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QWidget, QPushButton, 
    QVBoxLayout, QHBoxLayout,
    QListWidget, QLabel,
    QFileDialog, QTabWidget,
    QFormLayout, QTextEdit
)

class VariableEditor(QTabWidget):
    def __init__(self):
        super().__init__()

        self.setMaximumHeight(300)
        self.nothing_tab = QWidget()
        self.addTab(self.nothing_tab, "Nothing")
        self.build_default_tab()


    def update_tabs(self, files):
        self.clear()

        if len(files) == 0:
            self.addTab(self.nothing_tab, "Nothing")
            return 

        for f in files:
            self.addTab(self.var_tab(f), str(f))

    def var_tab(self, wb: WorkBookFile) -> QWidget:
        wid = QWidget()
        layout = QFormLayout()
        
        slow_learner_input = QTextEdit(wb.slow_learners_threshold or "0")
        advanced_learner_input = QTextEdit(wb.advanced_learners_threshold or "0")
        layout.addRow(QLabel("Variables"))
        layout.addRow("Slow Learners Threshold", slow_learner_input)
        layout.addRow("Advanced Learners Threshold", advanced_learner_input)

        slow_learner_input.textChanged.connect(lambda : wb.set_slow_learners_threshold(int(slow_learner_input.toPlainText())))
        advanced_learner_input.textChanged.connect(lambda : wb.set_advanced_learners_threshold(int(advanced_learner_input.toPlainText())))
        process = QPushButton("Process")
        
        process.clicked.connect( lambda : set_status_message_for("Processing") )
        layout.addRow(process)

        wid.setLayout(layout)
        return wid

    def build_default_tab(self):
        nothing_layout = QVBoxLayout()
        nothing_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        nothing_layout.addWidget(
            QLabel("Nothing to show")
        )

        self.nothing_tab.setLayout(nothing_layout)
