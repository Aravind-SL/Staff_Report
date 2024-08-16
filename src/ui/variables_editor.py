from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QWidget, QPushButton, QVBoxLayout,
    QListWidget, QLabel,
    QFileDialog, QTabWidget
)

class VariableEditor(QTabWidget):
    def __init__(self):
        super().__init__()

        self.nothing_tab = QWidget()
        self.addTab(self.nothing_tab, "Nothing")
        self.build_default_tab()


    def update_tabs(self, files):
        self.clear()

        if len(files) == 0:
            self.addTab(self.nothing_tab, "Nothing")
            return 

        for f in files:
            self.addTab(QLabel(f), f.split("/")[-1])


    def build_default_tab(self):
        nothing_layout = QVBoxLayout()
        nothing_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        nothing_layout.addWidget(
            QLabel("Nothing to show")
        )

        self.nothing_tab.setLayout(nothing_layout)
