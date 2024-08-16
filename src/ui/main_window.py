from .signals import get_signal_emitter
from PyQt6 import QtCore, QtGui, QtWidgets
from PyQt6.QtWidgets import  QMainWindow, QFileDialog

class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()

        
        self.setObjectName("MainWindow")
        self.resize(800, 600)

        self.centralwidget = QtWidgets.QWidget(parent=self)
        self.centralwidget.setObjectName("centralwidget")

        self.menubar = self.menuBar()
        self.menuFile = self.menubar.addMenu("File")
        self.menuFile.setObjectName("menuFile")


        open_new_file = self.menuFile.addAction("Open new File")
        open_new_file.triggered.connect(self.open_file)

        self.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(parent=self)
        self.statusbar.setObjectName("statusbar")
        self.setStatusBar(self.statusbar)

        self.label()

        get_signal_emitter().status_bar_signal_timeout.connect(
            self.statusBar().showMessage
        )

        get_signal_emitter().set_status_message.connect(
            self.statusBar().showMessage
        )


    def open_file(self) -> str:

        diag = QFileDialog()
        path = diag.getOpenFileNames(None, 'Open a File', '.',  "Excel Files (*.xls *.xlsx *.xlsm)")

        return path

    def label(self):
        self.menuFile.setTitle("File")
        self.setWindowTitle("Staff Report")
