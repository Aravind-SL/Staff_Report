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


        self.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(parent=self)
        self.statusbar.setObjectName("statusbar")
        self.setStatusBar(self.statusbar)

        get_signal_emitter().status_bar_signal_timeout.connect(
            self.statusBar().showMessage
        )

        get_signal_emitter().set_status_message.connect(
            self.statusBar().showMessage
        )

        self.setWindowTitle("Staff Report")

