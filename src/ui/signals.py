from typing import Optional
from PyQt6.QtCore import pyqtSignal, QObject

class SignalEmitter(QObject):
    status_bar_signal_timeout = pyqtSignal(str, int)
    set_status_message = pyqtSignal(str)


__instance = SignalEmitter()

def get_signal_emitter() -> SignalEmitter:
    return __instance

