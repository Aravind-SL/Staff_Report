from .signals import get_signal_emitter

def set_status_message(text: str):
    get_signal_emitter().set_status_message.emit(text)

def set_status_message_for(text: str, timeout=1000):
    get_signal_emitter().status_bar_signal_timeout.emit(text, timeout)

