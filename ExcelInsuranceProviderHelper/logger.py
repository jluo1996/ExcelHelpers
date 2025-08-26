from PyQt5.QtWidgets import QTextEdit
from PyQt5.QtGui import QTextCursor
from colorama import Fore

class Logger():
    def __init__(self, textedit : QTextEdit = None):
        self.text_edit = textedit

    def log_info(self, msg, scroll_to_bottom : bool = True):
        if self.text_edit:
            self._append_new_line(msg, 'black')   # Use default font color
            if scroll_to_bottom:
                self._scroll_to_bottom()
        self._print(msg, Fore.GREEN)

    def log_error(self, msg, scroll_to_bottom : bool = True):
        if self.text_edit:
            self._append_new_line(msg, 'red') 
            if scroll_to_bottom:
                self._scroll_to_bottom()
        self._print(msg, Fore.RED)

    def log_warning(self, msg, scroll_to_bottom: bool = True):
        if self.text_edit:
            self._append_new_line(msg, 'yellow') 
            if scroll_to_bottom:
                self._scroll_to_bottom()
        self._print(msg, Fore.YELLOW)

    def _scroll_to_bottom(self):
        self.text_edit.moveCursor(QTextCursor.End)       # Move cursor to the end
        self.text_edit.ensureCursorVisible()             # Scroll so the cursor is visible   

    def _append_new_line(self, msg, color):
        """
        Parameter:
        msg: Could be a list[str] or str
        """
        if isinstance(msg, str):
                safe_text = msg.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                self.text_edit.append(f'<span style="color:{color};">{safe_text}</span>')
        elif isinstance(msg, list) and all(isinstance(item, str) for item in msg):
            for line in msg:
                safe_text = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                self.text_edit.append(f'<span style="color:{color};">{safe_text}</span>')

    def _print(self, msg, color):
        """
        Parameter:
        msg: Could be a list[str] or str
        """
        if isinstance(msg, str):
                print(color + msg)
        elif isinstance(msg, list) and all(isinstance(item, str) for item in msg):
            for line in msg:
                print(color + line)

    