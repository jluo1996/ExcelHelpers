from PyQt5.QtWidgets import QTextBrowser
from PyQt5.QtGui import QTextCursor
from colorama import Fore

class Logger():
    def __init__(self, textedit : QTextBrowser = None):
        self.text_browser = textedit

    def log_info(self, msg, scroll_to_bottom : bool = True):
        if self.text_browser:
            self._append_new_line(msg, 'black')   # Use default font color
            if scroll_to_bottom:
                self._scroll_to_bottom()
        self._print(msg, Fore.GREEN)

    def log_error(self, msg, scroll_to_bottom : bool = True):
        if self.text_browser:
            self._append_new_line(msg, 'red') 
            if scroll_to_bottom:
                self._scroll_to_bottom()
        self._print(msg, Fore.RED)

    def log_warning(self, msg, scroll_to_bottom: bool = True):
        if self.text_browser:
            self._append_new_line(msg, '#FFA500') 
            if scroll_to_bottom:
                self._scroll_to_bottom()
        self._print(msg, Fore.YELLOW)

    def _scroll_to_bottom(self):
        self.text_browser.moveCursor(QTextCursor.End)       # Move cursor to the end
        self.text_browser.ensureCursorVisible()             # Scroll so the cursor is visible   

    def _append_new_line(self, msg, color):
        """
        Parameter:
        msg: Could be a list[str] or str
        """
        if isinstance(msg, str):
                self.text_browser.append(f'<span style="color:{color};">{msg}</span>')
        elif isinstance(msg, list) and all(isinstance(item, str) for item in msg):
            for line in msg:
                self.text_browser.append(f'<span style="color:{color};">{line}</span>')

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

    