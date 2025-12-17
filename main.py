# main.py
import sys
import traceback
from PySide6.QtWidgets import QApplication, QMessageBox
from gui import RosterApp

# Global exception handler to prevent silent crashes
def exception_hook(exctype, value, tb):
    traceback_formated = ''.join(traceback.format_exception(exctype, value, tb))
    print(traceback_formated, file=sys.stderr)
    if QApplication.instance():
        error_msg = f"Critical Error:\n{exctype.__name__}: {value}\n\n{traceback_formated}"
        QMessageBox.critical(None, "Error", error_msg)
    else:
        sys.__excepthook__(exctype, value, tb)

if __name__ == "__main__":
    sys.excepthook = exception_hook
    app = QApplication(sys.argv)
    window = RosterApp()
    window.show()
    sys.exit(app.exec())