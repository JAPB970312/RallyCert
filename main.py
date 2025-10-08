import sys
from PyQt6.QtWidgets import QApplication
from ui import App
from auto_updater import auto_update

if __name__ == '__main__':
    app = QApplication(sys.argv)
    auto_update(app)  # Verifica actualizaciones con QMessageBox
    window = App()
    window.show()
    sys.exit(app.exec())