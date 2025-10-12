# main.py
import sys
import os
from PyQt6.QtWidgets import QApplication
from ui import App
from auto_updater import auto_update

def get_local_commit_sha():
    """Lee la versión instalada desde la carpeta persistente del usuario."""
    try:
        base_dir = os.path.join(os.path.expanduser("~"), ".RallyCert")
        commit_file = os.path.join(base_dir, "commit.sha")
        if os.path.exists(commit_file):
            with open(commit_file, "r", encoding="utf-8") as f:
                sha = f.read().strip()
                if sha:
                    return sha[:7]  # SHA abreviado
        return "Desconocida"
    except Exception as e:
        print(f"⚠️ No se pudo leer la versión local: {e}")
        return "Error"

if __name__ == '__main__':
    # Cambiar al directorio del script
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
    app = QApplication(sys.argv)

    # 🎨 Aplicar estilo claro a todos los cuadros de diálogo (QMessageBox)
    app.setStyleSheet("""
        QMessageBox {
            background-color: #ffffff;
            color: #333333;
            font-size: 14px;
        }
        QMessageBox QLabel {
            color: #333333;
            background-color: #ffffff;
        }
        QMessageBox QPushButton {
            background-color: #4a90e2;
            color: white;
            border-radius: 6px;
            padding: 6px 12px;
            font-weight: bold;
        }
        QMessageBox QPushButton:hover {
            background-color: #357abd;
        }
        QMessageBox QPushButton:pressed {
            background-color: #2d6da3;
        }
    """)

    # 📦 Obtener versión local instalada
    version_local = get_local_commit_sha()
    print(f"📦 Versión instalada: {version_local}")

    # 🚀 Verificar actualizaciones
    auto_update(app)
    
    # 🪪 Iniciar aplicación principal
    window = App()
    window.setWindowTitle(f"RallyCert — v{version_local}")
    window.show()
    
    # Ejecutar aplicación
    exit_code = app.exec()
    sys.exit(exit_code)
