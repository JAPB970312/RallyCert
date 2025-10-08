# main.py
import sys
import os
from PyQt6.QtWidgets import QApplication
from ui import App
from auto_updater import auto_update

if __name__ == '__main__':
    # Cambiar al directorio del script para evitar problemas de rutas
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
    app = QApplication(sys.argv)
    
    # Verificar actualizaciones (solo en modo desarrollo)
    auto_update(app)
    
    # Iniciar aplicación principal
    window = App()
    window.show()
    
    # Ejecutar aplicación
    exit_code = app.exec()
    
    # Asegurarse de que todos los hilos se cierren correctamente
    sys.exit(exit_code)
