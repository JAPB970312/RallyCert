# auto_updater.py
import os
import sys
import shutil
import zipfile
import requests
import tempfile
from PyQt6.QtWidgets import QApplication, QMessageBox
from PyQt6.QtCore import QThread, pyqtSignal

# Configuración del repositorio GitHub
REPO_OWNER = "JAPB970312"
REPO_NAME = "RallyCert"
BRANCH = "main"

def get_app_dir():
    """Obtiene el directorio de la aplicación"""
    if getattr(sys, 'frozen', False):
        # Ejecutable empaquetado
        return os.path.dirname(sys.executable)
    else:
        # Modo desarrollo
        return os.path.dirname(os.path.abspath(__file__))

def get_commit_file_path():
    """Obtiene la ruta del archivo de commit"""
    app_dir = get_app_dir()
    return os.path.join(app_dir, "commit.sha")

class UpdateThread(QThread):
    """Hilo para manejar la actualización en segundo plano"""
    update_finished = pyqtSignal(bool, str)
    progress_update = pyqtSignal(str)
    
    def __init__(self, target_dir=None):
        super().__init__()
        self.target_dir = target_dir or get_app_dir()
        self.remote_sha = None
        self.local_sha = None
        self._is_running = True
    
    def run(self):
        try:
            self.progress_update.emit("🔍 Verificando actualizaciones...")
            
            # Verificar si hay actualización disponible
            update_needed, remote_sha = self.check_for_update()
            
            if not self._is_running:
                return
                
            if update_needed:
                self.progress_update.emit("📥 Descargando actualización...")
                success = self.download_and_extract_update(remote_sha)
                if success:
                    self.save_local_commit_sha(remote_sha)
                    self.update_finished.emit(True, "✅ Actualización completada correctamente. Reinicie la aplicación.")
                else:
                    self.update_finished.emit(False, "❌ Error durante la descarga o instalación")
            else:
                self.update_finished.emit(True, "✅ Ya tienes la versión más reciente.")
                
        except Exception as e:
            if self._is_running:
                self.update_finished.emit(False, f"❌ Error: {str(e)}")
    
    def stop(self):
        """Detener el hilo de manera segura"""
        self._is_running = False
        self.quit()
        self.wait(5000)

    def get_remote_commit_sha(self):
        """Obtiene el SHA del último commit remoto desde GitHub."""
        if not self._is_running:
            return None
            
        try:
            # Usar la API de GitHub para obtener información del último commit
            url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/commits/{BRANCH}"
            headers = {
                'User-Agent': 'RallyCert-Updater',
                'Accept': 'application/vnd.github.v3+json'
            }
            
            self.progress_update.emit("Conectando con GitHub...")
            response = requests.get(url, headers=headers, timeout=15)
            
            if response.status_code == 200:
                commit_data = response.json()
                sha = commit_data['sha']
                self.progress_update.emit(f"Última versión remota: {sha[:7]}")
                return sha
            else:
                self.progress_update.emit(f"Error API GitHub: {response.status_code}")
                return None
                
        except requests.exceptions.Timeout:
            self.progress_update.emit("⏰ Timeout al conectar con GitHub")
            return None
        except requests.exceptions.ConnectionError:
            self.progress_update.emit("🌐 Error de conexión - verifica tu internet")
            return None
        except Exception as e:
            self.progress_update.emit(f"⚠️ Error al obtener SHA remoto: {e}")
            return None

    def get_local_commit_sha(self):
        """Lee el SHA local guardado en un archivo."""
        commit_file = get_commit_file_path()
        if not os.path.exists(commit_file):
            return None
        try:
            with open(commit_file, "r", encoding="utf-8") as f:
                sha = f.read().strip()
                self.progress_update.emit(f"Versión local: {sha[:7] if sha else 'Ninguna'}")
                return sha
        except Exception as e:
            self.progress_update.emit(f"Error leyendo versión local: {e}")
            return None

    def save_local_commit_sha(self, sha):
        """Guarda el SHA local en un archivo."""
        try:
            commit_file = get_commit_file_path()
            with open(commit_file, "w", encoding="utf-8") as f:
                f.write(sha)
            self.progress_update.emit(f"Versión guardada: {sha[:7]}")
        except Exception as e:
            self.progress_update.emit(f"⚠️ Error al guardar SHA local: {e}")

    def check_for_update(self):
        """Verifica si hay una nueva versión disponible."""
        if not self._is_running:
            return False, None
            
        self.remote_sha = self.get_remote_commit_sha()
        self.local_sha = self.get_local_commit_sha()
        
        # Si no se pudo obtener el SHA remoto, no hay actualización
        if not self.remote_sha:
            return False, None
            
        # Si no hay SHA local, considerar que necesita actualización
        if not self.local_sha:
            return True, self.remote_sha
            
        update_needed = self.remote_sha != self.local_sha
        
        if update_needed:
            self.progress_update.emit(f"🔄 Actualización disponible: {self.local_sha[:7]} → {self.remote_sha[:7]}")
        else:
            self.progress_update.emit("✅ Versiones coinciden")
            
        return update_needed, self.remote_sha

    def download_and_extract_update(self, remote_sha):
        """Descarga y extrae la última versión del repositorio GitHub."""
        if not self._is_running:
            return False
            
        # URL directa al ZIP del repositorio
        zip_url = f"https://github.com/{REPO_OWNER}/{REPO_NAME}/archive/refs/heads/{BRANCH}.zip"
        
        # Crear directorios temporales
        temp_dir = tempfile.mkdtemp(prefix="rallycert_update_")
        zip_path = os.path.join(temp_dir, "update.zip")
        extract_dir = os.path.join(temp_dir, "extracted")

        try:
            if not self._is_running:
                return False
                
            self.progress_update.emit("📥 Descargando actualización desde GitHub...")
            
            # Descargar con manejo de errores
            try:
                headers = {'User-Agent': 'RallyCert-Updater'}
                response = requests.get(zip_url, headers=headers, timeout=60, stream=True)
                response.raise_for_status()
                
                # Guardar archivo ZIP
                with open(zip_path, "wb") as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        if not self._is_running:
                            return False
                        if chunk:
                            f.write(chunk)
                            
                self.progress_update.emit(f"✅ Descarga completada ({os.path.getsize(zip_path) // 1024} KB)")
                
            except requests.exceptions.Timeout:
                self.progress_update.emit("⏰ Timeout al descargar actualización")
                return False
            except requests.exceptions.ConnectionError:
                self.progress_update.emit("🌐 Error de conexión durante descarga")
                return False
            except Exception as e:
                self.progress_update.emit(f"❌ Error en descarga: {e}")
                return False

            if not self._is_running:
                return False

            self.progress_update.emit("📦 Extrayendo archivos...")
            
            # Extraer archivos
            with zipfile.ZipFile(zip_path, "r") as zip_ref:
                zip_ref.extractall(extract_dir)

            # La estructura extraída es: extracted/REPO_NAME-BRANCH/*
            extracted_folder = os.path.join(extract_dir, f"{REPO_NAME}-{BRANCH}")
            
            if not os.path.exists(extracted_folder):
                self.progress_update.emit("❌ Estructura de archivos incorrecta")
                return False

            self.progress_update.emit("🔄 Copiando archivos actualizados...")
            
            # Lista de archivos/carpetas a excluir de la actualización
            exclude_files = ['commit.sha', 'config.ini', 'user_settings.json', 'keys']
            exclude_folders = ['__pycache__', '.git', 'output', 'temp', 'venv', '.github']
            
            # Función para copiar archivos de manera segura
            def copy_files(src, dst):
                for item in os.listdir(src):
                    if not self._is_running:
                        return False
                        
                    src_path = os.path.join(src, item)
                    dst_path = os.path.join(dst, item)
                    
                    # Saltar archivos/carpetas excluidos
                    if item in exclude_files or item in exclude_folders:
                        continue
                    
                    try:
                        if os.path.isdir(src_path):
                            if not os.path.exists(dst_path):
                                os.makedirs(dst_path)
                            if not copy_files(src_path, dst_path):
                                return False
                        else:
                            # Crear directorio padre si no existe
                            os.makedirs(os.path.dirname(dst_path), exist_ok=True)
                            shutil.copy2(src_path, dst_path)
                            self.progress_update.emit(f"  📄 {item}")
                    except Exception as e:
                        self.progress_update.emit(f"⚠️ Error copiando {item}: {e}")
                        # Continuar con otros archivos
                        continue
                return True

            # Copiar archivos
            if not copy_files(extracted_folder, self.target_dir):
                return False

            self.progress_update.emit("✅ Archivos copiados correctamente")
            return True

        except Exception as e:
            self.progress_update.emit(f"❌ Error durante la actualización: {e}")
            return False
        finally:
            # Limpieza de archivos temporales
            self.cleanup_temp_files(temp_dir)

    def cleanup_temp_files(self, temp_dir):
        """Limpia archivos temporales de manera segura"""
        try:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
                self.progress_update.emit("🧹 Archivos temporales eliminados")
        except Exception as e:
            self.progress_update.emit(f"⚠️ Error limpiando temporales: {e}")


def prompt_user_for_update(remote_sha, local_sha):
    """Muestra un cuadro de diálogo para solicitar autorización."""
    msg = QMessageBox()
    msg.setWindowTitle("Actualización disponible")
    msg.setIcon(QMessageBox.Icon.Question)
    
    message_text = f"""
Se ha detectado una nueva versión de RallyCert.

Versión actual: {local_sha[:7] if local_sha else 'Ninguna'}
Nueva versión: {remote_sha[:7]}

¿Desea descargar e instalar la actualización ahora?

• La aplicación se cerrará automáticamente
• Se reiniciará después de la actualización
• Sus configuraciones se mantendrán
"""
    
    msg.setText(message_text)
    msg.setStandardButtons(
        QMessageBox.StandardButton.Yes | 
        QMessageBox.StandardButton.No
    )
    msg.setDefaultButton(QMessageBox.StandardButton.Yes)
    
    # Aplicar estilo consistente
    msg.setStyleSheet("""
        QMessageBox {
            background-color: #f8f9fa;
            color: #333333;
        }
        QMessageBox QLabel {
            color: #333333;
        }
        QMessageBox QPushButton {
            background-color: #4a90e2;
            color: white;
            border-radius: 6px;
            padding: 8px 16px;
            font-weight: bold;
            min-width: 80px;
        }
        QMessageBox QPushButton:hover {
            background-color: #357abd;
        }
    """)
    
    respuesta = msg.exec()
    return respuesta == QMessageBox.StandardButton.Yes


# Variable global para mantener referencia al hilo
_update_thread = None

def auto_update(app=None, target_dir=None):
    """Ejecuta la verificación y actualización completa con PyQt6."""
    global _update_thread
    
    # Si estamos en un ejecutable empaquetado, verificar actualizaciones
    if getattr(sys, 'frozen', False):
        print("📦 Ejecutable empaquetado - Verificando actualizaciones...")
    else:
        print("🔧 Modo desarrollo - Verificando actualizaciones...")
    
    try:
        # Verificación rápida sin hilo primero
        print("🔍 Verificando actualizaciones...")
        
        # Crear instancia temporal para verificación
        temp_updater = UpdateThread(target_dir)
        remote_sha = temp_updater.get_remote_commit_sha()
        local_sha = temp_updater.get_local_commit_sha()
        
        if not remote_sha:
            print("⚠️ No se pudo conectar con GitHub para verificar actualizaciones")
            return
        
        update_needed = remote_sha != local_sha
        
        if update_needed:
            print(f"🔄 Actualización disponible: {local_sha[:7] if local_sha else 'Ninguna'} → {remote_sha[:7]}")
            
            # Preguntar al usuario
            if prompt_user_for_update(remote_sha, local_sha):
                # Crear y ejecutar hilo de actualización
                _update_thread = UpdateThread(target_dir)
                
                def on_update_finished(success, message):
                    global _update_thread
                    print(message)
                    
                    if success:
                        # Mostrar mensaje de éxito
                        msg = QMessageBox()
                        msg.setWindowTitle("Actualización completada")
                        msg.setIcon(QMessageBox.Icon.Information)
                        msg.setText(f"{message}\n\nLa aplicación se cerrará ahora.")
                        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
                        msg.exec()
                        
                        # Cerrar la aplicación para aplicar cambios
                        if app:
                            app.quit()
                    else:
                        # Mostrar error
                        msg = QMessageBox()
                        msg.setWindowTitle("Error de actualización")
                        msg.setIcon(QMessageBox.Icon.Critical)
                        msg.setText(message)
                        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
                        msg.exec()
                    
                    # Limpiar referencia al hilo
                    _update_thread = None
                
                def on_progress_update(message):
                    print(f"🔄 {message}")
                
                _update_thread.update_finished.connect(on_update_finished)
                _update_thread.progress_update.connect(on_progress_update)
                _update_thread.start()
            else:
                print("❌ Actualización cancelada por el usuario.")
        else:
            print("✅ RallyCert ya está actualizado.")
            
    except Exception as e:
        print(f"❌ Error en el proceso de actualización: {e}")


def get_local_commit_sha():
    """Obtiene la versión local instalada"""
    try:
        commit_file = get_commit_file_path()
        if os.path.exists(commit_file):
            with open(commit_file, "r", encoding="utf-8") as f:
                sha = f.read().strip()
                return sha[:7] if sha else "Desconocida"
        return "Desconocida"
    except Exception as e:
        print(f"⚠️ No se pudo leer la versión local: {e}")
        return "Error"


if __name__ == "__main__":
    # Para testing
    app = QApplication(sys.argv)
    auto_update(app)
    sys.exit(app.exec())
