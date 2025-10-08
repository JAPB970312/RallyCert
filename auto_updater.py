# auto_updater.py
import os
import sys
import shutil
import zipfile
import requests
from PyQt6.QtWidgets import QApplication, QMessageBox
from PyQt6.QtCore import QThread, pyqtSignal

# Configuración del repositorio GitHub
REPO_OWNER = "JAPB970312"
REPO_NAME = "RallyCert"
BRANCH = "main"
COMMIT_FILE = "commit.sha"


class UpdateThread(QThread):
    """Hilo para manejar la actualización en segundo plano"""
    update_finished = pyqtSignal(bool, str)
    progress_update = pyqtSignal(str)
    
    def __init__(self, target_dir="."):
        super().__init__()
        self.target_dir = target_dir
        self.remote_sha = None
        self.local_sha = None
        self._is_running = True
    
    def run(self):
        try:
            self.progress_update.emit("🔍 Verificando actualizaciones...")
            
            # Verificar si hay actualización disponible
            update_needed = self.check_for_update()
            
            if not self._is_running:
                return
                
            if update_needed:
                self.progress_update.emit("📥 Descargando actualización...")
                success = self.download_and_extract_update()
                if success:
                    self.save_local_commit_sha(self.remote_sha)
                    self.update_finished.emit(True, "✅ Actualización completada correctamente.")
                else:
                    self.update_finished.emit(False, "❌ Error durante la actualización")
            else:
                self.update_finished.emit(True, "✅ Ya tienes la versión más reciente.")
                
        except Exception as e:
            if self._is_running:
                self.update_finished.emit(False, f"❌ Error: {str(e)}")
    
    def stop(self):
        """Detener el hilo de manera segura"""
        self._is_running = False
        self.quit()
        self.wait(5000)  # Esperar máximo 5 segundos

    def get_remote_commit_sha(self):
        """Obtiene el SHA del último commit remoto desde GitHub."""
        if not self._is_running:
            return None
            
        url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/commits?sha={BRANCH}"
        try:
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            return response.json()[0]["sha"]
        except requests.exceptions.Timeout:
            print("⏰ Timeout al conectar con GitHub")
            return None
        except requests.exceptions.ConnectionError:
            print("🌐 Error de conexión - verifica tu internet")
            return None
        except Exception as e:
            print(f"⚠️ Error al obtener el SHA remoto: {e}")
            return None

    def get_local_commit_sha(self, file_path=COMMIT_FILE):
        """Lee el SHA local guardado en un archivo."""
        if not os.path.exists(file_path):
            return None
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                return f.read().strip()
        except:
            return None

    def save_local_commit_sha(self, sha, file_path=COMMIT_FILE):
        """Guarda el SHA local en un archivo."""
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(sha)
        except Exception as e:
            print(f"⚠️ Error al guardar SHA local: {e}")

    def check_for_update(self):
        """Verifica si hay una nueva versión disponible."""
        if not self._is_running:
            return False
            
        self.remote_sha = self.get_remote_commit_sha()
        self.local_sha = self.get_local_commit_sha()
        
        # Si no se pudo obtener el SHA remoto, no hay actualización
        if not self.remote_sha:
            return False
            
        update_needed = self.remote_sha != self.local_sha
        return update_needed

    def download_and_extract_update(self):
        """Descarga y extrae la última versión del repositorio GitHub."""
        if not self._is_running:
            return False
            
        zip_url = f"https://github.com/{REPO_OWNER}/{REPO_NAME}/archive/refs/heads/{BRANCH}.zip"
        zip_path = "update.zip"
        temp_dir = "update_temp"

        try:
            if not self._is_running:
                return False
                
            self.progress_update.emit("📥 Descargando actualización...")
            
            # Descargar con manejo de errores de conexión
            try:
                response = requests.get(zip_url, timeout=30)
                response.raise_for_status()
            except requests.exceptions.Timeout:
                self.progress_update.emit("⏰ Timeout al descargar actualización")
                return False
            except requests.exceptions.ConnectionError:
                self.progress_update.emit("🌐 Error de conexión durante descarga")
                return False

            with open(zip_path, "wb") as f:
                f.write(response.content)

            if not self._is_running:
                self.cleanup_temp_files(zip_path, temp_dir)
                return False

            self.progress_update.emit("📦 Extrayendo actualización...")
            with zipfile.ZipFile(zip_path, "r") as zip_ref:
                zip_ref.extractall(temp_dir)

            extracted_folder = os.path.join(temp_dir, f"{REPO_NAME}-{BRANCH}")

            # Lista de archivos/carpetas a excluir de la actualización
            exclude_files = ['commit.sha', 'config.ini', 'user_settings.json']
            exclude_folders = ['__pycache__', '.git', 'output', 'temp', 'venv']

            self.progress_update.emit("🔄 Aplicando actualización...")
            
            for item in os.listdir(extracted_folder):
                if not self._is_running:
                    self.cleanup_temp_files(zip_path, temp_dir)
                    return False
                    
                if item in exclude_folders:
                    continue
                    
                src = os.path.join(extracted_folder, item)
                dst = os.path.join(self.target_dir, item)
                
                # Si es un archivo excluido, saltar
                if os.path.isfile(src) and item in exclude_files:
                    continue
                
                try:
                    if os.path.isdir(src):
                        if os.path.exists(dst):
                            shutil.rmtree(dst)
                        shutil.copytree(src, dst)
                    else:
                        shutil.copy2(src, dst)
                except Exception as e:
                    print(f"⚠️ Error al copiar {item}: {e}")

            self.progress_update.emit("🧹 Limpiando archivos temporales...")
            self.cleanup_temp_files(zip_path, temp_dir)

            return True

        except Exception as e:
            print(f"❌ Error durante la actualización: {e}")
            self.cleanup_temp_files(zip_path, temp_dir)
            return False

    def cleanup_temp_files(self, zip_path, temp_dir):
        """Limpia archivos temporales"""
        try:
            if os.path.exists(zip_path):
                os.remove(zip_path)
        except:
            pass
            
        try:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
        except:
            pass


def prompt_user_for_update(remote_sha, local_sha):
    """Muestra un cuadro de diálogo para solicitar autorización."""
    msg = (
        f"Se ha detectado una nueva versión en GitHub.\n\n"
        f"Versión instalada:\n{local_sha or 'Ninguna'}\n"
        f"Nueva versión disponible:\n{remote_sha}\n\n"
        f"¿Desea descargar e instalar la actualización?"
    )

    respuesta = QMessageBox.question(
        None,
        "Actualización disponible",
        msg,
        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
    )
    
    return respuesta == QMessageBox.StandardButton.Yes


# Variable global para mantener referencia al hilo
_update_thread = None

def auto_update(app=None, target_dir="."):
    """Ejecuta la verificación y actualización completa con PyQt6."""
    global _update_thread
    
    # Si estamos en un ejecutable empaquetado, no verificar actualizaciones automáticamente
    if getattr(sys, 'frozen', False):
        print("📦 Ejecutable empaquetado - omitiendo actualizaciones automáticas")
        return
    
    try:
        # Verificar actualización de manera síncrona primero
        print("🔍 Verificando actualizaciones...")
        
        # Verificación rápida sin hilo
        remote_sha = None
        local_sha = None
        
        try:
            url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/commits?sha={BRANCH}"
            response = requests.get(url, timeout=8)  # Timeout reducido
            response.raise_for_status()
            remote_sha = response.json()[0]["sha"]
            
            if os.path.exists(COMMIT_FILE):
                with open(COMMIT_FILE, "r", encoding="utf-8") as f:
                    local_sha = f.read().strip()
                    
            print(f"✅ SHA remoto: {remote_sha}")
            print(f"✅ SHA local: {local_sha}")
            
        except requests.exceptions.Timeout:
            print("⏰ Timeout: No se pudo conectar con GitHub (verifica tu conexión)")
            return
        except requests.exceptions.ConnectionError:
            print("🌐 Error de conexión: Verifica tu acceso a internet")
            return
        except Exception as e:
            print(f"⚠️ No se pudo verificar actualizaciones: {e}")
            return
        
        if not remote_sha:
            print("⚠️ No se pudo obtener información de actualizaciones")
            return
        
        update_needed = remote_sha != local_sha
        
        if update_needed:
            print("🔄 Actualización disponible")
            if prompt_user_for_update(remote_sha, local_sha):
                # Crear y ejecutar hilo de actualización
                _update_thread = UpdateThread(target_dir)
                
                def on_update_finished(success, message):
                    global _update_thread
                    print(message)
                    if success:
                        QMessageBox.information(
                            None, 
                            "Actualización", 
                            f"{message}\n\nReinicie la aplicación para aplicar los cambios."
                        )
                    else:
                        QMessageBox.critical(None, "Error de actualización", message)
                    
                    # Limpiar referencia al hilo
                    _update_thread = None
                
                def on_progress_update(message):
                    print(message)
                
                _update_thread.update_finished.connect(on_update_finished)
                _update_thread.progress_update.connect(on_progress_update)
                _update_thread.start()
            else:
                print("❌ Actualización cancelada por el usuario.")
        else:
            print("✅ RallyCert ya está actualizado.")
            
    except Exception as e:
        print(f"❌ Error en el proceso de actualización: {e}")


if __name__ == "__main__":
    # Para testing
    app = QApplication(sys.argv)
    auto_update(app)
    sys.exit(app.exec())
if __name__ == "__main__":
    auto_update()

