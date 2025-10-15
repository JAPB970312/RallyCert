# auto_updater.py
import os
import sys
import shutil
import zipfile
import requests
import tempfile
import subprocess
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

def get_user_data_dir():
    """Obtiene el directorio de datos de usuario (AppData)"""
    app_data_dir = os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "RallyCert")
    os.makedirs(app_data_dir, exist_ok=True)
    return app_data_dir

def get_commit_file_path():
    """Obtiene la ruta del archivo de commit en AppData"""
    return os.path.join(get_user_data_dir(), "commit.sha")

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

def is_admin_installation():
    """Verifica si la aplicación está instalada en una ubicación que requiere admin"""
    app_dir = get_app_dir()
    protected_paths = [
        os.environ.get('PROGRAMFILES', 'C:\\Program Files'),
        os.environ.get('PROGRAMFILES(X86)', 'C:\\Program Files (x86)'),
        os.environ.get('PROGRAMW6432', 'C:\\Program Files')
    ]
    
    for path in protected_paths:
        if path and app_dir.startswith(path):
            return True
    return False

def run_as_admin(command):
    """Ejecuta un comando con permisos de administrador"""
    try:
        import ctypes
        if ctypes.windll.shell32.ShellExecuteW(None, "runas", "cmd.exe", f'/c "{command}"', None, 1) > 32:
            return True
        return False
    except Exception:
        return False

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
        self.is_admin_install = is_admin_installation()
    
    def run(self):
        try:
            self.progress_update.emit("🔍 Verificando actualizaciones...")
            
            # Verificar si hay actualización disponible
            update_needed, remote_sha = self.check_for_update()
            
            if not self._is_running:
                return
                
            if update_needed:
                if self.is_admin_install:
                    self.progress_update.emit("⚠️ Instalación detectada en ubicación protegida...")
                    
                self.progress_update.emit("📥 Descargando actualización...")
                success = self.download_and_extract_update(remote_sha)
                if success:
                    self.save_local_commit_sha(remote_sha)
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
        self.wait(5000)

    def get_remote_commit_sha(self):
        """Obtiene el SHA del último commit remoto desde GitHub."""
        if not self._is_running:
            return None
            
        try:
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
        
        if not self.remote_sha:
            return False, None
            
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
            
        zip_url = f"https://github.com/{REPO_OWNER}/{REPO_NAME}/archive/refs/heads/{BRANCH}.zip"
        temp_dir = tempfile.mkdtemp(prefix="rallycert_update_")
        zip_path = os.path.join(temp_dir, "update.zip")
        extract_dir = os.path.join(temp_dir, "extracted")

        try:
            if not self._is_running:
                return False
                
            self.progress_update.emit("📥 Descargando actualización desde GitHub...")
            
            # Descargar archivo
            try:
                headers = {'User-Agent': 'RallyCert-Updater'}
                response = requests.get(zip_url, headers=headers, timeout=60, stream=True)
                response.raise_for_status()
                
                with open(zip_path, "wb") as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        if not self._is_running:
                            return False
                        if chunk:
                            f.write(chunk)
                            
                self.progress_update.emit(f"✅ Descarga completada ({os.path.getsize(zip_path) // 1024} KB)")
                
            except Exception as e:
                self.progress_update.emit(f"❌ Error en descarga: {e}")
                return False

            if not self._is_running:
                return False

            self.progress_update.emit("📦 Extrayendo archivos...")
            
            # Extraer archivos
            with zipfile.ZipFile(zip_path, "r") as zip_ref:
                zip_ref.extractall(extract_dir)

            extracted_folder = os.path.join(extract_dir, f"{REPO_NAME}-{BRANCH}")
            
            if not os.path.exists(extracted_folder):
                self.progress_update.emit("❌ Estructura de archivos incorrecta")
                return False

            self.progress_update.emit("🔄 Aplicando actualización...")
            
            # Para instalaciones con admin, usar método especial
            if self.is_admin_install:
                return self.update_with_admin_privileges(extracted_folder, temp_dir)
            else:
                return self.update_normal(extracted_folder, temp_dir)

        except Exception as e:
            self.progress_update.emit(f"❌ Error durante la actualización: {e}")
            return False

    def update_normal(self, extracted_folder, temp_dir):
        """Actualización normal para ubicaciones sin protección"""
        try:
            exclude_files = ['commit.sha', 'config.ini', 'user_settings.json', 'keys']
            exclude_folders = ['__pycache__', '.git', 'output', 'temp', 'venv', '.github']
            
            def copy_files(src, dst):
                for item in os.listdir(src):
                    if not self._is_running:
                        return False
                        
                    src_path = os.path.join(src, item)
                    dst_path = os.path.join(dst, item)
                    
                    if item in exclude_files or item in exclude_folders:
                        continue
                    
                    try:
                        if os.path.isdir(src_path):
                            if not os.path.exists(dst_path):
                                os.makedirs(dst_path)
                            if not copy_files(src_path, dst_path):
                                return False
                        else:
                            os.makedirs(os.path.dirname(dst_path), exist_ok=True)
                            shutil.copy2(src_path, dst_path)
                            self.progress_update.emit(f"  📄 {item}")
                    except Exception as e:
                        self.progress_update.emit(f"⚠️ Error copiando {item}: {e}")
                        continue
                return True

            if not copy_files(extracted_folder, self.target_dir):
                return False

            self.progress_update.emit("✅ Archivos actualizados correctamente")
            return True

        except Exception as e:
            self.progress_update.emit(f"❌ Error actualizando archivos: {e}")
            return False
        finally:
            self.cleanup_temp_files(temp_dir)

    def update_with_admin_privileges(self, extracted_folder, temp_dir):
        """Actualización para instalaciones que requieren admin"""
        try:
            # Crear script batch para actualización
            batch_script = os.path.join(temp_dir, "update.bat")
            log_file = os.path.join(temp_dir, "update.log")
            
            with open(batch_script, "w") as f:
                f.write(f"""@echo off
echo Iniciando actualización de RallyCert...
echo %date% %time% > "{log_file}"

REM Copiar archivos actualizados
xcopy "{extracted_folder}\\*" "{self.target_dir}" /E /Y /I /Q
if %errorlevel% neq 0 (
    echo ERROR: No se pudieron copiar los archivos >> "{log_file}"
    exit /b 1
)

echo Archivos actualizados correctamente >> "{log_file}"
echo Actualización completada exitosamente.
timeout /t 2 /nobreak >nul
""")
            
            # Ejecutar con privilegios de administrador
            self.progress_update.emit("🛡️ Solicitando permisos de administrador...")
            
            import ctypes
            result = ctypes.windll.shell32.ShellExecuteW(
                None, "runas", "cmd.exe", f'/c "{batch_script}"', None, 0
            )
            
            if result > 32:
                self.progress_update.emit("✅ Actualización iniciada con permisos de administrador")
                return True
            else:
                self.progress_update.emit("❌ No se pudieron obtener permisos de administrador")
                return False
                
        except Exception as e:
            self.progress_update.emit(f"❌ Error en actualización con admin: {e}")
            return False

    def cleanup_temp_files(self, temp_dir):
        """Limpia archivos temporales"""
        try:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
        except Exception:
            pass


def prompt_user_for_update(remote_sha, local_sha):
    """Muestra diálogo para confirmar actualización"""
    msg = QMessageBox()
    msg.setWindowTitle("Actualización disponible")
    msg.setIcon(QMessageBox.Icon.Question)
    
    message = f"""
Se ha detectado una nueva versión de RallyCert.

Versión actual: {local_sha[:7] if local_sha else 'Ninguna'}
Nueva versión: {remote_sha[:7]}

¿Desea descargar e instalar la actualización ahora?
"""
    
    if is_admin_installation():
        message += "\n\n⚠️ Se solicitarán permisos de administrador para completar la actualización."
    
    msg.setText(message)
    msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
    msg.setDefaultButton(QMessageBox.StandardButton.Yes)
    
    return msg.exec() == QMessageBox.StandardButton.Yes


# Variable global para el hilo de actualización
_update_thread = None

def auto_update(app=None):
    """Función principal de actualización automática"""
    global _update_thread
    
    try:
        # Verificación rápida
        temp_updater = UpdateThread()
        remote_sha = temp_updater.get_remote_commit_sha()
        local_sha = temp_updater.get_local_commit_sha()
        
        if not remote_sha:
            return
        
        update_needed = remote_sha != local_sha
        
        if update_needed and prompt_user_for_update(remote_sha, local_sha):
            _update_thread = UpdateThread()
            
            def on_finished(success, message):
                global _update_thread
                print(message)
                
                if success:
                    msg = QMessageBox()
                    msg.setWindowTitle("Actualización")
                    msg.setIcon(QMessageBox.Icon.Information)
                    
                    if is_admin_installation():
                        msg.setText("✅ Actualización iniciada. Se requieren permisos de administrador.\n\nLa aplicación se cerrará.")
                    else:
                        msg.setText("✅ Actualización completada.\n\nLa aplicación se cerrará.")
                    
                    msg.exec()
                    
                    if app:
                        app.quit()
                else:
                    msg = QMessageBox()
                    msg.setWindowTitle("Error")
                    msg.setIcon(QMessageBox.Icon.Critical)
                    msg.setText(message)
                    msg.exec()
                
                _update_thread = None
            
            def on_progress(message):
                print(f"🔄 {message}")
            
            _update_thread.update_finished.connect(on_finished)
            _update_thread.progress_update.connect(on_progress)
            _update_thread.start()
            
    except Exception as e:
        print(f"❌ Error en actualización automática: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    auto_update(app)
    sys.exit(app.exec())
    # Para testing
    app = QApplication(sys.argv)
    auto_update(app)
    sys.exit(app.exec())
