import os
import sys
import shutil
import zipfile
import requests
import tkinter as tk
from tkinter import messagebox

# Configuración del repositorio GitHub
REPO_OWNER = "JAPB970312"
REPO_NAME = "RallyCert"
BRANCH = "main"
COMMIT_FILE = "commit.sha"


# =============================
# 🔹 Funciones de utilidad
# =============================

def get_remote_commit_sha():
    """Obtiene el SHA del último commit remoto desde GitHub."""
    url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/commits?sha={BRANCH}"
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        return response.json()[0]["sha"]
    except Exception as e:
        print(f"⚠️ Error al obtener el SHA remoto: {e}")
        return None


def get_local_commit_sha(file_path=COMMIT_FILE):
    """Lee el SHA local guardado en un archivo."""
    if not os.path.exists(file_path):
        return None
    with open(file_path, "r", encoding="utf-8") as f:
        return f.read().strip()


def save_local_commit_sha(sha, file_path=COMMIT_FILE):
    """Guarda el SHA local en un archivo."""
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(sha)


def check_for_update():
    """Verifica si hay una nueva versión disponible."""
    remote_sha = get_remote_commit_sha()
    local_sha = get_local_commit_sha()
    update_needed = remote_sha and (remote_sha != local_sha)
    return update_needed, remote_sha


# =============================
# 🔹 Interacción con el usuario
# =============================

def prompt_user_for_update(remote_sha, local_sha):
    """Muestra un cuadro gráfico para solicitar autorización."""
    root = tk.Tk()
    root.withdraw()  # Oculta la ventana principal

    msg = (
        f"Se ha detectado una nueva versión en GitHub.\n\n"
        f"Versión instalada:\n{local_sha or 'Ninguna'}\n"
        f"Nueva versión disponible:\n{remote_sha}\n\n"
        f"¿Desea descargar e instalar la actualización?"
    )

    respuesta = messagebox.askyesno("Actualización disponible", msg)
    root.destroy()
    return respuesta


# =============================
# 🔹 Proceso de descarga
# =============================

def download_and_extract_update(target_dir="."):
    """Descarga y extrae la última versión del repositorio GitHub."""
    zip_url = f"https://github.com/{REPO_OWNER}/{REPO_NAME}/archive/refs/heads/{BRANCH}.zip"
    zip_path = "update.zip"
    temp_dir = "update_temp"

    try:
        print("📥 Descargando actualización desde GitHub...")
        response = requests.get(zip_url, timeout=30)
        response.raise_for_status()

        with open(zip_path, "wb") as f:
            f.write(response.content)

        print("📦 Extrayendo actualización...")
        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall(temp_dir)

        extracted_folder = os.path.join(temp_dir, f"{REPO_NAME}-{BRANCH}")

        for item in os.listdir(extracted_folder):
            src = os.path.join(extracted_folder, item)
            dst = os.path.join(target_dir, item)
            if os.path.isdir(src):
                if os.path.exists(dst):
                    shutil.rmtree(dst)
                shutil.copytree(src, dst)
            else:
                shutil.copy2(src, dst)

        print("🧹 Limpiando archivos temporales...")
        os.remove(zip_path)
        shutil.rmtree(temp_dir)

        print("✅ Actualización completada correctamente.")
        return True

    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Error en actualización", f"Ocurrió un error al actualizar:\n{e}")
        root.destroy()
        print(f"❌ Error durante la actualización: {e}")
        return False


# =============================
# 🔹 Proceso principal
# =============================

def auto_update(target_dir="."):
    """Ejecuta la verificación y actualización completa."""
    update_needed, new_sha = check_for_update()

    if not new_sha:
        print("⚠️ No se pudo verificar la versión remota (sin conexión o API caída).")
        return

    if update_needed:
        local_sha = get_local_commit_sha()
        if prompt_user_for_update(new_sha, local_sha):
            success = download_and_extract_update(target_dir)
            if success:
                save_local_commit_sha(new_sha)
                root = tk.Tk()
                root.withdraw()
                messagebox.showinfo("Actualización", "✅ RallyCert se ha actualizado correctamente.\nReinicie la aplicación.")
                root.destroy()
                os._exit(0)  # Cierra la app después de actualizar
        else:
            print("❌ Actualización cancelada por el usuario.")
    else:
        print("✅ RallyCert ya está actualizado.")


if __name__ == "__main__":
    auto_update()

