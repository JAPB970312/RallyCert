import os
import shutil
import zipfile
import requests

# Configuración del repositorio corregida
REPO_OWNER = "JAPB970312"
REPO_NAME = "Generador_Constancias"
BRANCH = "main"

def get_remote_commit_sha():
    """Obtiene el SHA del último commit en GitHub."""
    url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/commits?sha={BRANCH}"
    response = requests.get(url)
    if response.status_code == 200:
        return response.json()[0]["sha"]
    else:
        print("Error al obtener el SHA remoto.")
        return None

def get_local_commit_sha(file_path="commit.sha"):
    """Lee el SHA local guardado en un archivo."""
    if not os.path.exists(file_path):
        return None
    with open(file_path, "r") as f:
        return f.read().strip()

def save_local_commit_sha(sha, file_path="commit.sha"):
    """Guarda el SHA local en un archivo."""
    with open(file_path, "w") as f:
        f.write(sha)

def check_for_update():
    """Verifica si hay una nueva versión disponible."""
    remote_sha = get_remote_commit_sha()
    local_sha = get_local_commit_sha()
    return remote_sha != local_sha, remote_sha

def prompt_user_for_update():
    """Solicita autorización al usuario para actualizar."""
    response = input("Hay una nueva actualización disponible. ¿Desea instalarla ahora? (s/n): ")
    return response.lower() == "s"

def download_and_extract_update(target_dir="."):
    """Descarga y extrae la última versión del repositorio."""
    zip_url = f"https://github.com/{REPO_OWNER}/{REPO_NAME}/archive/refs/heads/{BRANCH}.zip"
    zip_path = "update.zip"

    # Descargar el archivo zip
    response = requests.get(zip_url)
    with open(zip_path, "wb") as f:
        f.write(response.content)

    # Extraer el contenido
    with zipfile.ZipFile(zip_path, "r") as zip_ref:
        zip_ref.extractall("update_temp")

    # Copiar archivos al directorio de destino
    extracted_folder = os.path.join("update_temp", f"{REPO_NAME}-{BRANCH}")
    for item in os.listdir(extracted_folder):
        s = os.path.join(extracted_folder, item)
        d = os.path.join(target_dir, item)
        if os.path.isdir(s):
            if os.path.exists(d):
                shutil.rmtree(d)
            shutil.copytree(s, d)
        else:
            shutil.copy2(s, d)

    # Eliminar archivos temporales
    os.remove(zip_path)
    shutil.rmtree("update_temp")

def auto_update(target_dir="."):
    """Proceso completo de verificación y actualización."""
    update_needed, new_sha = check_for_update()
    if update_needed:
        if prompt_user_for_update():
            download_and_extract_update(target_dir)
            save_local_commit_sha(new_sha)
            print("✅ Actualización completada.")
        else:
            print("❌ Actualización cancelada por el usuario.")
    else:
        print("✅ La aplicación ya está actualizada.")