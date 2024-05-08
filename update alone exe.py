import customtkinter as ctk
import os
import sys
import requests
import zipfile
import shutil

def download_file(url, save_path, label):
    label.configure(text="Baixando a atualização...")
    try:
        response = requests.get(url, stream=True)
        if response.status_code == 200:
            with open(save_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            label.configure(text="Download concluído. Extraindo...")
            unzip_file(save_path, os.path.dirname(save_path), label)
        else:
            label.configure(text=f"Erro no download: Status {response.status_code}")
    except Exception as e:
        label.configure(text=f"Erro no download: {str(e)}")

def unzip_file(zip_path, extract_to, label):
    label.configure(text="Extraindo o arquivo...")
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_to)
            extracted_folder_path = os.path.join(extract_to, 'ecommerce_assistant')
            for item in os.listdir(extracted_folder_path):
                source_item = os.path.join(extracted_folder_path, item)
                destination_item = os.path.join(extract_to, item)
                if os.path.isdir(destination_item):
                    shutil.rmtree(destination_item)
                elif os.path.exists(destination_item):
                    os.remove(destination_item)
                shutil.move(source_item, extract_to)
            os.rmdir(extracted_folder_path)
        label.configure(text="Atualização instalada. Por favor, reinicie o aplicativo.")
        os.remove(zip_path)
    except Exception as e:
        label.configure(text=f"Erro ao instalar: {str(e)}")

def get_latest_release_url(repo):
    try:
        response = requests.get(url=f"https://api.github.com/repos/{repo}/releases/latest")
        response.raise_for_status()
        latest_release = response.json()
        return latest_release['assets'][0]['browser_download_url']
    except requests.RequestException as e:
        ctk.CTkMessageBox.show_error(title="Erro", text=f"Erro ao buscar a última release: {e}")
        return None

def check_updates(label, repo):
    label.configure(text="Verificando atualizações, aguarde...")
    download_url = get_latest_release_url(repo)
    if download_url:
        label.configure(text="Atualização encontrada! Baixando...")
        executable_path = os.path.dirname(sys.executable)
        local_save_path = os.path.join(executable_path, "ecommerce_assistant_update.zip")
        download_file(download_url, local_save_path, label)
    else:
        label.configure(text="Nenhuma atualização disponível.")

def create_update_interface():
    status_label = ctk.CTkLabel(root, text="Inicializando...", wraplength=300)
    status_label.pack(pady=20)
    root.after(100, lambda: check_updates(status_label, "QuachRTX/E-Commerce-Assistant"))

if __name__ == "__main__":
    ctk.set_appearance_mode("light")  # Define o tema antes de iniciar a janela
    root = ctk.CTk()
    root.title("Updater E-Commerce Assistant")
    root.geometry("400x200")  # Definindo um tamanho fixo para a janela
    create_update_interface()
    root.mainloop()
