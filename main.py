__author__ = "Alexsander Stahnke"
__version__ = "1.5"
__status__ = "Beta"
__email__ = "alexsander.stahnke@grendene.com.br"
__github__ = "https://github.com/QuachRTX"

import tkinter as tk
from tkinter import PhotoImage
import customtkinter as ctk
import os
import sys
import subprocess
import requests
from packaging import version

import checkresultord
import checkpriceadjusts
import autoenri
import produtosdesativar

# Funções para abrir cada funcionalidade
def open_check_result_ord():
    checkresultord.main()

def open_check_price_adjusts():
    checkpriceadjusts.main()

def open_auto_enri():
    autoenri.main()

def open_produtos_desativar():
    produtosdesativar.main()

def get_latest_release_info(repo):
    try:
        response = requests.get(f"https://api.github.com/repos/{repo}/releases/latest")
        response.raise_for_status()
        latest_release = response.json()
        return latest_release['assets'][0]['browser_download_url'], latest_release['tag_name']
    except requests.RequestException as e:
        ctk.CTkMessageBox.show_error(title="Erro", text=f"Erro ao buscar a última release: {e}")
        return None, None

def is_new_version(current_version, latest_version):
    return version.parse(latest_version) > version.parse(current_version)

def check_for_updates():
    repo = "QuachRTX/E-Commerce-Assistant"
    download_url, latest_version = get_latest_release_info(repo)
    if download_url and is_new_version(__version__, latest_version):
        update_label.configure(text="Atualização encontrada! Preparando...")
        launch_update()
    else:
        update_label.configure(text="Nenhuma atualização necessária. Versão mais recente instalada.")

# Função para iniciar o processo de atualização
def launch_update():
    try:
        main_app_directory = os.path.dirname(sys.executable)
        update_executable_path = os.path.join(main_app_directory, "Update.exe")
        subprocess.Popen(update_executable_path)
        root.destroy()  # Fecha o aplicativo atual
    except Exception as e:
        update_label.configure(text=f"Erro ao iniciar o atualizador: {str(e)}")

root = ctk.CTk()
root.title("E-Commerce Assistant by Alexsander Stahnke")
root.geometry("800x600")  # Definindo um tamanho fixo para a janela
ctk.set_appearance_mode("light")  # Define o tema

# Carregar e mostrar a logo
logo_image = PhotoImage(file="_internal/.img/logo2.png")  # Ajuste para o caminho da sua logo
#logo_image = PhotoImage(file=r"C:\Users\Quach\Desktop\VSCode24\Desenvolvimento\E-Commerce Assistent Project\logo2.png")
logo_label = tk.Label(root, image=logo_image)
logo_label.grid(row=0, column=0, columnspan=3, pady=20)

# Definindo a largura dos botões
button_width = 200

btn_planilha1 = ctk.CTkButton(master=root, text="Check Ordenação", command=open_check_result_ord, width=button_width)
btn_planilha1.grid(row=1, column=0, padx=10, pady=20)

btn_planilha2 = ctk.CTkButton(master=root, text="Validador de Preços", command=open_check_price_adjusts, width=button_width)
btn_planilha2.grid(row=1, column=1, padx=10, pady=20)

btn_validar = ctk.CTkButton(master=root, text="Auto Enriquecimento", command=open_auto_enri, width=button_width)
btn_validar.grid(row=1, column=2, padx=10, pady=20)

btn_planilha3 = ctk.CTkButton(master=root, text="Relatório de Produtos - 0020", command=open_check_result_ord, width=button_width, state=ctk.DISABLED)
btn_planilha3.grid(row=2, column=0, padx=10, pady=20)

btn_planilha4 = ctk.CTkButton(master=root, text="Produtos Desativar", command=open_produtos_desativar, width=button_width)
btn_planilha4.grid(row=2, column=1, padx=10, pady=20)

# Botão de atualização com um ícone
update_icon = PhotoImage(file="_internal/.img/update_logo2.png")  # Ajuste para o caminho da sua logo
#update_icon = PhotoImage(file=r"C:\Users\Quach\Desktop\teste download\update_logo2.png")
update_button = ctk.CTkButton(master=root, image=update_icon, text="", width=10, height=10, corner_radius=100, fg_color="lightgray", hover_color="lightgray", command=check_for_updates)
update_button.place(relx=1.0, rely=1.0, anchor="se")

# Label para mostrar informações sobre atualizações
update_label = ctk.CTkLabel(root, text="")
update_label.grid(row=3, column=0, columnspan=3, pady=10)

# Adicionando um texto para sinalizar a versão do aplicativo
versao_label = ctk.CTkLabel(root, text="v1.5  ")
versao_label.place(relx=0.94, rely=1.0, anchor="se")

root.mainloop()
