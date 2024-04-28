__author__ = "Alexsander Stahnke"
__version__ = "1.5"
__status__ = "Beta"
__email__ = "alexsander.stahnke@grendene.com.br"
__github__ = "https://github.com/QuachRTX"

import tkinter as tk
from tkinter import PhotoImage
import customtkinter
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

root = customtkinter.CTk()
root.title("E-Commerce Assistant by Alexsander Stahnke")
root.geometry("800x600")  # Definindo um tamanho fixo para a janela
customtkinter.set_appearance_mode("light")  # Define o tema

# Carregar e mostrar a logo
#logo_image = PhotoImage(file="_internal\.img\logo2.png")
logo_image = PhotoImage(file=r"C:\Users\Quach\Desktop\VSCode24\Desenvolvimento\E-Commerce Assistent Project\logo2.png")  # Ajuste para o caminho da sua logo
logo_label = tk.Label(root, image=logo_image)
logo_label.grid(row=0, column=0, columnspan=3, pady=20)  # Centraliza a logo acima dos botões

# Definindo a largura dos botões para garantir que todos tenham o mesmo tamanho e adicionando-os diretamente ao root usando grid
button_width = 200

btn_planilha1 = customtkinter.CTkButton(master=root, text="Check Ordenação", command=open_check_result_ord, width=button_width)
btn_planilha1.grid(row=1, column=0, padx=10, pady=20)

btn_planilha2 = customtkinter.CTkButton(master=root, text="Validador de Preços", command=open_check_price_adjusts, width=button_width)
btn_planilha2.grid(row=1, column=1, padx=10, pady=20)

btn_validar = customtkinter.CTkButton(master=root, text="Auto Enriquecimento", command=open_auto_enri, width=button_width)
btn_validar.grid(row=1, column=2, padx=10, pady=20)

btn_planilha1 = customtkinter.CTkButton(master=root, text="Relatório de Produtos - 0020", command=open_check_result_ord, width=button_width, state=customtkinter.DISABLED)
btn_planilha1.grid(row=2, column=0, padx=10, pady=20)

btn_planilha2 = customtkinter.CTkButton(master=root, text="Produtos Desativar", command=open_produtos_desativar, width=button_width)
btn_planilha2.grid(row=2, column=1, padx=10, pady=20)

# Adicionando um texto para sinalizar a versão do aplicativo
versao_label = customtkinter.CTkLabel(root, text="v1.5  ")
versao_label.place(relx=1.0, rely=1.0, anchor="se")  # Ancora o label no canto inferior direito


root.mainloop()
