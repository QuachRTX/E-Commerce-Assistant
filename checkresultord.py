__author__ = "Alexsander Stahnke"
__version__ = "1.4"
__status__ = "Stable"
__email__ = "alexsander.stahnke@grendene.com.br"
__github__ = "https://github.com/QuachRTX"

from tkinter import filedialog
from tkinter import messagebox
from bs4 import BeautifulSoup
import pandas as pd
import customtkinter as ctk

def extrair_produtos_e_codigos(arquivo_html):
    with open(arquivo_html, "r", encoding="utf-8") as file:
        html_content = file.read()
    
    soup = BeautifulSoup(html_content, 'html.parser')
    produtos_e_codigos = []
    
    for link in soup.find_all('a', href=True):
        href = link['href']
        if "/p/" in href and "?" in href:
            path_parts = href.split("/p/")[1].split("?")[0]
            product_name, product_code = path_parts.split("/") if "/" in path_parts else (path_parts, "")
            if product_name and product_code:
                produtos_e_codigos.append((product_name.replace("-", " ").title(), product_code))
    return produtos_e_codigos

def remover_duplicatas_preservando_ordem(lista):
    seen = set()
    seen_add = seen.add
    return [x for x in lista if not (x in seen or seen_add(x))]

def exportar_para_excel(produtos_e_codigos, nome_arquivo):
    df = pd.DataFrame(produtos_e_codigos, columns=['Produto', 'Código'])
    df.to_excel(nome_arquivo, index=False)

def iniciar_processo_checkord():
    arquivo_html = filedialog.askopenfilename(initialdir="/", title="Selecione o arquivo HTML",
                                               filetypes=(("Arquivos HTML", "*.html"), ("todos os arquivos", "*.*")))
    if arquivo_html:
        produtos_e_codigos = extrair_produtos_e_codigos(arquivo_html)
        produtos_e_codigos_unicos = remover_duplicatas_preservando_ordem(produtos_e_codigos)
        nome_arquivo_excel = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                          filetypes=[("Excel files", "*.xlsx")])
        if nome_arquivo_excel:
            exportar_para_excel(produtos_e_codigos_unicos, nome_arquivo_excel)
            # Correção para exibir o pop-up de confirmação
            messagebox.showinfo("Sucesso", f"Arquivo '{nome_arquivo_excel}' foi salvo com sucesso.")
        else:
            messagebox.showinfo("Cancelado", "Operação cancelada, local para salvar o arquivo não foi selecionado.")
    else:
        messagebox.showinfo("Cancelado", "Operação cancelada, arquivo HTML não foi selecionado.")


def main():
    # Configuração da janela principal
    root = ctk.CTk()
    root.title("Check Result Ord")
    root.geometry("400x300")

    # Label de instrução para o usuário
    instrucao_label = ctk.CTkLabel(master=root, text="Por favor, selecione um arquivo HTML", height=10)
    instrucao_label.grid(row=0, column=0, columnspan=3, pady=(150, 0))  # Ajuste para centralizar e posicionar corretamente

    # Botão para iniciar o processo de seleção do arquivo e análise
    btn_checkord = ctk.CTkButton(master=root, text="Selecionar Arquivo HTML", command=iniciar_processo_checkord, width=200)
    btn_checkord.grid(row=1, column=0, columnspan=3, pady=20)  # Ajuste para centralizar abaixo do rótulo

    # Centraliza os widgets na janela
    root.grid_rowconfigure(0, weight=1)
    root.grid_rowconfigure(1, weight=1)
    root.grid_columnconfigure(0, weight=1)
    root.grid_columnconfigure(1, weight=1)
    root.grid_columnconfigure(2, weight=1)

    root.mainloop()

if __name__ == "__main__":
    main()
