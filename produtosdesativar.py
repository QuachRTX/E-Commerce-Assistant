__author__ = "Alexsander Stahnke"
__version__ = "1.1"
__status__ = "Beta"
__email__ = "quach.vrc@gmail.com"
__github__ = "https://github.com/QuachRTX"

from tkinter import filedialog, messagebox
import pandas as pd
import customtkinter as ctk
import os
from shutil import copy2
import datetime

def processar_diretorio(diretorio_loja, nome_arquivo):
    caminho_arquivo = os.path.join(diretorio_loja, nome_arquivo)

    # Verifica se o arquivo existe
    if not os.path.exists(caminho_arquivo):
        messagebox.showerror("Erro", f"{nome_arquivo} não encontrado.")
        return

    caminho_temporario = os.path.join(diretorio_loja, 'temp_' + nome_arquivo)
    
    # Criando cópia temporária
    copy2(caminho_arquivo, caminho_temporario)

    # Carregar dados
    df = pd.read_excel(caminho_temporario, sheet_name='ESTOQUE_COR', header=5)
    
    # Aplicar filtros
    filtro = (df['QTDE REPOR'] == 0) & \
             (df['Reserva'] == 0) & \
             (df['Disponivel'] == 0) & \
             (df['Dias no estoque (venda)'] >= 60) & \
             (df['Reposições Pendentes'] == 0) & \
             (df['STATUS PRODUTO'] == 'FORA DE LINHA')
 
    df_filtrado = df[filtro]

    # Selecionar colunas
    colunas_selecionadas = ['SKU_COR', 'Código', 'Descrição', 'Cod_Cor']
    df_final = df_filtrado[colunas_selecionadas]

    # Escolher local para salvar o arquivo
    data_atual = datetime.datetime.now().strftime("%d.%m")
    nome_arquivo_saida = f'Produtos Desativar - {data_atual}.xlsx'
    output_path = filedialog.asksaveasfilename(initialdir=diretorio_loja, title="Salvar como",
                                               filetypes=[("Excel files", "*.xlsx")],
                                               defaultextension=".xlsx", initialfile=nome_arquivo_saida)
    if output_path:
        # Salvar resultado
        df_final.to_excel(output_path, index=False)

        # Remover arquivo temporário
        os.remove(caminho_temporario)

        messagebox.showinfo("Sucesso", f"Arquivo '{nome_arquivo_saida}' foi salvo com sucesso em '{output_path}'.")
    else:
        messagebox.showinfo("Cancelado", "Operação cancelada, local para salvar o arquivo não foi selecionado.")
        os.remove(caminho_temporario)

def main():
    user_name = os.getlogin()  # Obtém o nome de usuário do sistema
    root = ctk.CTk()
    root.title("Produtos Desativar")
    root.geometry("500x400")

    bold_font = ("Helvetica", 12, "bold")
    instrucao_label = ctk.CTkLabel(master=root, text="Selecione a loja para listar os produtos sem estoque/reposição", height=10)
    instrucao_label.pack(pady=40)


    root.mainloop()

if __name__ == "__main__":
    main()
