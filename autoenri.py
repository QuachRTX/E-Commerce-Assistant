__author__ = "Alexsander Stahnke"
__version__ = "1.1"
__status__ = "Development"
__email__ = "alexsander.stahnke@grendene.com.br"
__github__ = "https://github.com/QuachRTX"

import pandas as pd
from openpyxl import load_workbook
import customtkinter as ctk
from tkinter import filedialog, messagebox
import os

ctk.set_appearance_mode("System")  # Define o tema do sistema, claro ou escuro
ctk.set_default_color_theme("blue")  # Define o tema de cor para azul


def salvar_como():
    return filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])

def selecione_arquivo(titulo, tipo_arquivo):
    return filedialog.askopenfilename(title=titulo, filetypes=[(tipo_arquivo, "*.xlsx"), ("Todos os arquivos", "*.*")])

def selecionar_enriquecimento():
    global arquivo_enriquecimento, botao_executar  # Adicionado botao_executar ao global para alterar seu estado
    arquivo_enriquecimento = selecione_arquivo("Selecionar planilha de enriquecimento", "Excel")
    if arquivo_enriquecimento:
        label_enriquecimento.configure(text=f"Selecionado: {os.path.basename(arquivo_enriquecimento)}")
        botao_executar.configure(state="normal")  # Habilita o botão após a seleção do arquivo

def selecionar_pesos_medidas():
    global arquivo_pesos_medidas
    arquivo_pesos_medidas = selecione_arquivo("Selecionar pesos e medidas", "Excel")
    if arquivo_pesos_medidas:
        label_pesos_medidas.configure(text=f"Selecionado: {os.path.basename(arquivo_pesos_medidas)}")

def selecionar_descricao():
    global arquivo_descricao
    arquivo_descricao = selecione_arquivo("Selecionar descrição", "Excel")
    if arquivo_descricao:
        label_descricao.configure(text=f"Selecionado: {os.path.basename(arquivo_descricao)}")

def ajustar_largura_colunas(ws):
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name
        for cell in col:
            try:  # Necessary to avoid error on empty cells
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2) * 1.1  # Ajuste conforme necessário
        ws.column_dimensions[column].width = adjusted_width

def atualizar_pesos(df_enriquecimento, arquivo_pesos):
    if arquivo_pesos:
        df_pesos = pd.read_excel(arquivo_pesos, sheet_name='Antiga')
        pesos_lookup = df_pesos.set_index('Modelo')['Peso (g) PÉ'].to_dict()

        def encontrar_peso(codigo):
            peso = pesos_lookup.get(codigo)
            return f"Peso aproximado: {peso}g" if peso else None

        for index, row in df_enriquecimento.iterrows():
            peso_info = encontrar_peso(row['Código Produto'])
            if peso_info and pd.isna(row['Características']):
                df_enriquecimento.at[index, 'Características'] = peso_info

def atualizar_descricao(df_enriquecimento, arquivo_descricao):
    if arquivo_descricao:
        df_descricao = pd.read_excel(arquivo_descricao)
        # Atualizar descrições dos produtos
        for index, row in df_enriquecimento.iterrows():
            codigo = row['Código Produto']
            if codigo in df_descricao['Product Code'].values:
                descricao = df_descricao[df_descricao['Product Code'] == codigo]['Product Description PT'].iloc[0]
                # Verificar se a célula já está preenchida
                if pd.isna(row['Descrição Amigavel']):
                    df_enriquecimento.at[index, 'Descrição Amigavel'] = descricao

def atualizar_planilha(planilha_enriquecimento, arquivo_pesos, arquivo_descricao, arquivo_saida):
    df_enriquecimento = pd.read_excel(planilha_enriquecimento)

    # Atualizar nomes dos produtos com as regras especificadas
    for index, row in df_enriquecimento.iterrows():
        # Verificar se a célula já está preenchida antes de atualizar o nome
        if pd.isna(row['Nome Enriquecido']):
            nome_original = row['Descrição do Item']
            # Ignorar a linha se 'Descrição do Item' estiver vazia
            if pd.isna(nome_original):
                continue  # Pula para a próxima iteração do loop sem fazer mais nada
            nome_modificado = ' '.join(word.capitalize() for word in nome_original.split())
            nome_modificado = nome_modificado.replace('Ad', '').replace('Bb', 'Baby').replace('Inf', 'Infantil')
            df_enriquecimento.at[index, 'Nome Enriquecido'] = nome_modificado

        # Atualizar cor dos produtos com as regras especificadas
        if pd.isna(row['Cor Enriquecida']):
            cor_original = row['Descrição da Cor'].split('/')[0].split()[0]
            cor_modificada = cor_original.capitalize()
            cor_modificada = cor_modificada.replace('Vidro', 'Transparente').replace('Ouro', 'Dourado')
            df_enriquecimento.at[index, 'Cor Enriquecida'] = cor_modificada

    # Exemplo de atualização de pesos
    atualizar_pesos(df_enriquecimento, arquivo_pesos)
    atualizar_descricao(df_enriquecimento, arquivo_descricao)

    df_enriquecimento.to_excel(arquivo_saida, index=False)
    wb = load_workbook(arquivo_saida)
    ws = wb.active
    ajustar_largura_colunas(ws)
    wb.save(arquivo_saida)

def executar_processo():
    global arquivo_saida
    arquivo_saida = salvar_como()
    if arquivo_enriquecimento and arquivo_saida:  # Só verifica o arquivo de enriquecimento e a saída
        try:
            # Corrigido para passar todos os argumentos necessários
            atualizar_planilha(arquivo_enriquecimento, arquivo_pesos_medidas, arquivo_descricao, arquivo_saida)
            messagebox.showinfo("Sucesso", "O arquivo foi salvo com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
    else:
        messagebox.showinfo("Aviso", "Por favor, selecione o arquivo de enriquecimento necessário e escolha onde salvar o resultado.")

def main():
    global label_enriquecimento, label_pesos_medidas, label_descricao, arquivo_enriquecimento, arquivo_pesos_medidas, arquivo_descricao, botao_executar  # Adicionar botao_executar aqui
    arquivo_enriquecimento = arquivo_pesos_medidas = arquivo_descricao = None
    root = ctk.CTk()
    root.title("Auto Enriquecimento")
    root.geometry("600x500")

    ctk.CTkLabel(root, text="Selecionar Planilha de Enriquecimento").pack(pady=10)
    ctk.CTkButton(root, text="Selecionar", command=selecionar_enriquecimento).pack()
    label_enriquecimento = ctk.CTkLabel(root, text="")
    label_enriquecimento.pack(pady=10)

    ctk.CTkLabel(root, text="Selecionar Pesos e Medidas (Opcional)").pack(pady=10)
    ctk.CTkButton(root, text="Selecionar", command=selecionar_pesos_medidas).pack()
    label_pesos_medidas = ctk.CTkLabel(root, text="")
    label_pesos_medidas.pack(pady=10)

    ctk.CTkLabel(root, text="Selecionar Descrição (Opcional)").pack(pady=10)
    ctk.CTkButton(root, text="Selecionar", command=selecionar_descricao).pack()
    label_descricao = ctk.CTkLabel(root, text="")
    label_descricao.pack(pady=10)

    botao_executar = ctk.CTkButton(root, text="Executar Enriquecimento e Salvar", command=executar_processo, state="disabled")  # Botão começa desabilitado
    botao_executar.pack(pady=20)

    #Aviso de desenvolvimento
    bold_font = ("Helvetica", 10, "bold")
    versao_label = ctk.CTkLabel(root, text="EM DESENVOLVIMENTO  ", text_color='red', font=bold_font)
    versao_label.place(relx=1.0, rely=1.0, anchor="se")  # Ancora o label no canto inferior direito

    root.mainloop()

if __name__ == "__main__":
    main()