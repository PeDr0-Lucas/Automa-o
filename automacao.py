import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog

def selecionar_arquivo():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    arquivo_selecionado = filedialog.askopenfilename(
        title="Selecione um arquivo Excel",
        filetypes=[("Excel files", "*.xlsx;*.xls")]
    )
    return arquivo_selecionado

def selecionar_pasta_para_salvar():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    pasta_selecionada = filedialog.askdirectory(title="Selecione a pasta para salvar o arquivo")
    return pasta_selecionada

# Selecionar o arquivo
caminho_arquivo_original = selecionar_arquivo()
if not caminho_arquivo_original:
    print("Nenhum arquivo selecionado. O programa será encerrado.")
    exit()

print("Arquivo selecionado pelo usuário:", caminho_arquivo_original)

# Selecionar a pasta para salvar o novo arquivo
diretorio_salvar = selecionar_pasta_para_salvar()
if not diretorio_salvar:
    print("Nenhuma pasta selecionada. O programa será encerrado.")
    exit()

# Ler o arquivo original
df = pd.read_excel(caminho_arquivo_original)

# Definir as colunas a serem excluídas
colunas_para_excluir = ['Data de Admissão', 'Idade']
colunas_existentes = [col for col in colunas_para_excluir if col in df.columns]
df = df.drop(columns=colunas_existentes, errors='ignore')

# Salvar em um novo arquivo Excel na pasta selecionada
novo_caminho_arquivo = os.path.join(diretorio_salvar, 'novo_arquivo.xlsx')
df.to_excel(novo_caminho_arquivo, index=False)
print("Arquivo salvo com sucesso em:", novo_caminho_arquivo)

# Manter o terminal aberto
input("Pressione Enter para sair...")
