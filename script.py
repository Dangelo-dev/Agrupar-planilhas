import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk

# Função para solicitar ao usuário para selecionar um diretório
def selecionar_diretorio():
    diretorio = filedialog.askdirectory()
    if diretorio:
        print("Diretorio selecionado:", diretorio)
        processar_arquivos(diretorio)
        root.destroy() # Fecha a janela após selecionar o diretório
    else:
        print("Nenhum diretorio selecionado.")

# Função para processar os arquivos e mostrar ao usuário do que está sendo carregado
def processar_arquivos(diretorio):
    files = [arquivo for arquivo in os.listdir(diretorio) if arquivo.endswith('xlsx')]
    num_files = len(files)
    dfs = [] # Inicializa a lista para armazenar os DataFrames lidos de cada arquivo
    if num_files > 0:
        progresso["maximum"] = num_files
        for i, arquivo in enumerate(files):
            print('Carregando arquivo {0}...'.format(arquivo))
            try:
                df = pd.read_excel(os.path.join(diretorio, arquivo)) # Juntando os arquivos dentro do diretório selecionado
                dfs.append(df)  # Adicionando o DataFrame à lista
                progresso["value"] = i + 1
                root.update_idletasks()
            except Exception as Err:
                print(f'Erro ao ler arquivo {arquivo}: {Err}')
        if dfs:        
            df_principal = pd.concat(dfs) # concatenando os arquivos do DataFrame principal
            df_principal.to_excel('Todas_lojas.xlsx', index=False) # Criando um arquivo com todas as planilhas dentro da pasta juntas
            print("Arquivos combinados e salvos como Todas_lojas.xlsx")
        else:
            print('Nenhum DataFrame válido encontrado nos arquivos .xlsx.')
    else:
        print('Nenhum arquivo .xlsx encontrado no diretório.')

# Função para salvar o DataFrame principal como arquivo .xlsx em um local selecionado pelo usuário
def salvar_arquivo(df_principal):
    local_arquivo = filedialog.asksaveasfilename(defaultextension=".xlsx")
    if local_arquivo:
        df_principal.to_excel(local_arquivo, index=False)
        print(f"Arquivo salvo como {local_arquivo}")

# Criar a janela principal
root = tk.Tk()
root.title("Juntar arquivos '.xlsx'")

# Definir as dimensões da janela
largura_janela = 400
altura_janela = 200
pos_x = (root.winfo_screenwidth() - largura_janela) // 2
pos_y = (root.winfo_screenheight() - altura_janela) // 2
root.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")

# Criar um frame para organizar os elementos
frame = tk.Frame(root)
frame.place(relx=0.5, rely=0.5, anchor="center")

# Botão para selecionar diretório
btn_selecionar = tk.Button(frame, text="Selecionar diretório", command=selecionar_diretorio)
btn_selecionar.pack()

# Barra de progresso
progresso = ttk.Progressbar(frame, orient="horizontal", length=300, mode="determinate")
progresso.pack(pady=10)

# Executar o loop da aplicação
root.mainloop()