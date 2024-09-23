import os
import pandas as pd
import pyodbc

# listinha com dados do SQL Server para configuração
server = 'GDB-CIT-09\SQLEXPRESS'
database = 'UNDB_BANHEIROS'
username = 'SA'
password = '12090607'
conn = pyodbc.connect(f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}') #conecta com o SQL server usando essas credenciais


pasta = r'C:\Users\joao.souza\OneDrive - undb.edu.br\Formulários' # Pasta dos arquivos


data_frames = [] # Lista para armazenar os DataFrames


for arquivo in os.listdir(pasta): #Para cada item na pasta # Função os.listdir: é usada para listar todos os arquivos presentes no diretório especificado
    if arquivo.endswith('.xlsx'):
        """
        Docstring Endwith: 
            -> Ela retoma o súfixo que for determinado em formato de string no parâmetro
            Parâmetros:
                f: '.xlsx' Para arquivos excel
                
        """
        caminho_arquivo = os.path.join(pasta, arquivo) # Ele cria um caminho completo para cada arquivo que será lido #pasta: a pasta, #arquivo: cada arquivo
        df = pd.read_excel(caminho_arquivo) # Lê o arquivo de a cordo com o loop e coloca em um dataframe
        data_frames.append(df) # Adiciona a lista de data_frames

dataframe_completo = pd.concat(data_frames, ignore_index=True) # Concatena todos os DataFrames da lista "data_frames" em um único DataFrame
print(dataframe_completo.columns)
print('Server conectado com sucesso!')


