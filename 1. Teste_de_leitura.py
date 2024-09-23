import pandas as pd
import os
arquivos = [] 

pasta = r'C:\Users\joao.souza\OneDrive - undb.edu.br\Formulários' #Caminho para a pasta.

for f in os.listdir(pasta): #Para cada item na pasta # Função os.listdir: é usada para listar todos os arquivos presentes no diretório especificado
    if f.endswith('.xlsx'): #Se o arquivo for  terminado em ".xlsx"  #Função "endswith": 
        """
        Docstring Endwith: 
            -> Ela retoma o súfixo que for determinado em formato de string no parâmetro
            Parâmetros:
                f: '.xlsx' Para arquivos excel
                
        """
        arquivos.append(f) #então adicione a lista arquivos.


data_frames = [] #Lista para armazenar os DataFrames

for nome_arquivo in arquivos: #Para cada arquivo na lista [arquivos]:
    arquivo = os.path.join(pasta, nome_arquivo) #Cria um caminho para cada arquivo
    df = pd.read_excel(arquivo) 
    '''
    
    Docstring read_excel:
        -> Essa função lê cada arquivo que vai passar pelo loop do for e carrega os dados para colocar no dataframe
        ''' 
    data_frames.append(df)  #Adiciona o dataFrame à lista.
    
    print(f'Arquivo lido: {nome_arquivo}') 
    print(f'Número de linhas no arquivo {nome_arquivo}: {len(df)}') 

dataframe_completo = pd.concat(data_frames, ignore_index=True) #Concatena (junta) todos os dataframes presentes na lista "dataframes" em um dataframe (Importante lembrar que é uma variável que está sendo atribuida, pois com uma lista foi impossivel usar funções como "head()")

print(f'{"PRIMEIRAS LINHAS DO DATAFRAME":-^70}')