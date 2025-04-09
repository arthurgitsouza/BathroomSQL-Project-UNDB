import pandas as pd
import os
import pyodbc
from dotenv import load_dotenv

load_dotenv()

# Configuração dos dados do SQL Server
server = r'GDB-CIT-09\SQLEXPRESS'
database = 'UNDB_BANHEIROS'
username = os.getenv('DB_USER')
password = os.getenv('DB_PASSWORD')
conn = pyodbc.connect(f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}') #conecta com o SQL server usando essas credenciais

cursor = conn.cursor()

cursor.execute("DELETE FROM RESPOSTAS") #Limpar a tabela de respostas antes de inserir novos dados
conn.commit()

pasta = r'C:\Users\joao.souza\OneDrive - undb.edu.br\Formulários'  #Caminho para a pasta.

data_frames = []  #Lista para armazenar os DataFrames

# Processar todos os arquivos Excel na pasta
for arquivo in os.listdir(pasta): #Para cada item na pasta # Função os.listdir: é usada para listar todos os arquivos presentes no diretório especificado
    if arquivo.endswith('.xlsx'):
        """
        Docstring Endwith: 
            -> Ela retoma o súfixo que for determinado em formato de string no parâmetro
            Parâmetros:
                f: '.xlsx' Para arquivos excel    
        """
        caminho_arquivo = os.path.join(pasta, arquivo) #Ele cria um caminho completo para cada arquivo que será lido #pasta: a pasta, #arquivo: cada arquivo
        df = pd.read_excel(caminho_arquivo) #Lê o arquivo de a cordo com o loop e coloca em um dataframe
        
        nome_banheiro = arquivo.split('.')[0] #Utiliza o nome do arquivo como nome do banheiro
        
        df = df.rename(columns={ #Renomeia as colunas para facilitar a manipulação
            'Hora de início': 'HRINICIO',
            'Hora de conclusão': 'HRFIM',
            'Este local encontra-se LIMPO e devidamente higienizado para seu uso?': 'WCLIMPO',
            'Faltou algum insumo? (papel, sabonete, etc)': 'WCINSUMO',
            'Identificou algum equipamento danificado?': 'WCEQUIPDANIFICADO',
            'Quais insumos estão faltando nesse local?': 'WCINSUMODESCRICAO',
            'Quais equipamentos você identificou como danificados?': 'WCEQUIPDANIFICADODESCRICAO',
            'De uma forma geral, qual seu índice de satisfação enquanto usuário deste local? (0 à 5)': 'WCNOTA'
        })
        
        df['WCNOTA'] = pd.to_numeric(df['WCNOTA'], errors='coerce').fillna(0).astype(int) # Converte WCNOTA para inteiro
        
        df['NOMEBANHEIRO'] = nome_banheiro #Adiciona o nome do banheiro ao DataFrame
        
        data_frames.append(df) #Adiciona a lista de data_frames

dataframe_completo = pd.concat(data_frames, ignore_index=True) #Concatena todos os DataFrames da lista "data_frames" em um único DataFrame

for index, row in dataframe_completo.iterrows(): # index: para cada indice da linha no dataframe | row: Contém os valores de uma linha específica, onde as colunas são os rótulos e os dados são os valores correspondentes.
    #Verifica se o banheiro já existe
    cursor.execute("SELECT IDWC FROM BANHEIROS WHERE NOMEBANHEIRO = ?", (row['NOMEBANHEIRO'],))
    result = cursor.fetchone()
    
    if result:
        codbanheiro = result[0]
    else:
        cursor.execute("INSERT INTO BANHEIROS (NOMEBANHEIRO) VALUES (?)", (row['NOMEBANHEIRO'],))
        conn.commit()
        cursor.execute("SELECT @@IDENTITY AS IDWC")
        codbanheiro = cursor.fetchone()[0]
    
    #Convertendo valores para os tipos que o SQL aceite e truncando as respostas se necessário (devido ao limite de 255 caracteres)
    hrinicio = row['HRINICIO'] if pd.notnull(row['HRINICIO']) else ''
    hrfim = row['HRFIM'] if pd.notnull(row['HRFIM']) else ''
    wclimpo = row['WCLIMPO'] if pd.notnull(row['WCLIMPO']) else 'Não'
    wcinsumo = row['WCINSUMO'] if pd.notnull(row['WCINSUMO']) else 'Não'
    wcequipdanificado = row['WCEQUIPDANIFICADO'] if pd.notnull(row['WCEQUIPDANIFICADO']) else 'Não'
    wcinsumodescricao = row['WCINSUMODESCRICAO'] if pd.notnull(row['WCINSUMODESCRICAO']) else ''
    wcequipdanificadodescricao = row['WCEQUIPDANIFICADODESCRICAO'][:100] if pd.notnull(row['WCEQUIPDANIFICADODESCRICAO']) else ''  # Trunca para 100 caracteres

    #Verifica se a resposta já existe
    cursor.execute('''
        SELECT 1 FROM RESPOSTAS WHERE HRINICIO = ? AND CODBANHEIRO = ?
        ''', 
        hrinicio, codbanheiro
    )
    
    if cursor.fetchone() is None:
        #Insere a resposta na tabela RESPOSTAS
        cursor.execute('''
            INSERT INTO dbo.RESPOSTAS (HRINICIO, HRFIM, WCLIMPO, WCINSUMO, WCEQUIPDANIFICADO, WCINSUMODESCRICAO, WCEQUIPDANIFICADODESCRICAO, WCNOTA, CODBANHEIRO)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''',
            hrinicio,
            hrfim,
            wclimpo,
            wcinsumo,
            wcequipdanificado,
            wcinsumodescricao,
            wcequipdanificadodescricao,
            row['WCNOTA'],
            codbanheiro
        )

#Commit das transações
conn.commit()
cursor.close()
conn.close()

print("Banheiros e respostas inseridos com sucesso!")
