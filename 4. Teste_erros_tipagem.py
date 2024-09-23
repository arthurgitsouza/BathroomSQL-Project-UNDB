import os
import pandas as pd
import pyodbc

server = r'GDB-CIT-09\SQLEXPRESS'
database = 'UNDB_BANHEIROS'
username = 'SA'
password = '12090607'
conn = pyodbc.connect(f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}')

pasta = r'C:\Users\joao.souza\OneDrive - undb.edu.br\Formulários'

data_frames = []  

for arquivo in os.listdir(pasta):
    if arquivo.endswith('.xlsx'):
        caminho_arquivo = os.path.join(pasta, arquivo)
        df = pd.read_excel(caminho_arquivo)
        
        nome_banheiro = arquivo.split('.')[0]
        
        df = df.rename(columns={
            'Hora de início': 'HRINICIO',
            'Hora de conclusão': 'HRFIM',
            'Este local encontra-se LIMPO e devidamente higienizado para seu uso?': 'WCLIMPO',
            'Faltou algum insumo? (papel, sabonete, etc)': 'WCINSUMO',
            'Identificou algum equipamento danificado?': 'WCEQUIPDANIFICADO',
            'Quais insumos estão faltando nesse local?': 'WCINSUMODESCRICAO',
            'Quais equipamentos você identificou como danificados?': 'WCEQUIPDANIFICADODESCRICAO',
            'De uma forma geral, qual seu índice de satisfação enquanto usuário deste local? (0 à 5)': 'WCNOTA'
        })
        df['WCNOTA'] = pd.to_numeric(df['WCNOTA'], errors='coerce').fillna(0).astype(int)  
        df['NOMEBANHEIRO'] = nome_banheiro 
        data_frames.append(df) 

dataframe_completo = pd.concat(data_frames, ignore_index=True) 

cursor = conn.cursor()

for index, row in dataframe_completo.iterrows():
    cursor.execute("SELECT IDWC FROM BANHEIROS WHERE NOMEBANHEIRO = ?", (row['NOMEBANHEIRO'],))
    result = cursor.fetchone()
    if result:
        codbanheiro = result[0]
    else:
        cursor.execute("INSERT INTO BANHEIROS (NOMEBANHEIRO) VALUES (?)", (row['NOMEBANHEIRO'],))
        conn.commit()  
        cursor.execute("SELECT @@IDENTITY AS IDWC")
        codbanheiro = cursor.fetchone()[0]
    
    hrinicio = row['HRINICIO'] if pd.notnull(row['HRINICIO']) else ''
    hrfim = row['HRFIM'] if pd.notnull(row['HRFIM']) else ''
    wclimpo = row['WCLIMPO'] if pd.notnull(row['WCLIMPO']) else 'Não'
    wcinsumo = row['WCINSUMO'] if pd.notnull(row['WCINSUMO']) else 'Não'
    wcequipdanificado = row['WCEQUIPDANIFICADO'] if pd.notnull(row['WCEQUIPDANIFICADO']) else 'Não'
    wcinsumodescricao = row['WCINSUMODESCRICAO'] if pd.notnull(row['WCINSUMODESCRICAO']) else ''
    wcequipdanificadodescricao = row['WCEQUIPDANIFICADODESCRICAO'][:255] if pd.notnull(row['WCEQUIPDANIFICADODESCRICAO']) else ''  # Trunca para 255 caracteres

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