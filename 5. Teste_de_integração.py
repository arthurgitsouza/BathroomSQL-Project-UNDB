import os
import pandas as pd
import pyodbc

# Configuração dos dados do SQL Server
server = r'GDB-CIT-09\SQLEXPRESS'
database = 'UNDB_BANHEIROS'
username = 'SA'
password = '12090607'
conn = pyodbc.connect(f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}')

pasta = r'C:\Users\joao.souza\OneDrive - undb.edu.br\Formulários'  # Pasta dos arquivos

data_frames = []  # Lista para armazenar os DataFrames

# Processar todos os arquivos Excel na pasta
for arquivo in os.listdir(pasta):
    if arquivo.endswith('.xlsx'):
        caminho_arquivo = os.path.join(pasta, arquivo)
        df = pd.read_excel(caminho_arquivo)
        
        # Supondo que o nome do banheiro está no nome do arquivo, por exemplo, "BANHEIRO_1.xlsx"
        nome_banheiro = arquivo.split('.')[0]
        
        # Renomeia as colunas para facilitar a manipulação
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
        
        # Converte WCNOTA para inteiro, garantindo que valores não numéricos sejam tratados
        df['WCNOTA'] = pd.to_numeric(df['WCNOTA'], errors='coerce').fillna(0).astype(int)
        
        # Adiciona o nome do banheiro ao DataFrame
        df['NOMEBANHEIRO'] = nome_banheiro
        
        # Adiciona o DataFrame à lista
        data_frames.append(df)

# Concatena todos os DataFrames em um único DataFrame
dataframe_completo = pd.concat(data_frames, ignore_index=True)

cursor = conn.cursor()

# Loop para inserir os dados no banco de dados
for index, row in dataframe_completo.iterrows():
    # Verifica se o banheiro já existe
    cursor.execute("SELECT IDWC FROM BANHEIROS WHERE NOMEBANHEIRO = ?", (row['NOMEBANHEIRO'],))
    result = cursor.fetchone()
    
    if result:
        codbanheiro = result[0]
    else:
        # Insere o novo banheiro e recupera o ID gerado
        cursor.execute("INSERT INTO BANHEIROS (NOMEBANHEIRO) VALUES (?)", (row['NOMEBANHEIRO'],))
        conn.commit()  # É necessário fazer o commit para obter o ID gerado
        cursor.execute("SELECT @@IDENTITY AS IDWC")
        codbanheiro = cursor.fetchone()[0]
    
    # Converte valores para os tipos adequados e trunca se necessário
    hrinicio = row['HRINICIO'] if pd.notnull(row['HRINICIO']) else ''
    hrfim = row['HRFIM'] if pd.notnull(row['HRFIM']) else ''
    wclimpo = row['WCLIMPO'] if pd.notnull(row['WCLIMPO']) else 'Não'
    wcinsumo = row['WCINSUMO'] if pd.notnull(row['WCINSUMO']) else 'Não'
    wcequipdanificado = row['WCEQUIPDANIFICADO'] if pd.notnull(row['WCEQUIPDANIFICADO']) else 'Não'
    wcinsumodescricao = row['WCINSUMODESCRICAO'] if pd.notnull(row['WCINSUMODESCRICAO']) else ''
    wcequipdanificadodescricao = row['WCEQUIPDANIFICADODESCRICAO'][:100] if pd.notnull(row['WCEQUIPDANIFICADODESCRICAO']) else ''  # Trunca para 100 caracteres

    # Verifica se a resposta já existe
    cursor.execute('''
        SELECT 1 FROM RESPOSTAS WHERE HRINICIO = ? AND NOMEBANHEIRO = ?
        ''', 
        hrinicio, row['NOMEBANHEIRO']
    )
    
    if cursor.fetchone() is None:
        # Insere a resposta na tabela RESPOSTAS
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

# Commit das transações
conn.commit()
cursor.close()
conn.close()

print("Banheiros e respostas inseridos com sucesso!")