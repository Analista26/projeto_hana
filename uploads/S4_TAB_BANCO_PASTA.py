import pandas as pd
import pyodbc
from tqdm import tqdm
import re
import os

# === CONFIGURAÇÕES ===
#pasta_txt    = r'Z:/00 Pastas de trabalho/ExportATOS_TOS4/PM - Equipamentos/'  # Pasta onde estão os arquivos
pasta_txt    = r'E:/Carga SQL Server/Dados/AUSP/'



nome_tabela  = 'AUSP_M01'
server       = 'SPLCPVMSQLQA,23002'
database     = 'MDM_ONDA3'
username     = 'usr_mdm'
password     = 'HSw0raUlpVcBC'
sep          = '\t'
encoding     = 'utf-8'

import os

def renomear_arquivos_sem_extensao(pasta, extensao='.txt'):
    """
    Renomeia todos os arquivos na pasta que não possuem extensão,
    acrescentando a extensão especificada.
    """
    arquivos = os.listdir(pasta)
    for arquivo in arquivos:
        caminho_antigo = os.path.join(pasta, arquivo)
        if os.path.isfile(caminho_antigo) and '.' not in arquivo:
            novo_nome = arquivo + extensao
            caminho_novo = os.path.join(pasta, novo_nome)
            os.rename(caminho_antigo, caminho_novo)
            print(f"🔄 Renomeado: {arquivo} -> {novo_nome}")

            
# === RENOMEIA ARQUIVOS SAP SEM EXTENSÃO PARA .TXT ===
renomear_arquivos_sem_extensao(pasta_txt)

# === FUNÇÃO DE FORMATAÇÃO ===
def formatar_celula(x):
    if pd.isna(x) or x == '':
        return None
    s = str(x).strip()
    if re.fullmatch(r'\d+[.,]\d+', s):
        f = float(s.replace(',', '.'))
        decimais = len(s) - max(s.find('.'), s.find(',')) - 1
        return f"{f:.{decimais}f}".replace('.', ',')
    return s

# === CONEXÃO SQL SERVER ===
conn_str = (
    f"DRIVER={{ODBC Driver 17 for SQL Server}};"
    f"SERVER={server};DATABASE={database};"
    f"UID={username};PWD={password};"
    "TrustServerCertificate=yes;"
)
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# === VERIFICA SE A TABELA EXISTE ===
check_table_sql = f"""
IF OBJECT_ID('dbo.{nome_tabela}', 'U') IS NULL
BEGIN
    SELECT 0; -- tabela não existe
END
ELSE
BEGIN
    SELECT 1; -- tabela existe
END
"""
cursor.execute(check_table_sql)
existe = cursor.fetchone()[0]

# === CRIA A TABELA CASO NÃO EXISTA ===
if not existe:
    # Lê apenas um arquivo de exemplo para criar a estrutura
    #arquivos = [f for f in os.listdir(pasta_txt) if f.endswith('.txt')]
    arquivos = [f for f in os.listdir(pasta_txt)
            if os.path.isfile(os.path.join(pasta_txt, f)) and ('.' not in f or f.endswith('.txt'))]

    if not arquivos:
        print("❌ Nenhum arquivo .txt encontrado na pasta.")
        exit()

    caminho_exemplo = os.path.join(pasta_txt, arquivos[0])
    df_exemplo = pd.read_csv(caminho_exemplo, sep=sep, encoding=encoding, dtype=str, keep_default_na=False, engine='python', on_bad_lines='skip')
    df_exemplo = df_exemplo.apply(lambda col: col.astype(str).map(formatar_celula))

    colunas_sql = ",\n    ".join(f"[{col}] VARCHAR(MAX)" for col in df_exemplo.columns)
    create_stmt = f"CREATE TABLE dbo.{nome_tabela} ({colunas_sql});"
    cursor.execute(create_stmt)
    conn.commit()
    print(f"✅ Tabela '{nome_tabela}' criada.")

# === INSERE OS DADOS DE TODOS OS ARQUIVOS ===
arquivos = [f for f in os.listdir(pasta_txt) if f.endswith('.txt')]
for arquivo in arquivos:
    caminho_txt = os.path.join(pasta_txt, arquivo)
    print(f"🔄 Processando {arquivo}")

    df_raw = pd.read_csv(
        caminho_txt,
        sep=sep,
        encoding=encoding,
        dtype=str,
        keep_default_na=False,
        engine='python',
        on_bad_lines='skip'
    )

    df = df_raw.apply(lambda col: col.astype(str).map(formatar_celula))

    # === INSERE NO SQL SERVER ===
    cols = ", ".join(f"[{c}]" for c in df.columns)
    placeholders = ", ".join("?" for _ in df.columns)
    insert_sql = f"INSERT INTO dbo.{nome_tabela} ({cols}) VALUES ({placeholders})"

    for _, row in tqdm(df.iterrows(), total=len(df), desc=f"Inserindo {arquivo}"):
        try:
            cursor.execute(insert_sql, tuple(row))
        except Exception as e:
            print(f"❌ Erro na linha {row.name}: {e}")

    conn.commit()
    print(f"✅ Inserção concluída para {arquivo}")

cursor.close()
conn.close()
print("🎯 Todos os arquivos foram processados com sucesso.")
