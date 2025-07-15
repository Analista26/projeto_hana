import pandas as pd
import pyodbc
from tqdm import tqdm

# === CONFIGURAÇÕES ===
caminho_excel = r'Z:/00 Pastas de trabalho/Asantos/09 - Demanda/Tabelas_Auxiliar/Organização_Empresa.xlsx'
nome_tabela = 'ORG_EMP'
server = 'SPLCPVMSQLQA,23002'
database = 'MDM_ONDA3'
username = 'usr_mdm'
password = 'HSw0raUlpVcBC'

# === LER O ARQUIVO EXCEL COMO OBJETO STRING (preserva tudo) ===
# Usa engine openpyxl e converte tudo explicitamente para texto
df_raw = pd.read_excel(caminho_excel, dtype=str)

# Converte tudo para string, mantendo formatação original do Excel
def formatar_celula(x):
    if pd.isnull(x):
        return None
    elif isinstance(x, float):
        # Formata float como no Excel (com vírgula), removendo notação científica
        return format(x, 'f').replace('.', ',')
    else:
        return str(x).strip()

df = df_raw.applymap(formatar_celula)

# === GERAR SCRIPT CREATE TABLE (todas as colunas como NVARCHAR(MAX)) ===
colunas_sql = ",\n    ".join(f"[{col}] VARCHAR(MAX)" for col in df.columns)

create_stmt = f"""
IF OBJECT_ID('dbo.{nome_tabela}', 'U') IS NOT NULL
    DROP TABLE dbo.{nome_tabela};

CREATE TABLE dbo.{nome_tabela} (
    {colunas_sql}
);
"""
print("Script de criação da tabela:\n", create_stmt)

# === CONEXÃO COM SQL SERVER ===
conn_str = f"""
    DRIVER={{ODBC Driver 17 for SQL Server}};
    SERVER={server};
    DATABASE={database};
    UID={username};
    PWD={password};
    TrustServerCertificate=yes;
"""
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# === CRIAR TABELA ===
cursor.execute(create_stmt)
conn.commit()

# === INSERIR OS DADOS COM tqdm ===
cols = ", ".join(f"[{col}]" for col in df.columns)
placeholders = ", ".join("?" for _ in df.columns)
insert_sql = f"INSERT INTO dbo.{nome_tabela} ({cols}) VALUES ({placeholders})"

for _, row in tqdm(df.iterrows(), total=len(df), desc=f"Inserindo linhas em {nome_tabela}"):
    try:
        cursor.execute(insert_sql, tuple(row))
    except Exception as e:
        print(f"\n❌ Erro ao inserir linha {row.name}: {e}")

conn.commit()
cursor.close()
conn.close()
print(f"✅ Tabela '{nome_tabela}' criada e dados inseridos com sucesso.")
