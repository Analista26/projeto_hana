import pandas as pd
import pyodbc
from tqdm import tqdm

import pandas as pd
import pyodbc
from tqdm import tqdm
import re

# === CONFIGURAÇÕES ===
#caminho_txt = r'Z:/00 Pastas de trabalho/ExportATOS_TOS4/PM - Equipamentos/EXPORT_TO_S4-FLEET_0001.txt'
caminho_txt = r'Z:/00 Pastas de trabalho/ExportATOS_TOS4/MM Classe/EXPORT_TO_S4-SWOR_0001.txt'
#caminho_txt = r'Z:/00 Pastas de trabalho/ExportATOS_TOS4/MM Produto e Caracteristica/EXPORT_TO_S4-MPOP_V_0001.txt'
#caminho_txt   = r'Z:/00 Pastas de trabalho/ExportATOS_TOS4/BP/EXPORT_TO_S4-LFA1_0001.txt'
nome_tabela   = 'SWOR_M01' # XXX XXX XXX XXX  XXX XXX XXX

server      = 'SPLCPVMSQLQA,23002'
database    = 'MDM_ONDA3'
username    = 'usr_mdm'
password    = 'HSw0raUlpVcBC'
sep         = '\t'
encoding    = 'utf-8'

# === LÊ O TXT COMO STRINGS PUROS ===
df_raw = pd.read_csv(caminho_txt, sep=sep, encoding=encoding, dtype=str, keep_default_na=False)
"""
df_raw = pd.read_csv(
    caminho_txt,
    sep=sep,
    encoding=encoding,
    dtype=str,
    keep_default_na=False,
    engine='python',
    on_bad_lines='skip'  # ou error_bad_lines=False em pandas <1.3
)
"""

# === FORMATAÇÃO DE CÉLULAS ===
def formatar_celula(x):
    if pd.isna(x) or x == '':
        return None
    s = str(x).strip()
    # detecta decimais no texto original
    if re.fullmatch(r'\d+[.,]\d+', s):
        f = float(s.replace(',', '.'))
        decimais = len(s) - max(s.find('.'), s.find(',')) - 1
        return f"{f:.{decimais}f}".replace('.', ',')
    return s


df = df_raw.apply(lambda col: col.astype(str).map(formatar_celula))


# === CRIAÇÃO DA TABELA ===
colunas_sql = ",\n    ".join(f"[{col}] VARCHAR(MAX)" for col in df.columns)
create_stmt = f"""
IF OBJECT_ID('dbo.{nome_tabela}', 'U') IS NOT NULL
    DROP TABLE dbo.{nome_tabela};

CREATE TABLE dbo.{nome_tabela} (
    {colunas_sql}
);
"""
print(create_stmt)

# === INSERE NO SQL SERVER ===
conn_str = (
    f"DRIVER={{ODBC Driver 17 for SQL Server}};"
    f"SERVER={server};DATABASE={database};"
    f"UID={username};PWD={password};"
    "TrustServerCertificate=yes;"
)
conn   = pyodbc.connect(conn_str)
cursor = conn.cursor()
cursor.execute(create_stmt)
conn.commit()

cols         = ", ".join(f"[{c}]" for c in df.columns)
placeholders = ", ".join("?" for _ in df.columns)
insert_sql   = f"INSERT INTO dbo.{nome_tabela} ({cols}) VALUES ({placeholders})"

for _, row in tqdm(df.iterrows(), total=len(df), desc=f"Inserindo em {nome_tabela}"):
    try:
        cursor.execute(insert_sql, tuple(row))
    except Exception as e:
        print(f"Erro na linha {row.name}: {e}")

conn.commit()
cursor.close()
conn.close()
print("✅ Pronto!")
