import pandas as pd
from sqlalchemy import create_engine
import urllib
import os

# === CONFIGURAÇÃO DE CONEXÃO COM SQLALCHEMY ===
conn_str = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=SPLCPVMSQLQA,23002;"
    "DATABASE=MDM_ONDA3;"
    "UID=usr_mdm;"
    "PWD=HSw0raUlpVcBC;"
    "TrustServerCertificate=yes;"
)
conn_str_escaped = urllib.parse.quote_plus(conn_str)
engine = create_engine(f"mssql+pyodbc:///?odbc_connect={conn_str_escaped}")

# === QUERY PARA TRAZER TABELA x COLUNA ===
query = """
select
    tab.name as table_name, 
    col.name as column_name
from sys.tables as tab
    inner join sys.columns as col
        on tab.object_id = col.object_id
"""

# === EXECUÇÃO DA QUERY USANDO SQLALCHEMY ===
df_sql = pd.read_sql(query, engine)

# === CARREGAR SEU EXCEL EXISTENTE ===
excel_path = r"Z:/00 Pastas de trabalho/Asantos/07 - Dev/Mapeamento_Fonecedor.xlsx"

if not os.path.exists(excel_path):
    print(f"❌ Arquivo não encontrado: {excel_path}")
    exit(1)
else:
    print(f"✅ Arquivo encontrado: {excel_path}")

df_excel = pd.read_excel(excel_path, engine="openpyxl")

# === MAPEAMENTO: BUSCAR TABELA DE CADA CAMPO ===
def encontrar_tabela(campo):
    resultado = df_sql[df_sql['column_name'].str.upper() == str(campo).strip().upper()]
    if not resultado.empty:
        return ', '.join(resultado['table_name'].unique())
    else:
        return 'NÃO ENCONTRADO'

# === APLICAR NA COLUNA I (8º índice) E CRIAR NOVA COLUNA COM RESULTADO ===
coluna_i_nome = df_excel.columns[8]
df_excel['Tabela_encontrada'] = df_excel[coluna_i_nome].apply(encontrar_tabela)

# === SALVAR RESULTADO ===
output_path = excel_path.replace(".xlsx", "_Tabelas.xlsx")
df_excel.to_excel(output_path, index=False)
print(f"✅ Mapeamento concluído. Arquivo salvo em: {output_path}")
