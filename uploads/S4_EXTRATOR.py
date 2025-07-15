#S4_EXTRATOR
import os
import pandas as pd
from sqlalchemy import create_engine
import urllib
from tqdm import tqdm

# === CONFIGURAÇÕES GERAIS ===

# Conexão SQL Server
conn_str = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=SeuServidor;"
    "DATABASE=SeuBanco;"
    "UID=SeuUsuario;"
    "PWD=SuaSenha;"
    "TrustServerCertificate=yes;"
)
conn_str_escaped = urllib.parse.quote_plus(conn_str)
engine = create_engine(f"mssql+pyodbc:///?odbc_connect={conn_str_escaped}")

# Diretório base de saída
output_base_dir = r"F:\BackOffice_GERAL\Projeto S4 Hana\Onda 2\9. Saneamento\Extraidos"
os.makedirs(output_base_dir, exist_ok=True)

# === INPUTS ===

# Excel com lista de tabelas
excel_tabelas = r"F:\BackOffice_GERAL\Projeto S4 Hana\Onda 2\9. Saneamento\Tabelas_Extracao.xlsx"
df_tabelas = pd.read_excel(excel_tabelas)

# Excel com lista de itens a extrair
excel_lista_itens = r"F:\BackOffice_GERAL\Projeto S4 Hana\Onda 2\9. Saneamento\Lista_Itens_Extracao.xlsx"
df_lista_itens = pd.read_excel(excel_lista_itens)

# Detecta automaticamente a coluna presente (ex: MATNR ou LIFNR)
coluna_itens = df_lista_itens.columns[0].strip().upper()
lista_itens = df_lista_itens.iloc[:, 0].astype(str).str.strip().str.upper().tolist()

print(f"[INFO] {len(lista_itens)} itens carregados ({coluna_itens}).")

# === Função para dividir em batches ===

def dividir_em_batches(lista, tamanho_batch):
    for i in range(0, len(lista), tamanho_batch):
        yield lista[i:i + tamanho_batch]

# === Mapeamento coluna_chave -> coluna_filtro real no SQL ===

# Ajuste conforme sua estrutura real no banco SAP S4
mapeamento_filtros = {
    "MATNR": "PRODUCT",
    "MARTNR": "PRODUCT",
    "LIFNR": "PARTNER",
    "KUNNR": "CUSTOMER",
    "PARTNER": "PARTNER",
}

# === PROCESSAMENTO ===

for _, row in tqdm(df_tabelas.iterrows(), total=len(df_tabelas), desc="Extraindo tabelas"):
    tabela_nome = row["TABELA"]
    coluna_chave = row["COLUNA_CHAVE"].strip().upper()
    coluna_filtro = mapeamento_filtros.get(coluna_chave, coluna_chave)  # default = coluna_chave

    try:
        print(f"[INFO] Processando tabela: {tabela_nome} | chave: {coluna_chave} | filtro: {coluna_filtro}")

        # Cria diretório específico para cada tabela
        output_dir = os.path.join(output_base_dir, f"S_{tabela_nome.split('_')[0].upper()}")
        os.makedirs(output_dir, exist_ok=True)

        # === Processa em batches ===
        df_total = pd.DataFrame()

        for batch in dividir_em_batches(lista_itens, 500):  # Batches de 500 itens
            placeholders = ",".join([f"'{item}'" for item in batch])

            query = f"""
                SELECT *
                FROM "{tabela_nome}"
                WHERE {coluna_filtro} IN ({placeholders})
            """

            with engine.connect() as conn:
                df_batch = pd.read_sql_query(query, conn)

            df_total = pd.concat([df_total, df_batch], ignore_index=True)

        print(f"[INFO] {len(df_total)} registros totais extraídos de {tabela_nome}.")

        # === Ajustes de formatação ===
        if coluna_chave in df_total.columns:
            if coluna_chave in ["MATNR", "MARTNR"]:
                df_total[coluna_chave] = df_total[coluna_chave].astype(str).str.zfill(18)
            elif coluna_chave in ["LIFNR", "KUNNR", "PARTNER"]:
                df_total[coluna_chave] = df_total[coluna_chave].astype(str).str.zfill(10)

        # === Salvar Excel ===
        output_path = os.path.join(output_dir, f"{tabela_nome}.xlsx")
        df_total.to_excel(output_path, index=False)
        print(f"[✅] Tabela {tabela_nome} salva em {output_path}")

    except Exception as e:
        print(f"[ERRO] Falha ao processar tabela {tabela_nome}: {e}")
