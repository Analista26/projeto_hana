import pandas as pd
import math

# Caminho do arquivo de entrada
arquivo_entrada = r'Z:/00 Pastas de trabalho/Asantos/08 - Lista_Carga/Lista_Carga_Cacteristicas.xlsx'

# Carregar o Excel
df = pd.read_excel(arquivo_entrada, dtype=str)  # garante leitura como string

# Garante MATNR como string com 18 caracteres preenchidos com zeros à esquerda
#if 'MATNR' in df.columns:
#    df['MATNR'] = df['MATNR'].astype(str).str.zfill(18)
if 'ATNAM' in df.columns:
    df['ATNAM'] = df['ATNAM'].astype(str).str.strip()  
    

# Quantidade máxima por arquivo
max_por_arquivo = 10000

# Número total de splits necessários
num_arquivos = math.ceil(len(df) / max_por_arquivo)

# Pasta de saída
pasta_saida = r'Z:/00 Pastas de trabalho/Asantos/08 - Lista_Carga/Caracteristicas', exist_ok=True
# Garante que a pasta de destino exista

for i in range(num_arquivos):
    inicio = i * max_por_arquivo
    fim = inicio + max_por_arquivo
    df_split = df.iloc[inicio:fim]

    nome_arquivo = f'{pasta_saida}\\Lista_Carga_{i+1}.xlsx'
    df_split.to_excel(nome_arquivo, index=False)
    print(f"✅ Criado: {nome_arquivo} ({len(df_split)} linhas)")

print("✔️ Finalizado!")
