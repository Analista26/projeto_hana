import pandas as pd
import math
import os

# Caminho do arquivo de entrada
arquivo_entrada = r'Z:/00 Pastas de trabalho/Asantos/08 - Lista_Carga/Lista_Carga_Cacteristicas.xlsx'

# Carregar o Excel
df = pd.read_excel(arquivo_entrada, dtype=str)  # garante leitura como string

# Garante limpeza do campo ATNAM
if 'ATNAM' in df.columns:
    df['ATNAM'] = df['ATNAM'].astype(str).str.strip()  

# Quantidade máxima por arquivo
max_por_arquivo = 10000

# Número total de splits necessários
num_arquivos = math.ceil(len(df) / max_por_arquivo)

# Pasta de saída
pasta_saida = r'Z:/00 Pastas de trabalho/Asantos/08 - Lista_Carga/Caracteristicas'
os.makedirs(pasta_saida, exist_ok=True)  # ✅ Garante que a pasta exista

# Loop para dividir e salvar os arquivos
for i in range(num_arquivos):
    inicio = i * max_por_arquivo
    fim = inicio + max_por_arquivo
    df_split = df.iloc[inicio:fim]

    nome_arquivo = os.path.join(pasta_saida, f'Lista_Carga_{i+1}.xlsx')
    df_split.to_excel(nome_arquivo, index=False)
    print(f"✅ Criado: {nome_arquivo} ({len(df_split)} linhas)")

print("✔️ Finalizado!")
