import pandas as pd

arquivo = 'F:/BackOffice_GERAL/Projeto S4 Hana/Onda 2/9. Saneamento/Asantos/25 . Saneamento/FORNECEDOR/EXPORT_TO_S4-ADRC_0001.txt'

# Tente separadores comuns
separadores = ['\t', ';', ',', '|']

for sep in separadores:
    try:
        df = pd.read_csv(arquivo, sep=sep, nrows=5)
        print(f'\n✅ Separador "{sep}" lido com sucesso. Colunas detectadas:')
        print(df.columns)
    except Exception as e:
        print(f'\n❌ Erro com separador "{sep}": {e}')
