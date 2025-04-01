import pandas as pd

# Tente abrir o arquivo
try:
    arquivo = "CADASTRO FITAS RIO.xlsx"  # Troque para .xlsx se salvar como outro formato
    with pd.ExcelFile(arquivo, engine='openpyxl') as xls:
        sheet_names = xls.sheet_names  # Obtém o nome das abas
        print(f"Aba(s) encontrada(s) no arquivo: {sheet_names}")

        # Certifique-se de que a aba tem o nome correto
        aba_certa = "BASE FITAS LEGATO"  # Ajuste conforme necessário
        if aba_certa not in sheet_names:
            print(f"Erro: a aba '{aba_certa}' não foi encontrada no arquivo.")
        else:
            df = pd.read_excel(xls, sheet_name=aba_certa)
            print("Colunas encontradas no Excel:")
            print(df.columns)
            print("Primeiras 5 linhas do arquivo:")
            print(df.head())
except Exception as e:
    print(f"Erro ao carregar o arquivo: {e}")
