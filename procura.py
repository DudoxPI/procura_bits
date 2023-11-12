import pandas as pd

def procurar_palavra_em_arquivo_xlsx(arquivo_xlsx, palavra_a_procurar):
    # Carregue o arquivo XLSX em um DataFrame do pandas
    df = pd.read_excel(arquivo_xlsx)
    pd.set_option('display.max_colwidth', None)
    
    # Encontre as linhas onde a palavra foi encontrada
    linhas_encontradas = df[df.loc[:, df.columns != 'endereco'].apply(lambda row: row.astype(str).str.contains(palavra_a_procurar, case=False).any(), axis=1)]

    if not linhas_encontradas.empty:
        print("Palavra encontrada nas seguintes linhas:")
        print("-------------------------------------------------------------")
        print(linhas_encontradas)
    else:
        print("Palavra não encontrada no arquivo.")

# Nome do arquivo XLSX que você deseja pesquisar
arquivo_xlsx = "bits.xlsx"

# Palavra que você deseja procurar (a pesquisa é insensível a maiúsculas e minúsculas)
while True:
    palavra_a_procurar = input('Digite palavra para procurar: ')

    # Chame a função para procurar a palavra e exibir as linhas onde ela foi encontrada
    procurar_palavra_em_arquivo_xlsx(arquivo_xlsx, palavra_a_procurar)
