import pandas as pd
import os
import webbrowser
from unidecode import unidecode  # Importe a função unidecode

caminho_arquivo = "bits.xlsx" # Especifique o caminho para o seu arquivo XLSX
df = pd.read_excel(caminho_arquivo) # Leia o arquivo XLSX em um DataFrame do Pandas
pd.set_option('display.max_colwidth', None)#aumenta tamanho da coluna link
pd.set_option('display.max_columns', None)#mostra todas as colunas
msg = ''

while True:
    print(msg)
    valor_procurado = input('Digite palavra para busca em Sumários: ') # Valor que você deseja procurar na primeira coluna
    os.system('cls')
        
    if valor_procurado == '':
        msg = '*** Campo não pode ser vazio!'
        os.system('cls')
        continue
    
    # Realize a pesquisa na primeira coluna usando unidecode para normalização
    palavras_chave = unidecode(valor_procurado).lower().split()  # Divide as palavras da consulta
    resultados = df[df.iloc[:, 0].apply(lambda x: all(word in unidecode(x).lower() for word in palavras_chave))]
    #resultados = df[df.iloc[:, 0].apply(lambda x: unidecode(x).lower()).str.contains(unidecode(valor_procurado).lower())]
    #sem_coluna = (resultados.drop(columns=['url'])) # retira coluna endereço do resultado

    # Verifica se foram encontrados resultados
    if not resultados.empty:
        if resultados.shape[0] > 1:
            print(f'*** Foram encontrados [ {resultados.shape[0]} ] resultados para: -> {valor_procurado}')
        else:
            print(f"*** Foi encontrado [ {resultados.shape[0]} ] resultado para: -> {valor_procurado}")
        print('')
        print('')

        for index, row in resultados.iterrows():
            print(f"---> {row['Sumario']}.......... Página: {row['Pagina']}")
            print('')
            print(f"     Cliente: {row['Cliente']}")
            print('')
            print(f"     BIT: {row['BIT']} - {row['Titulo']}") 
            print('')
            print(f"     Link: {row['Link']}\n")
            #print('-------------------------------------------------------------------------------')
            msg = ''
    else:
        msg = (f"*** Nenhum resultado encontrada para: ->  {valor_procurado}")
        os.system('cls')
        continue

    print('*** Clique duas vezes no link, copie e cole no navegador!')
    

