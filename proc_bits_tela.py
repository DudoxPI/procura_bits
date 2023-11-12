import tkinter as tk
import pandas as pd
from unidecode import unidecode

def pesquisar(event=None):
    valor_procurado = entrada_busca.get()
    resultado_text.delete(1.0, tk.END)

    if valor_procurado == '':
        resultado_text.insert(tk.END, '*** Campo não pode ser vazio!')
        return

    palavras_chave = unidecode(valor_procurado).lower().split()  # Divide as palavras da consulta
    resultados = df[df.iloc[:, 0].apply(lambda x: all(word in unidecode(x).lower() for word in palavras_chave))]

    if not resultados.empty:
        resultado_text.insert(tk.END, f'*** Foram encontrados {resultados.shape[0]} resultados para: -> {valor_procurado}\n\n')
        for index, row in resultados.iterrows():
            resultado_text.insert(tk.END, f"---> {row['Sumario']}.............. Página: {row['Pagina']}\n")
            resultado_text.insert(tk.END, f"     Cliente: {row['Cliente']}\n")
            resultado_text.insert(tk.END, f"     BIT: {row['BIT']} - {row['Titulo']}\n")
            resultado_text.insert(tk.END, f"     Link: {row['Link']}\n\n\n")
    else:
        resultado_text.insert(tk.END, f'*** Nenhum resultado encontrada para: ->  {valor_procurado}')

# Leitura do arquivo Excel
caminho_arquivo = "bits.xlsx"
df = pd.read_excel(caminho_arquivo)

# Configuração da janela
root = tk.Tk()
root.title('Pesquisa de Sumários')
root.geometry('1024x768')

# Widget da entrada e botão de pesquisa
frame_busca = tk.Frame(root)
frame_busca.pack(pady=10)

label_busca = tk.Label(frame_busca, text='Buscar em Sumários:')
label_busca.pack(side=tk.LEFT)

entrada_busca = tk.Entry(frame_busca, width=40)
entrada_busca.pack(side=tk.LEFT, padx=10)
entrada_busca.bind('<Return>', pesquisar)

botao_pesquisar = tk.Button(frame_busca, text='Pesquisar', command=pesquisar)
botao_pesquisar.pack(side=tk.LEFT)

# Widget de resultados
resultado_text = tk.Text(root, wrap=tk.WORD)
resultado_text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

# Loop principal
root.mainloop()



#pyinstaller -w -F  procura3.py