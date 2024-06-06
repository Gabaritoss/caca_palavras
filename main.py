# importando bibliotecas

import openpyxl # Necessaria para o pandas ler arquivos excel
import pandas as pd # Serve para manipular, importar e exportar data sets
import numpy as np # Adiciona e melhora funções e operações matematicas e estatisticas
import re # Serve para manipular strings

# criação das variaveis

lista_planilhas = []
arquivo_usuario = []
novo_arquivo = 0

# armazenar nome dos arquivos

while novo_arquivo != "":
  novo_arquivo = input("insira o nome das planilhas ou aperte enter para finalizar: ")
  arquivo_usuario.append(f"{novo_arquivo}")
  print(f"{novo_arquivo} adicionado.")
  
arquivo_usuario.remove("")
print(f"lista de planilhas carregadas: {arquivo_usuario}")
lista_planilhas += list(arquivo_usuario)

# Executa algoritmo na lista de planilhas

for planilha in lista_planilhas:
    
    # Armazenando as planilhas em um data frame

    ## IMPORTANTE: colocar o caminho completo da planilha

    try:
        data_frame = pd.read_excel(f"C:/Users/Gabriel/Desktop/caça_palavras/arquivos_iniciais/{planilha}.xlsx")
        
    except FileNotFoundError:
        print(f"arquivo {planilha} não encontrado, revise o nome do arquivo e garanta que esta em formato .xlsx.")
    
    else:

        # Determina a coluna que o arquivo vai ler
      
        frases = data_frame['DESCRIPCION'].apply(str)
        
        # limpeza do descritivo

        frases_sem_numeros = [re.sub(r'[0-9]', '', frase) for frase in frases]
        frases_sem_numeros_maiusculas = [frase.upper() for frase in frases_sem_numeros]
        frases_sem_kg_fat = [re.sub(r'\bKG\b|\bFAT\b|\bDE\b|[.,-:;%+&\*()]|\b\w\w\b', ' ', frase) for frase in frases_sem_numeros_maiusculas]
        frases_sem_palavras_curtas = [re.sub(r'\b\w\b', '', frase) for frase in frases_sem_kg_fat]

        # Dicionário para armazenar palavras repetidas

        palavras_repetidas = {}
        for frase_sem_palavra_curta in frases_sem_palavras_curtas:
            for palavra in frase_sem_palavra_curta.split():
                if palavra in palavras_repetidas:
                    palavras_repetidas[palavra] += 1
                else:
                    palavras_repetidas[palavra] = 1

        # Criação do data frame

        df_palavras_repetidas = pd.DataFrame.from_dict(palavras_repetidas, orient='index', columns=['Contagem'])
        df_palavras_repetidas.reset_index(inplace=True)
        df_palavras_repetidas.columns = ['Palavra', 'Contagem']
        
        # Criação do arquivo excel

        ## IMPORTANTE: colocar o caminho completo da planilha

        df_palavras_repetidas.to_excel(f'C:/Users/Gabriel/Desktop/caça_palavras/arquivos_finais/{planilha}_final.xlsx', index=False)
        
        print(f"arquivo {planilha}_final criado com sucesso")
