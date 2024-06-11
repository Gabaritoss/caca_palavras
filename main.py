# importando bibliotecas

import openpyxl # Necessaria para o pandas ler arquivos excel
import pandas as pd # Serve para manipular, importar e exportar data sets
import numpy as np # Adiciona e melhora funções e operações matematicas e estatisticas
import re # Serve para manipular strings

# criação das variaveis

lista_planilhas = []
arquivo_usuario = []
novo_arquivo = 0

# Função caca palavras

def caça_palavras(data_frame, planilha):
    """
    Função para analisar um DataFrame de descrições, 
    limpar os textos, contar as palavras repetidas e 
    salvar os resultados em um arquivo Excel.

    Argumentos:
        data_frame (pandas.DataFrame): DataFrame com a coluna 'DESCRIPCION'.
        planilha (str): Nome do arquivo de saída.
    """

    frases = data_frame['DESCRIPCION'].apply(str) # Armazena a coluna descripcion na variavel frases

    # Limpeza do descritivo

    frases_sem_numeros = [re.sub(r'[0-9]', '', frase) for frase in frases] # remove numeros
    frases_sem_numeros_maiusculas = [frase.upper() for frase in frases_sem_numeros] # converte em maiusculo
    frases_sem_kg_fat = [re.sub(r'\bKG\b|\bFAT\b|\bDE\b|[.,-:;%+&\*()]', ' ', frase) for frase in frases_sem_numeros_maiusculas] # remove kg e fat
    frases_limpas = [re.sub(r'\b\w\b|\b\w\w\b', '', frase) for frase in frases_sem_kg_fat] # remove palavras com 2 ou menos caracteres

    # Dicionário para armazenar palavras repetidas

    palavras_repetidas = {}
    for frases_limpas in frases_limpas:
        for palavra in frases_limpas.split():
            if palavra in palavras_repetidas:
                palavras_repetidas[palavra] += 1
            else:
                palavras_repetidas[palavra] = 1

    # Criação do data frame

    df_palavras_repetidas = pd.DataFrame.from_dict(palavras_repetidas, orient='index', columns=['Contagem'])
    df_palavras_repetidas.reset_index(inplace=True)
    df_palavras_repetidas.columns = ['Palavra', 'Contagem']
    
    # Cálculos

    Contagem = df_palavras_repetidas["Contagem"]

    df_palavras_repetidas['Share'] = Contagem / Contagem.sum() # Calcula o share para cada palavra

    # Criação do arquivo excel

    ## IMPORTANTE: colocar o caminho completo da planilha
    df_palavras_repetidas.to_excel(f'C:/Users/Gabriel/Desktop/caça_palavras/arquivos_finais/{planilha}_final.xlsx', index=False)
    print(f"arquivo {planilha}_final criado com sucesso")

# armazenar nome dos arquivos determinados pelo usuario

while novo_arquivo != "":
  novo_arquivo = input("insira o nome das planilhas ou aperte enter para finalizar: ") # Solicita os arquivos
  arquivo_usuario.append(f"{novo_arquivo}") # Adiciona eles em uma lista
  print(f"{novo_arquivo} adicionado.")
  
arquivo_usuario.remove("") # Remove o espaço em branco do enter, condição para sair do loop while
print(f"lista de planilhas carregadas: {arquivo_usuario}") # Retorna a lista de arquivos carregados
lista_planilhas += list(arquivo_usuario) # Adiciona os arquivos do usuario a lista de planilhas

# Executa algoritmo na lista de planilhas

for planilha in lista_planilhas:
    
    # Armazenando as planilhas em um data frame

    ## IMPORTANTE: colocar o caminho completo da planilha

    try:
        data_frame = pd.read_excel(f"C:/Users/Gabriel/Desktop/caça_palavras/arquivos_iniciais/{planilha}.xlsx") # Le o arquivo excel do usuario
        
    except FileNotFoundError:
        print(f"arquivo {planilha} não encontrado, revise o nome do arquivo e garanta que esta em formato .xlsx.") # Sinaliza caso haja algum erro na etapa anterior
    
    else:
        caça_palavras(data_frame, planilha) # Executa a função caça_palavras com os arquivos do usuario.