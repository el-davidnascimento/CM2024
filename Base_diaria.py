###### CODIGO PARA CAPTAR AS INFORMAÇÕES DA PLANILHA DE QUANTIDADE DE ALUNOS #############

# Importar as bibliotecas necessárias
import pandas as pd

# Carrega os DataFrames dos arquivos Excel
tratado_terceiro = pd.read_excel(r'C:\Users\Victor\Downloads\Tratado_terceiro.xlsx')
tratado_quarto = pd.read_excel(r'C:\Users\Victor\Downloads\Tratado_quarto.xlsx')

# Encontra os nomes que existem no Tratado_terceiro e não no Tratado_quarto
nomes_a_serem_adicionados = tratado_terceiro[~tratado_terceiro['NOME DO ALUNO_x'].isin(tratado_quarto['NOME DO ALUNO_x'])]

# Salva os nomes em um quinto arquivo
quinto_arquivo = r'C:\Users\Victor\Downloads\Tratado_quinto.xlsx'
nomes_a_serem_adicionados.to_excel(quinto_arquivo, index=False)