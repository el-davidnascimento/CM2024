###### CODIGO PARA CAPTAR AS INFORMAÇÕES DA PLANILHA DE QUANTIDADE DE ALUNOS #############

# Importar as bibliotecas necessárias
from datetime import datetime as date, timedelta
import openpyxl
import pandas as pd

####################### ARQUIVO REGULAR E INTEGRAL ###########################################

# Caminho para o arquivo do educacional
caminho_arquivo_regeducacional = r'G:\.shortcut-targets-by-id\1kArAZwgCxrjbQwQOPEzeJLtMUll3VVJ7\13. Dados\13.8. Campanha de Matrícula\13.8.1. Auxiliar\REGULAR E INTEGRAL.xlsx'

# Ler o arquivo Excel
dados_excel_regeducacional = pd.read_excel(caminho_arquivo_regeducacional)

# Obter a data atual
data_atual = date.today()

# Criar um dicionário com os dados da tabela
planilha_regeducacional = {
    'PLETIVO': dados_excel_regeducacional['PLETIVO'],
    'RA': dados_excel_regeducacional['RA'],
    'NOME DO ALUNO': dados_excel_regeducacional['ALUNO'],
    'HABILITACAO': dados_excel_regeducacional['HABILITACAO'],
    'CODTURMA': dados_excel_regeducacional['SITUACAO'],
    'CURSO': dados_excel_regeducacional['CURSO']

}

# Criar o DataFrame a partir do dicionário
df_regeducacional = pd.DataFrame(planilha_regeducacional)

# Concatenar as colunas "RA", "PLETIVO" e "HABILITACAO" em ambos os DataFrames
df_regeducacional['DATA_INSERT'] = data_atual
df_regeducacional['CHAVE'] = df_regeducacional['RA'].astype(str) + '&' + df_regeducacional['PLETIVO'].astype(str) + '&' + df_regeducacional['HABILITACAO'].astype(str)

# Filtrar as linhas em que a coluna 'CURSO' não está vazia
df_regeducacional = df_regeducacional[df_regeducacional['CODTURMA'].notnull()]

# Caminho para o arquivo tratado
Tratado_regeducacional = r'C:\Users\Victor\Downloads\Tratado_regeducacional.xlsx'

# Salvar o DataFrame no arquivo Excel
df_regeducacional.to_excel(Tratado_regeducacional, index=False)


####################### ARQUIVO DO EDUCACIONAL ###########################################

# Caminho para o arquivo do educacional
caminho_arquivo_educacional = r'G:\.shortcut-targets-by-id\1kArAZwgCxrjbQwQOPEzeJLtMUll3VVJ7\13. Dados\13.8. Campanha de Matrícula\13.8.2. Base\Quantidade de Alunos - 2024.xlsx'

# Ler o arquivo Excel
dados_excel_educacional = pd.read_excel(caminho_arquivo_educacional)

# Obter a data atual
data_atual = date.today()

# Criar um dicionário com os dados da tabela
planilha_educacional = {
    'CODCOLIGADA': dados_excel_educacional['CODCOLIGADA'],
    'DATAMATRICULA': dados_excel_educacional['DATAMATRICULA'],
    'PLETIVO': dados_excel_educacional['PLETIVO'],
    'MATREMAT': dados_excel_educacional['MATREMAT'],
    'RA': dados_excel_educacional['RA'],
    'NOME DO ALUNO': dados_excel_educacional['ALUNO'],
    'HABILITACAO': dados_excel_educacional['HABILITACAO'],
    'CODTURMA': dados_excel_educacional['CODTURMA'],
    'TURNO': dados_excel_educacional['TURNO'],
    'CURSO': dados_excel_educacional['CURSO']
}

# Criar o DataFrame a partir do dicionário
df_educacional = pd.DataFrame(planilha_educacional)

# Concatenar as colunas "RA", "PLETIVO" e "HABILITACAO" em ambos os DataFrames
df_educacional['DATA_INSERT'] = data_atual
df_educacional['CHAVE'] = df_educacional['RA'].astype(str) + '&' + df_educacional['PLETIVO'].astype(str) + '&' + df_educacional['HABILITACAO'].astype(str)

# Filtrar as linhas em que a coluna 'CURSO' não está vazia
df_educacional = df_educacional[df_educacional['CURSO'].notnull()]

# Remover linhas duplicadas com base nas colunas especificadas
df_educacional.drop_duplicates(subset=['RA', 'NOME DO ALUNO', 'HABILITACAO', 'CODTURMA', 'TURNO', 'CURSO'], keep='first', inplace=True)

# Caminho para o arquivo tratado
Tratado_educacional = r'C:\Users\Victor\Downloads\Tratado_educacional.xlsx'

# Salvar o DataFrame no arquivo Excel
df_educacional.to_excel(Tratado_educacional, index=False)

####################### TERCEIRO DATA FRAME #########################
# Realizar a junção (merge) dos DataFrames usando a coluna chave
df_terceiro = pd.merge(df_regeducacional, df_educacional, on='CHAVE', how='inner')

# Caminho para o arquivo tratado do terceiro DataFrame
Tratado_terceiro = r'C:\Users\Victor\Downloads\Tratado_terceiro.xlsx'

# Salvar o terceiro DataFrame no arquivo Excel
df_terceiro.to_excel(Tratado_terceiro, index=False)
