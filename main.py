###### CODIGO PARA A CAPTAÇÃO DAS INFORMAÇÕES DE MATRICULA #############
#
# O INTUITO É PURAMENTE DE ANALISE DE DADOS

# Importar as bibliotecas necessárias
from datetime import datetime as date
import openpyxl
import pandas as pd

# Caminho para o arquivo do educacional
caminho_arquivo_educacional = r'G:\.shortcut-targets-by-id\1kArAZwgCxrjbQwQOPEzeJLtMUll3VVJ7\13. Dados\13.8. Campanha de Matrícula\13.8.2. Base\Quantidade de Alunos - 2023.xlsx'

# Ler o arquivo Excel
dados_excel_educacional = pd.read_excel(caminho_arquivo_educacional)

# Obter a data atual
data_atual = date.today()

# Criar um dicionário com os dados da tabela
planilha_educacional = {
    'CODCOLIGADA': dados_excel_educacional['CODCOLIGADA'],
    'DATAMATRICULA': [data_atual] * len(dados_excel_educacional),
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

# Filtrar as linhas em que a coluna 'CURSO' não está vazia
df_educacional = df_educacional[df_educacional['CURSO'].notnull()]

# Remover linhas duplicadas com base nas colunas especificadas
df_educacional.drop_duplicates(subset=['RA', 'NOME DO ALUNO', 'HABILITACAO', 'CODTURMA', 'TURNO', 'CURSO'], keep='first', inplace=True)

# Caminho para o arquivo tratado
Tratado_educacional = r'C:\Users\Victor\Downloads\Tratado_educacional.xlsx'

# Salvar o DataFrame no arquivo Excel
df_educacional.to_excel(Tratado_educacional, index=False)

#############################################################################################################################################

# Caminho para o arquivo do lançamento
caminho_arquivo_lancamento = r'G:\.shortcut-targets-by-id\1kArAZwgCxrjbQwQOPEzeJLtMUll3VVJ7\13. Dados\13.8. Campanha de Matrícula\2023 - Lancamento.xlsx'

# Ler o arquivo Excel
dados_excel_lancamento = pd.read_excel(caminho_arquivo_lancamento)

# Criar um dicionário com os dados da tabela
planilha_lancamento = {
    'DATA DE CRIACAO': data_atual,
    'CODCOLIGADA': dados_excel_lancamento['CODCOLIGADA'],
    'RA': dados_excel_lancamento['RA'],
    'NOME DO ALUNO': dados_excel_lancamento['ALUNO'],
    'NATUREZA_FINANCEIRA': dados_excel_lancamento['NATUREZA_FINANCEIRA'],
    'COMPETENCIA': dados_excel_lancamento['COMPETENCIA'],
    'MESES': dados_excel_lancamento['MESES'],

}

# Criar o DataFrame a partir do dicionário
df_lancamento = pd.DataFrame(planilha_lancamento)

# Remover linhas duplicadas com base nas colunas especificadas
df_lancamento.drop_duplicates(subset=['CODCOLIGADA', 'RA', 'NOME DO ALUNO', 'NATUREZA_FINANCEIRA', 'MESES'], keep='first', inplace=True)

# Filtrar as linhas em que a coluna 'CURSO' não está vazia
df_lancamento = df_lancamento[df_lancamento['MESES'].notnull()]

# Caminho para o arquivo tratado
Tratado_lancamento = r'C:\Users\Victor\Downloads\Tratado_lancamento.xlsx'

# Salvar o DataFrame no arquivo Excel
df_lancamento.to_excel(Tratado_lancamento, index=False)