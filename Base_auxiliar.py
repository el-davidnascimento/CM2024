import shutil

# Define o caminho do arquivo de origem e destino
caminho_origem = r'C:\Users\Victor\Downloads\Tratado_terceiro.xlsx'
caminho_destino = r'C:\Users\Victor\Downloads\Tratado_quarto.xlsx'

# Copia o arquivo de origem para o destino com um novo nome
shutil.copyfile(caminho_origem, caminho_destino)

