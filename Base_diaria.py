###### CODIGO PARA CAPTAR AS INFORMAÇÕES DA PLANILHA DE QUANTIDADE DE ALUNOS #############

# Importar as bibliotecas necessárias
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


# Carrega os DataFrames dos arquivos Excel
tratado_terceiro = pd.read_excel(r'C:\Users\Victor\Downloads\Tratado_terceiro.xlsx')
tratado_quarto = pd.read_excel(r'C:\Users\Victor\Downloads\Tratado_quarto.xlsx')

# Encontra os nomes que existem no Tratado_terceiro e não no Tratado_quarto
nomes_a_serem_adicionados = tratado_terceiro[~tratado_terceiro['NOME DO ALUNO_x'].isin(tratado_quarto['NOME DO ALUNO_x'])]

# Salva os nomes em um quinto arquivo
quinto_arquivo = r'C:\Users\Victor\Downloads\Tratado_quinto.xlsx'
nomes_a_serem_adicionados.to_excel(quinto_arquivo, index=False)

# Enviar para o email
# Configurações de e-mail
email_de = 'bdmultiverso@gmail.com'  # Seu e-mail
senha = 'Aristoteles@123'  # Sua senha
email_para = 'david.nascimento@multiversoeducacao.com.br'  # E-mail do destinatário
assunto = 'Testando o envio da tabela por email'
mensagem = 'Eaiii cara, será que deu certo enviar?.'

# Cria a mensagem de e-mail
msg = MIMEMultipart()
msg['From'] = email_de
msg['To'] = email_para
msg['Subject'] = assunto

# Adiciona o corpo da mensagem
msg.attach(MIMEText(mensagem, 'plain'))

# Adiciona o arquivo como anexo
with open(quinto_arquivo, "rb") as arquivo:
    part = MIMEApplication(arquivo.read(), Name="Tratado_quinto.xlsx")

part['Content-Disposition'] = f'attachment; filename="{quinto_arquivo}"'
msg.attach(part)

# Configura o servidor SMTP
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()

# Efetua login no servidor
server.login(email_de, senha)

# Envia o e-mail
server.sendmail(email_de, email_para, msg.as_string())

# Fecha a conexão com o servidor SMTP
server.quit()