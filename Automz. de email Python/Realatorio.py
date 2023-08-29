import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# Importar base de dados
tabela_medias = pd.read_excel('Faculdade (version 1).xlsb')

# Visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_medias)

# Criar nova coluna para mensurar o número de repetições
tabela_medias['nova_coluna'] = [1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1]
mediador  = tabela_medias[['Faculdade', 'nova_coluna']].groupby('Faculdade').sum()
mediador2 = tabela_medias[['Estado', 'nova_coluna']].groupby('Estado').sum()

# Notas por faculdade
print(100*'_')
notas = tabela_medias[['Faculdade','Estado', 'Media']]
print(notas)

# Média nacional das faculdades
print(100*'_')
media_nacional = tabela_medias[['Faculdade', 'Media']].groupby('Faculdade').sum()
media_nacional['Media'] = media_nacional['Media'] / mediador['nova_coluna']
media_nacional['Media'] = round(media_nacional['Media'], 1)
print(media_nacional)

# Média dos estados
print(100*'_')
media_estados = tabela_medias[['Estado', 'Media']].groupby('Estado').sum()
media_estados['Media'] = media_estados['Media'] / mediador2['nova_coluna']
media_estados['Media'] = round(media_estados['Media'], 1)
print(media_estados)

# Configurações do e-mail
smtp_server = 'exemplo@.com' #Digite seu smtp server(servidor de protocolo de email, você encontrarnas configurações do seu email)
smtp_port = 123 #Digite sua porta de segurança smtp (você encontrarnas configurações do seu email)
sender_email = input('Digite o seu email: ')
sender_password = input('Digite o sua senha: ')
receiver_email = input('Qual o email do destinatário: ')

# Criar a representação HTML da tabela media_nacional
media_nacional = media_nacional.to_html(classes='table', escape=False)

# Criar o conteúdo do e-mail
message = MIMEMultipart()
message['From'] = sender_email
message['To'] = receiver_email
message['Subject'] = 'Relatório de médias por faculdade e estados'

# Criar o conteúdo do e-mail
message = MIMEMultipart()
message['From'] = sender_email
message['To'] = receiver_email
message['Subject'] = 'Relatório de médias por faculdade e estados'

# Texto do e-mail
html_content = f'''
<!DOCTYPE html>
<html>
<head>
<style>
    table {{
        border-collapse: collapse;
        width: 100%;
        font-family: Arial, sans-serif;
    }}
    th, td {{
        border: 1px solid #dddddd;
        text-align: left;
        padding: 8px;
    }}
    th {{
        background-color: #f2f2f2;
    }}
</style>
</head>
<body>
<p>Prezados,</p>
<p>Segue abaixo um relatório mostrando a média das faculdades e o desempenho médio dos alunos por estado e faculdade.</p>

<h3>As notas por faculdade:</h3>
{notas.to_html(classes="table", escape=False)}

<h3>Média nacional das faculdades:</h3>
{media_nacional}

<h3>Média dos estados:</h3>
{media_estados.to_html(classes="table", escape=False)}

<p>Dúvidas? Estou sempre à disposição &lt;3</p>
<p>Att..<br>Antônio Lucas</p>
</body>
</html>
'''

# Anexar o conteúdo HTML ao e-mail
message.attach(MIMEText(html_content, 'html'))

# Autenticação e envio do e-mail...

# Autenticação e envio do e-mail
try:
    smtp = smtplib.SMTP(smtp_server, smtp_port)
    smtp.starttls()
    smtp.login(sender_email, sender_password)
    smtp.sendmail(sender_email, receiver_email, message.as_string())
    smtp.quit()
    print('E-mail enviado com sucesso!')
except Exception as e:
    print('Erro ao enviar o e-mail:', e)