import smtplib
import codecs
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header

# Constantes que serão usadas no envio do e-mail
# arquivos csv com lista de todos destinatários e seus respectivos e-mails
csvDestinatarios = 'destinatarios.csv'
# remetente fake
emailRemetente = 'remetente@email.com'
# senha fake
senhaEmailRemetente = 'senhaForte@123456'
templateHTML = 'template.html'
# endereço de correio fake
enderecoCorreio = 'correio.email.com'
csvRepasse = 'repasses.csv'
today = datetime.date.today()
date = today.strftime('%d/%m/%Y')
# banco fake
banco = 'Banco de Testes'
# agência fake
agencia = '1234'
# conta fake
conta = '98765-2'
# empresa fake
empresa = 'Empresa de Teste'
# cnpj fake gerado no https://www.4devs.com.br/gerador_de_cnpj
cnpj = '27.691.892/0001-70'
assuntoEmail = 'Relatório de Repasses'
# email fake
emailFrom = 'pendencias@email.com'

# abre o arquivos dos destinatários
with open(csvDestinatarios) as f:
    dados = f.read().splitlines()
f.close()
dados.pop(0)
for l in dados:
    l.strip('\n\r')

destinatarios = {}
for dado in dados:
    temp = dado.split(';')
    destinatarios[temp[0].strip().decode('utf-8')] = temp[1].strip().split('/')

smtp = smtplib.SMTP(enderecoCorreio, 587)
smtp.starttls()
smtp.login(emailRemetente, senhaEmailLogin)

with codecs.open(templateHTML, 'r', 'utf-8-sig') as f:
    html = f.read()

with open(csvRepasse) as f:
    dados = f.read().splitlines()
f.close()
dados.pop(0)
for l in dados:
    l.strip('\n\r')

for dado in dados:
    temp = dado.split(';')
    nomeDestinatario = temp[0].strip().decode('utf-8')
    valor = temp[1]
    email = html.replace('{{{cartorio}}}', nomeDestinatario)
    email = email.replace('{{{valor}}}', valor)
    email = email.replace('{{{data}}}', date)
    email = email.replace('{{{banco}}}', banco)
    email = email.replace('{{{agencia}}}', agencia)
    email = email.replace('{{{conta}}}', conta)
    email = email.replace('{{{empresa}}}', empresa)
    email = email.replace('{{{cnpj}}}', cnpj)
    to = destinatarios[nomeDestinatario]
    to.append(emailRemetente)

    msg = MIMEMultipart('alternative')
    msg['Subject'] = assuntoEmail
    msg['From'] = emailRemetente
    msg['To'] = ','.join(to)
    emailHtml = MIMEText(email.encode('utf-8'), 'html', 'utf-8')
    msg.attach(emailHtml)

    smtp.sendmail(emailFrom, to, msg.as_string())
smtp.quit()
