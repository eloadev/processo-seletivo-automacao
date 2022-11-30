import smtplib
import sys
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from avaliaprecos import AvaliaPrecos
import pandas


def get_valores_excel(excel_sheet):
    data_frame = (pandas.read_excel(excel_sheet)).values.tolist()
    lista_produtos = []
    for x in range(len(data_frame)):
        lista_produtos.append(data_frame[x][0])
    return lista_produtos


def envia_email(send_from, send_to, subject, text, senha):
    mensagem = MIMEMultipart()
    mensagem['From'] = send_from
    mensagem['To'] = send_to
    mensagem['Subject'] = subject
    mensagem.attach(MIMEText(text))

    part = MIMEBase('application', "octet-stream")
    part.set_payload(open("relatorio.xlsx", "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="relatorio.xlsx"')
    mensagem.attach(part)
    try:
        smtp = smtplib.SMTP('smtp.gmail.com', 587)
        smtp.starttls()
        smtp.login(send_from, senha)
        smtp.sendmail(send_from, send_to, mensagem.as_string())
    except Exception:
        log = open("log.txt", 'a')
        log.write(datetime.now().strftime('%d/%m/%y %H:%M:%S') +
                  "Erro {} ao enviar e-mail!".format(sys.exc_info()[0]))
    finally:
        log = open("log.txt", 'a')
        log.write(datetime.now().strftime('%d/%m/%y %H:%M:%S') +
                  " - Email enviado com sucesso!")
    smtp.quit()


excel = "Exemplo.xlsx"
produtos = get_valores_excel(excel)
avaliaprecos = AvaliaPrecos()
avaliaprecos.gerador_relatorio_excel(produtos)

send_from = "elofakes@gmail.com"
send_to = "eloamello126@gmail.com"
subject = "Relatório"
text = "Relatório em anexo"
senha = "jsrbsgfxwhgcbsfu"

envia_email(send_from, send_to, subject, text, senha)