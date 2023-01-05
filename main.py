import smtplib
import sys
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from avaliaprecos import AvaliaPrecos
import pandas
import configparser


def get_valores_excel(excel_sheet):
    data_frame = (pandas.read_excel(excel_sheet)).values.tolist()
    lista_produtos = []
    for x in range(len(data_frame)):
        lista_produtos.append(data_frame[x][0])
    return lista_produtos


def envia_email(send_from, send_to, subject, text, password):
    message = MIMEMultipart()
    message['From'] = send_from
    message['To'] = send_to
    message['Subject'] = subject
    message.attach(MIMEText(text))

    part = MIMEBase('application', "octet-stream")
    part.set_payload(open("relatorio.xlsx", "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="relatorio.xlsx"')
    message.attach(part)

    try:
        smtp = smtplib.SMTP('smtp.gmail.com', 587)
        smtp.starttls()
        smtp.login(send_from, password)
        smtp.sendmail(send_from, send_to, message.as_string())
        log = open("log.txt", 'a')
        log.write(datetime.now().strftime('%d/%m/%y %H:%M:%S') +
                  " - Email enviado com sucesso!\n")
        smtp.quit()
    except Exception:
        log = open("log.txt", 'a')
        log.write(datetime.now().strftime('%d/%m/%y %H:%M:%S') +
                  "Erro {} ao enviar e-mail! \n".format(sys.exc_info()[0]))
        return


excel = "Exemplo.xlsx"
products = get_valores_excel(excel)
avaliaprecos = AvaliaPrecos()
avaliaprecos.gerador_relatorio_excel(products)

config = configparser.ConfigParser()
config.read('email.ini')

send_from = config['DEFAULT'].get('send_from')
send_to = config['DEFAULT'].get('send_to')
subject = config['DEFAULT'].get('subject')
text = config['DEFAULT'].get('text')
password = config['DEFAULT'].get('password')

envia_email(send_from, send_to, subject, text, password)
