import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText

def configurar_server(server_email, server_senha):
    try:
        server = smtplib.SMTP("endereço smtp aqui", 123)
        server.starttls()
        server.login(server_email, server_senha)
        return server
    except Exception as e:
        print(f"Erro ao configurar o servidor SMTP: {e}")
        return None

def criar_email(origem_email, lista_email, assunto_email, corpo_email, arquivo):
    msg = MIMEMultipart()
    msg['From'] = origem_email
    msg['To'] = lista_email
    msg['Subject'] = assunto_email
    msg.attach(MIMEText(corpo_email, 'plain'))

    try:
        with open(arquivo, 'rb') as anexar_arquivo:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(anexar_arquivo.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={arquivo}')
            msg.attach(part)
    except Exception as e:
        print(f"Erro ao anexar o arquiv o: {e}")
    return msg

def enviar_email(server, msg):
    try:
        server.send_message(msg)
        print("E-mail enviado com sucesso")
    except Exception as e:
        print(f"Erro ao enviar o e-mail: {e}")

def mandar_email ():
    server_email = "email.exemplo@email.com"
    server_senha = "coloque a senha aqui!"
    origem_email = server_email
    assunto_email = "coloque o assunto do email aqui!"
    corpo_email = "coloque o corpo email aqui!"
    lista_email = ["email_1@email.com", "email_2@email.com"]
    lista_arquivos = ["planilhas/planilha_1.xlsx", "planilhas/planilha_2.xlsx"]

    server = configurar_server(server_email, server_senha)
    if server is None:
        return
    for destinatario, arquivo in zip(lista_email, lista_arquivos):
       msg = criar_email(origem_email, destinatario, assunto_email, corpo_email, arquivo)
       enviar_email(server, msg)
       
    server.quit()

mandar_email()
