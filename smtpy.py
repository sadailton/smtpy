import os
import mimetypes
import csv
import json
import pandas as pd
import smtplib
import imaplib
from unidecode import unidecode
import urllib.parse
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from email import encoders
from email import parser
from datetime import datetime


def get_mime_type(caminho_arquivo: str) -> str:
    
    mime_type = mimetypes.guess_type(caminho_arquivo)
    
    return mime_type

def anexa_arquivo(mensagem, caminho_anexo, nome_anexo) -> None:
    mime_type = get_mime_type(caminho_anexo)
    extensao_arquivo = os.path.splitext(caminho_anexo)[1]

    if not(nome_anexo):
        # Obter somente o nome do arquivo
        nome_anexo = os.path.basename(caminho_anexo)
                       
    with open(caminho_anexo, 'rb') as arquivo:


        if mime_type and mime_type[0].startswith('text'):
            anexo = MIMEText(arquivo.read().decode('utf-8'), 'plain')
        
        elif mime_type and mime_type[0].startswith('image'):
            anexo = MIMEImage(arquivo.read())
            anexo.add_header('Content-ID', 'convite_aula_encerramento')

        elif mime_type:
            anexo = MIMEApplication(arquivo.read(), _subtype=mime_type[0].split('/')[1])

        else:
            print(f"Não foi possível determinar o tipo MIME para o arquivo { caminho_anexo }")
            return False

        # A função unidecode() retira os acentos e cedilhas de uma string
        
        if nome_anexo == '':
            nome_anexo = unidecode(nome_anexo) 
        else:
            nome_anexo = unidecode(nome_anexo) + extensao_arquivo
        
        anexo.add_header('Content-Disposition', f'attachment; filename={ nome_anexo }') 

        mensagem.attach(anexo)

def envia_email(servidor_smtp: str, porta: int, usuario: str, senha: str, \
                email_remetente: str, email_destinatario: str, emails_cc: list = [], emails_bcc: list = [], assunto: str = '', mensagem: str = '', caminho_anexo: list = [], nome_arquivo_anexo: str = '') -> bool:
    """Função que envia e-mail usando a biblioteca smtplib.

        Args:
            servidor_smtp (str): Servidor de e-mail que irá enviar o(s) e-mail(s). Pode ser passado o FQDN ou o IP.
            porta (int): Porta do servidor SMTP. Defaults para 587
            usuario (str): Login do usuário da conta do servidor de e-mail.
            senha (str): Senha do login do usuário da conta do servidor de e-mail.
            email_remetente (str):  Endereço de e-mail do remetente.
            email_destinatario (str):  Endereço de e-mail do destinatário.
            assunto (str, optional): Assunto do e-mail. Defaults to ''.
            mensagem (str, optional): Corpo da mensagem do e-mail. Pode ser em formato HTML. Defaults to ''.
            caminho_anexo (list, optional): Lista contento os arquivos a serem enviados como anexo no e-mail. Defaults to ''.
            nome_arquivo_anexo (str, optional): Nome do arquivo a ser anexado. Caso não seja informado e a variável 'caminho_anexo' estiver preenchida,
                                                será usado o nome dos arquivos. Defaults to ''.

        Returns:
            bool: Retorna True se o e-mail for enviado com sucesso e False caso contrário.
    """

    try:
        
        #Configura os parametros do e-mail
        msg = MIMEMultipart()
        msg['From'] = email_remetente
        msg['To'] = email_destinatario        
        msg["Subject"] = assunto

        for email_cc in emails_cc:
            msg['Cc'] = email_cc
        
        for email_bcc in emails_cc:
            msg['Bcc'] = email_bcc

        # Adiciona o corpo do e-mail
        msg.attach(MIMEText(mensagem, 'html', "utf-8"))
        #msg = msg.as_string().encode('ascii')

        # Verificar se há um caminho de anexo válido
        if caminho_anexo:
            for anexo in caminho_anexo:
                anexa_arquivo(msg, anexo, nome_arquivo_anexo)

        # Try to log in to server and send email
        with smtplib.SMTP(servidor_smtp, porta) as servidor:
            servidor.starttls() # Secure the connection
            servidor.login(usuario, senha)
            
            # sendmail() retorna vazio se não houve erro no envio
            
            response = servidor.sendmail(email_remetente, [email_destinatario] + emails_cc + emails_bcc, msg.as_string())
            
            servidor.quit()
            return response


    except Exception as erro:
        return erro
    
def get_received_from_mta(email_message, header='X-Failed-Recipients'):
    
    received_from_mta = None
    received_header_failed_recipients = email_message.get_all(header)
    #received_header_error_message = email_message

    #print(received_header_error_message)

    #if received_header_failed_recipients:
        #for header in received_header_failed_recipients:
            #print(header)
            #print(received_header_error_message)
            #if 'X-Failed-Recipients' in header:
            #    received_from_mta = header
                #break

    return received_header_failed_recipients

def check_bounces(email: str, senha: str, num_emails: int = 5):
    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    mail.login(email, senha)
    mail.select("INBOX")

    result, data = mail.search(None, "ALL")

    if result == "OK":
        email_ids = data[0].split()
        start_index = max(0, len(email_ids) - num_emails)
        email_ids_to_fetch = email_ids[start_index:]
        for email_id in email_ids_to_fetch:
            result, data = mail.fetch(email_id, "(RFC822)")
            if result == "OK":
                raw_email = data[0][1]
                email_message = parser.BytesParser().parsebytes(raw_email)
                
                destinatario = get_received_from_mta(email_message)
                assunto = get_received_from_mta(email_message, header='Subject')
                email_body = get_email_body(email_message)

                log_failed_email_delivery_from_inbox(destinatario, assunto, email_body)

    mail.close()
    mail.logout()

def get_email_body(email_message):
    if email_message.is_multipart():
        # Caso o e-mail seja multipart (possui várias partes), procuramos pelo texto plano
        for part in email_message.walk():
            if part.get_content_type() == 'text/plain':
                return part.get_payload(decode=True).decode('utf-8')
    else:
        # Caso contrário, retornamos diretamente o corpo do e-mail
        return email_message.get_payload(decode=True).decode('utf-8')


def log_failed_email_delivery_from_inbox(destinatario, assunto, email_body):
   cabecalho = ['Destinatario', 'Assunto', 'Mensagem']
   arquivo_log = 'arquivos\\delivery_failed.csv'
   objeto = [destinatario, assunto, email_body]

   grava_csv(cabecalho, objeto, arquivo_log)


def registrar_status_envio(dados: dict, arquivo_log: str = 'envio_email_log.csv') -> None:
    """Registro de log do envio de e-mail.

    Args:
        dados (dict): Dicionario contendo os campos do arquivo de log.
        Exemplo de dicionario: 
        log_info: dict = {'data': '',
                    'hora': '',
                    'nome': '',
                    'destinatario': '',
                    'status': ''
                    }
        arquivo_log (str, optional): Nome do arquivo de log. Defaults to 'envio_email_log.csv'.
    """
    agora = datetime.now()
    data = agora.strftime("%d-%m-%Y")
    hora = agora.strftime("%H:%M:%S")

    dados["data"] = data
    dados["hora"] = hora

    json_to_csv(dados, f"logs/{arquivo_log}")


'''
def registrar_status_envio(cabecalho: list, dados: list, flag: bool, arquivo: str):
    
    agora = datetime.now()
    data = agora.strftime("%d-%m-%Y")
    hora = agora.strftime("%H:%M:%S")
    
    status = ""

    if flag:
        status = "Enviado"
    else:
        status = "Falha"

    cabecalho = ['Data', 'Hora', 'Remetente', 'Destinatario', 'Status']
    log = [data, hora, remetente, destinatario, status]
    grava_csv(cabecalho, log, arquivo)
'''

def json_to_csv(json_data, csv_filename):
    try:
        # Load existing CSV file to check if it's empty
        existing_df = pd.read_csv(csv_filename)
        header_exists = not existing_df.empty
    except (FileNotFoundError, pd.errors.EmptyDataError):
        header_exists = False

    with open(csv_filename, 'a', newline='', encoding='utf-8') as csv_file:
        # Only write header if it doesn't exist
        if not header_exists:
            df = pd.json_normalize(json_data)
            df.to_csv(csv_file, index=False, header=True)
        else:
            df = pd.json_normalize(json_data)
            df.to_csv(csv_file, index=False, header=False, mode='a')


def grava_csv(cabecalho: list = [], objeto: list = [], caminho_arquivo: str = '') -> bool:

    try:
        arquivo_existe = os.path.isfile(caminho_arquivo)

        with open(caminho_arquivo, "a", encoding='utf-8', newline='') as arquivo:
            writer = csv.writer(arquivo)

            # Verificar se o arquivo está vazio
            arquivo_vazio = arquivo.tell() == 0

            if not(arquivo_existe) and arquivo_vazio:
                writer.writerow(cabecalho)
            
            writer.writerow(objeto)
        return True
    
    except FileNotFoundError as erro:
        print("Arquivo não encontrado", erro)
        return False
    except Exception as erro:
        print(erro)
        return False
