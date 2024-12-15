import yaml
import csv
from smtpy  import *

def load_smtp_config(config_file: str) -> object:
    
    with open(config_file, 'r') as file:
        smtp_config = yaml.safe_load(file)
    
    return smtp_config

def main():
    
    smtp_config = load_smtp_config('email_config.yml')
    lista_contatos: str = smtp_config['lista_contatos']
    
    smtp_server: str = smtp_config['servidor_smtp']
    smtp_porta: int = smtp_config['porta']
    smtp_usuario: str = smtp_config['usuario']
    smtp_senha: str = smtp_config['senha']
    smtp_remetente: str = smtp_config['remetente']
    smtp_log: str = smtp_config['arquivo_log']
    
    email_assunto: str = ""
    email_corpo : str = ""
    
    anexo: list = ['']
    
    log_info: dict = {'data': '',
                    'hora': '',
                    'nome': '',
                    'destinatario': '',
                    'status': ''
                    }
    
    with open(lista_contatos, 'r', newline='', encoding='utf8') as csvfile:
        conteudo_csv = csv.DictReader(csvfile)
        
        for linha in conteudo_csv:
            nome = linha['nome']
            sobrenome = linha['sobrenome']
            destinatario = linha['email']
            
            # Corpo do e-mail em html
            email_corpo: str = f"""<p>Olá {nome},</p>
                                <p>Este é um email de teste.</p>
                                
                                <p>Atenciosamente,<br>
                                Equipe de teste SMTPy 
                                """
            
            log_info['nome'] = f"{nome} {sobrenome}"
            log_info['destinatario'] = f"{destinatario}"
            
            resposta = envia_email(smtp_server, smtp_porta, smtp_usuario, smtp_senha, smtp_remetente, destinatario, [], [], email_assunto, email_corpo, anexo, "arquivo_json")
            
            if  resposta == {}:
                log_info['status'] = "Sucesso"
            else:
                log_info['status'] = resposta
                
            registrar_status_envio(log_info, smtp_log)

if __name__ == '__main__':
    main()