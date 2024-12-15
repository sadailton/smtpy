# Projeto de mala direta

**Autor:** Adailton Saraiva
**Python:** 3.10.x

#### Instalando as dependencias

[!NOTE]
É altamente recomendado executar este projeto em um ambiente virtual.

```
pip install -r requirements.txt
```

#### Arquivo de configuração

O arquivo `email_config.yml` contém a configuração do servidor de e-mail utilizado para o disparo de e-mails. Por padrão o envio de e-mail é feito usando a porta 587 que é o smtp seguro.

[!WARNING]
Essa biblioteca não suporta envio de e-mail usando smtp sem TLS (porta 25).

```yml
# Configuração do servidor SMTP
servidor_smtp: "smtp.dominio.com"
porta: "587"
usuario: "fulano"
senha: "asd@123"
remetente: "no-reply@dominio.com"


# Arquivo CSV contendo a lista de destinatários
# O programa espera um arquivo contento os campos: nome, sobrenome, email
lista_contatos: "contatos.csv"

# Arquivo de log
arquivo_log: "envio_email_log.csv"

```

#### Arquivo de contatos

O programa espera um arquivo csv contendo a lista de destinatários no formato abaixo:

```csv
nome,sobrenome,email
Fulano,de Tal,fulano.tal@email.com
Ciclano,de Tal,ciclano.tal@email.com
```

#### Arquivo de log

Os logs de envio serão salvos no arquivo definido na variável `arquivo_log: ` no arquivo de configuração. Se a variável não for definida, será gerado um arquivo com o nome de `envio_email_log.csv` por padrão.

[!NOTE]
O arquivo de log é salvo dentro da pasta `logs`.

Abaixo um exemplo do conteúdo do arquivo de log:

[!NOTE]
Os campos `data` e `hora` são preenchidos automáticamente obtidos do sistema.

```csv
data,hora,nome,destinatario,status
01-01-2025,13:17:19,Fulano de Tal,Sucesso
01-01-2025,13:18:19,Fulano de Tal,<mensagem de erro>

```

#### Executando o código

Para executar o código para disparar os e-mails, execute o programa principal:

```
python3 main.py
```