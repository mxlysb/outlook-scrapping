import imaplib
import email
import PyPDF4
import os
import requests
from io import BytesIO
from urllib.parse import urlparse
from bs4 import BeautifulSoup
from requests_oauthlib import OAuth2Session
from oauthlib.oauth2 import BackendApplicationClient
import base64
import json
from msal import PublicClientApplication
from exchangelib import Credentials, Account

# Define the app credentials
client_id = '44396eda-0f9d-456d-9ed0-a7244cceac13'
client_secret = 'y258Q~6uDoocWqMmsDj7PM4_4v5e~ZMXWPEklbg6'

email_login = 'administrador@billapp.com.br'

# Define the necessary permissions
scope = ['https://graph.microsoft.com/.default']

# Define o URL de autorização e o URL de token
AUTHORITY_URL = f"https://login.microsoftonline.com/e2d62140-be09-4eb0-98a4-b31d13d73626"
TOKEN_URL = f"https://login.microsoftonline.com/e2d62140-be09-4eb0-98a4-b31d13d73626/oauth2/v2.0/token"

# Cria uma instância do cliente MSAL
app = PublicClientApplication(client_id, authority=AUTHORITY_URL)

# Faz a autenticação do usuário
result = None
accounts = app.get_accounts(username=email_login)
if accounts:
    # Obtém um token para o usuário autenticado
    result = app.acquire_token_silent(scope, account=accounts[0])

if not result:
    # Faz a autenticação interativa do usuário
    result = app.acquire_token_interactive(scope)

print(result)

# Obtém o token de acesso
access_token = result["access_token"]

# Usa o token de acesso para fazer uma requisição à API do Outlook
response = requests.get("https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages?$top=10", headers={"Authorization": f"Bearer {access_token}"})

# Imprime a resposta
print(response.json())



def download_pdf(url):
    """
    Faz o download do arquivo PDF a partir de uma URL.
    :param url: a URL do arquivo PDF
    """
    filename = url.split("/")[-1]

    # Obtém o conteúdo do arquivo PDF
    response = requests.get(url)
    pdf_content = response.content

    # Verifica se o conteúdo é um PDF válido
    try:
        pdf_reader = PyPDF4.PdfFileReader(BytesIO(pdf_content))
        pdf_reader.getNumPages()
    except PyPDF4.utils.PdfReadError:
        print(f'O arquivo {filename} não é um PDF válido')
        return

    # Escreve o conteúdo do PDF em um arquivo local
    with open(filename, "wb") as f:
        f.write(pdf_content)
        print(f"O arquivo {filename} foi baixado com sucesso!")


def process_email_part(part):
    """
    Processa uma parte do email (pode ser uma mensagem, anexo ou corpo HTML)
    :param part: a parte do email a ser processada
    """
    content_type = part.get_content_type()
    content_disposition = str(part.get("Content-Disposition"))

    if content_type == "text/html" and "attachment" not in content_disposition:
        # Verifica se é um email HTML com o corpo dentro de uma tag <div>
        soup = BeautifulSoup(part.get_payload(decode=True), "html.parser")
        body = soup.find("div")

        # Encontra todos os links dentro do corpo do e-mail
        for link in body.find_all("a"):
            href = link.get("href")
            parsed_href = urlparse(href)

            if parsed_href.scheme and parsed_href.netloc and parsed_href.path.endswith('.pdf'):
                download_pdf(href)

    elif part.get_content_maintype() == 'multipart':
        pass
    elif part.get('Content-Disposition') is None:
        pass
    elif part.get_content_type() == 'application/pdf':
        # Pegando o nome do arquivo em anexo
        fileName = part.get_filename()

        # Criamos um arquivo com o mesmo nome na pasta local
        with open(fileName, 'wb') as file:
            # Escrevendo o binário do anexo no arquivo
            file.write(part.get_payload(decode=True))

        # Verificando se o arquivo é realmente um PDF
        with open(fileName, 'rb') as pdf_file:
            try:
                pdf_reader = PyPDF4.PdfFileReader(pdf_file)
                pdf_reader.getNumPages()
            except PyPDF4.utils.PdfReadError:
                os.remove(fileName)  # remover o arquivo se não for um PDF válido
                print(f'O arquivo {fileName} não é um PDF válido')

        file.close()

def mark_email_as_read(access_token, email_id):
    """
    Marca um e-mail como lido.
    :param access_token: o token de acesso à API do Outlook
    :param email_id: o ID do e-mail a ser marcado como lido
    """
    url = f"https://graph.microsoft.com/v1.0/me/messages/{email_id}"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    data = {
        "isRead": True
    }
    response = requests.patch(url, headers=headers, json=data)
    response.raise_for_status()

def process_email(connectionObject, num):
    """
    Processa um email completo, incluindo todas as partes (mensagens, anexos, corpo HTML)
    :param connectionObject: objeto de conexão com a caixa de entrada do email
    :param num: número do email na caixa de entrada
    """
    result, dados = connectionObject.fetch(num, '(RFC822)')
    text_email = dados[0][1]
    text_email = text_email.decode('utf-8')
    text_email = email.message_from_string(text_email)

    for part in text_email.walk():
        process_email_part(part)
    
    email_id = result.split()[0].decode('utf-8')
    mark_email_as_read(access_token, email_id)

