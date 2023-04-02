import imaplib
import email
import PyPDF4
import msal
import os
import requests
from io import BytesIO
from urllib.parse import urlparse
from bs4 import BeautifulSoup
from requests_oauthlib import OAuth2Session


def authenticate_with_oauth2(client_id, client_secret, redirect_uri, scope):
    """
    Authenticate with OAuth2 and return the access token.

    :param client_id: the client ID of the OAuth2 app
    :param client_secret: the client secret of the OAuth2 app
    :param redirect_uri: the redirect URI of the OAuth2 app
    :param scope: the list of OAuth2 scopes to request
    """
    oauth = OAuth2Session(client_id, redirect_uri=redirect_uri, scope=scope)

    authorization_url, state = oauth.authorization_url('https://login.microsoftonline.com/common/oauth2/v2.0/authorize')

    # Redirect the user to the authorization URL and request the authorization code

    # Get the authorization code from the user

    token = oauth.fetch_token('https://login.microsoftonline.com/common/oauth2/v2.0/token', authorization_response='URL de redirecionamento com o código de autorização', client_secret=client_secret)

    return token['access_token']


# Define the app credentials
client_id = 'ID do aplicativo'
client_secret = 'Segredo do cliente'
redirect_uri = 'URL de redirecionamento'

# Define the necessary permissions
scope = ['https://outlook.office.com/IMAP.AccessAsUser.All']

# Authenticate with OAuth2
access_token = authenticate_with_oauth2(client_id, client_secret, redirect_uri, scope)

# Connect to the Outlook email server with IMAP
email_server = imaplib.IMAP4_SSL("outlook.office365.com")

# Set the credentials
email_login = "administrador@billapp.com.br"
email_password = "3i11@99P"

# Authenticate using the access token
email_server.authenticate('XOAUTH2', lambda x: f'user={email_login}\x01auth=Bearer {access_token}\x01\x01')

# Loop through the inbox
email_server.select(mailbox='inbox', readonly=True)
responses, email_ids = email_server.search(None, 'UNSEEN')

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

