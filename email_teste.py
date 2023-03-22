import imaplib
import email
import os
import base64
import requests
import json

# Credenciais OAuth2
client_id = 'ae05af2b-1606-4630-9df6-0d1869ce3304'
client_secret = 'e9fb09b1-8343-4413-b88b-47a3b517a0d2'
username = 'conta@billapp.com.br'
password = 'Ram12717'
scopes = ['https://outlook.office.com/mail.read']
tenant_id = 'e2d62140-be09-4eb0-98a4-b31d13d73626'

# Obter token de acesso
api_url = 'https://outlook.office.com/api/v2.0/me/messages'
oauth_url = 'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token'

data = {
    'grant_type': 'password',
    'client_id': client_id,
    'client_secret': client_secret,
    'username': username,
    'password': password,
    'scope': scopes
}
response = requests.post(oauth_url, data=data)

# verificar a resposta da solicitação POST
if response.status_code != 200:
    print('Erro:', response.status_code, response.text)
    exit()

# verificar o JSON retornado
json_response = response.json()
if 'access_token' not in json_response:
    print('Erro: a chave "access_token" não foi encontrada no JSON:', json_response)
    exit()

access_token = response.json()['access_token']

headers = {
    'Authorization': 'Bearer ' + access_token,
    'Accept': 'application/json',
    'Content-Type': 'application/json'
}
response = requests.get(api_url, headers=headers)
messages = response.json()['value']

# exibir as informações da mensagem
for message in messages:
    print('From:', message['From']['EmailAddress']['Name'])
    print('Subject:', message['Subject'])
    print('Body:', message['Body']['Content'])
    print()


#Conectando ao servidor do outlook com IMAP
#objCon = imaplib.IMAP4_SSL("outlook.office365.com")

#Credenciais
#login = "conta@billapp.com.br"
#senha = "Ram12717"

#objCon.login(login, senha)

#Loopar a caixa de entrada
#objCon.select(mailbox='inbox', readonly=True)
#respostas, idDosEmails = objCon.search(None, 'All')

#for num in idDosEmails[0].split():
    #decodificando o email e jogando em uma variavel as partes
    #resultado, dados = objCon.fetch(num, '(RFC822)')
    #texto_do_email = dados[0][1]
   #texto_do_email = texto_do_email.decode('utf-8')
    #texto_do_email = email.message_from_string(texto_do_email)

    #print(texto_do_email)
    #for part in texto_do_email.walk():
     #   if part.get_content_maintype() == 'multipart':
      #      continue
       # if part.get('Content-Disposition') is None:
        #    continue
        #Pegando o nome do arquivo em anexo
        #fileName = part.get_filename()

        #Criamos um arquivo com o mesmo nome na pasta local
        #arquivo = open(fileName, 'wb')

        #Escrevendo o binário do anexo no arquivo
        #arquivo.write(part.get_payload(decode=True))
        #arquivo.close()
