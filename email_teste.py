import json
import requests
import base64

# Configurações de autenticação
client_id = 'ae05af2b-1606-4630-9df6-0d1869ce3304'
client_secret = 'e9fb09b1-8343-4413-b88b-47a3b517a0d2'
tenant_id = 'e2d62140-be09-4eb0-98a4-b31d13d73626'
scopes = ['https://graph.microsoft.com/.default']
username = 'conta@billapp.com.br'
password = 'Ram12717'

# URL para obter o access token
token_url = f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token'

# Credenciais para autenticação OAuth2
data = {
    'client_id': client_id,
    'scope': ' '.join(scopes),
    'client_secret': client_secret,
    'grant_type': 'password',
    'username': username,
    'password': password
}

# Codifica as credenciais como base64
credentials = f"{client_id}:{client_secret}"
credentials_bytes = credentials.encode('ascii')
credentials_b64 = base64.b64encode(credentials_bytes).decode('ascii')

# Define o cabeçalho HTTP com as credenciais
headers = {
    'Authorization': f'Basic {credentials_b64}',
    'Content-Type': 'application/x-www-form-urlencoded'
}

# Faz a requisição HTTP para obter o access token
response = requests.post(token_url, data=data, headers=headers)
response.raise_for_status()

# Extrai o access token da resposta
access_token = response.json()['access_token']

# Define o cabeçalho HTTP com o token de acesso
headers = {'Authorization': f'Bearer {access_token}'}

# URL da API para listar as mensagens
api_url = 'https://graph.microsoft.com/v1.0/me/messages'

# Faz a requisição HTTP para listar as mensagens
response = requests.get(api_url, headers=headers)
response.raise_for_status()

# Extrai as mensagens da resposta
messages = response.json()['value']

# Exibe as informações das mensagens
for message in messages:
    print(f"De: {message['from']['emailAddress']['name']}")
    print(f"Assunto: {message['subject']}")
    print(f"Corpo: {message['body']['content']}")
    print("-" * 40)


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
