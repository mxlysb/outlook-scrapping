import imaplib
import email

#Conectando ao servidor do outlook com IMAP
objCon = imaplib.IMAP4_SSL("outlook.office365.com")

#Credenciais
login = "billteste165@outlook.com"
senha = "@teste123"

objCon.login(login, senha)

#Loopar a caixa de entrada
objCon.select(mailbox='inbox', readonly=True)
respostas, idDosEmails = objCon.search(None, 'All')

for num in idDosEmails[0].split():
    #decodificando o email e jogando em uma variavel as partes
    resultado, dados = objCon.fetch(num, '(RFC822)')
    texto_do_email = dados[0][1]
    texto_do_email = texto_do_email.decode('utf-8')
    texto_do_email = email.message_from_string(texto_do_email)

    print(texto_do_email)
    for part in texto_do_email.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        #Pegando o nome do arquivo em anexo
        fileName = part.get_filename()

        #Verificando se é um arquivo PDF
        if fileName and part.get_content_type() == 'application/pdf':
            #Criamos um arquivo com o mesmo nome na pasta local
            with open(fileName, 'wb') as arquivo:
                #Escrevendo o binário do anexo no arquivo
                arquivo.write(part.get_payload(decode=True))
                arquivo.close()

