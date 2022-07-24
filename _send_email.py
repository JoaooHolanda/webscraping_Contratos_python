



def sendEmail():
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    import pandas as pd
    import pandas as pd

    tabela = pd.read_excel('Contratos.xlsx')

  

    #setando o notificar por email o cliente!
    host ='smtp.office365.com' #imutavel, a n ser que haja erro no server
    port = '587' #imutavel, a n ser que haja erro no server
    login = 'joao.holanda@normatel.com.br'
    senha = 'Brasil2021'


    server = smtplib.SMTP(host,port)      


    #setando corpo do email
    corpo = """<html lang="br">
    <head>
        <meta charset="UTF-8">
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <link href="edit.css" rel="stylesheet" type="text/css">
        <title>Document</title>
    </head>
    <body>

        
    <p id="end">Att, João Lucas Silva Holanda</p>

    <p>Sua Automação Sobre os Contratos estão prontas.</p>
    <p>Ao utilizar, Por gentileza, retirar da pasta! </p>
    <p>Se acontecer algum problema, Contacte-me</p>
    <p></p>
    <p></p>
    <p>João Lucas Holanda</p>
    <p>Estagiario Do DP</p>

    </body>"""
  #configurações próprias do server (hotmail)
    server.ehlo()
    server.starttls()
    server.login(login,senha)



    email_msg = MIMEMultipart()

    
    email_msg['From'] = login
    email_msg['To'] = tabela.loc[0,'Email']
    email_msg['Subject'] = 'Robô Programado'
    email_msg.attach(MIMEText(corpo,'html'))

    #enviando

    server.sendmail(email_msg["From"],email_msg["To"],email_msg.as_string())
    
    server.quit()


