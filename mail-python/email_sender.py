import smtplib
from email.message import EmailMessage
from string import Template
from pathlib import Path

# We create a template object
html = Template(Path('index.html').read_text())

# We create a mail
email = EmailMessage()
email['from_addr'] = 'manu_from_example@gmail.com' 
email['To'] = 'manu_to_example@gmail.com' 
email['subject'] = 'Sending mail from Python' # asunto
email.set_content(html.substitute({'name':'TinTin'}), 'html') # dentro del contenido de correo, usamos el objeto Template

# conectando a mi cuenta de gmail
with smtplib.SMTP_SSL(host='smtp.gmail.com', port=465) as smtp:  
    # empezando protocolo de saludo
    smtp.ehlo()
    smtp.starttls

    # haciendo login
    smtp.login('manu_from_example@gmail.com', 'dqbivurhxcbp')

    # enviando el objeto correo
    smtp.send_message(email)

    # comprobación de que el mensaje se ha enviado
    print('all good')