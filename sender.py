import email, smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


smtp_server = 'smtp.gmail.com'
port = 465

sender = 'jan.luczewski@gmail.com'
password = input('podaj haslo: ')

receiver = 'artur.sieminski@crist.com.pl'
message = MIMEMultipart('alternative')
message['Subject'] = 'Odbiory'
message['From'] = sender
message['To'] = receiver
text = """\

pozdrowionka
"""

html = """\
<html>
    <body>
        <h1>Oto papiery </h1>
        <h3>załącznik podsyłam w formacie zip</h3>
        <h4>zapisz załącznik</h4>
        <h4>najedź myszą na plik w zapisanym folderze, użyj 7-zip aby rozpakować</h4>
        <h2>Pozdrawiam ,wracam w piątek</h2>     
    </body>
</html>
"""
part1 = MIMEText(text, 'plain')
part2 = MIMEText(html, 'html')


filename = 'odbiory_27.xlsx.zip'


with open(filename, 'rb') as attachment:
    part_a = MIMEBase('application' ,'octet-stream')
    part_a.set_payload(attachment.read())

encoders.encode_base64(part_a)


part_a.add_header(
    'Content_Disposition',
    f'attachment; filename= {filename}',
)


message.attach(part1)
message.attach(part2)
message.attach(part_a)



context = ssl.create_default_context()

with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
    server.login(sender, password)
    server.sendmail(sender, receiver, message.as_string())


