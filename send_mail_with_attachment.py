# Python code to illustrate Sending mail with attachments
# from your Gmail account

# libraries to be imported
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

print("email sendign started")
begin_time = datetime.datetime.now()

fromaddr = "ds2020.ipec@gmail.com"
password = "Dsipec2020@"
toaddr = ['06akshay2002@gmail.com','2003012005@ipec.org.in','2003012061@ipec.org.in','2003010051@ipec.org.in','akshay2002singh@gmail.com','2003011061@ipec.org.in','2003011054@ipec.org.in']

# for sender in toaddr:
    # instance of MIMEMultipart
msg = MIMEMultipart()

    # storing the senders email address
msg['From'] = fromaddr

    # storing the receivers email address
# msg['To'] = sender

    # storing the subject
msg['Subject'] = "Subject of the Mail"

    # string to store the body of the mail
body = "Body_of_the_mail"

    # attach the body with the msg instance
msg.attach(MIMEText(body, 'plain'))

    # open the file to be sent
filename = "a.txt"
attachment = open("./a.txt", "rb")

    # instance of MIMEBase and named as p
p = MIMEBase('application', 'octet-stream')

    # To change the payload into encoded form
p.set_payload((attachment).read())

    # encode into base64
encoders.encode_base64(p)

p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    # attach the instance 'p' to instance 'msg'
msg.attach(p)

    # creates SMTP session
s = smtplib.SMTP('smtp.gmail.com', 587)

    # start TLS for security
s.starttls()

    # Authentication
s.login(fromaddr, password)

    # Converts the Multipart msg into a string
text = msg.as_string()


for sender in toaddr:
    # sending the mail
    msg['To'] = sender
    s.sendmail(fromaddr, sender, text)

    # terminating the session
s.quit()


print(datetime.datetime.now() - begin_time)