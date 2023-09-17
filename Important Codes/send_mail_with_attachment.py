# Python code to illustrate Sending mail with attachments
# from your Gmail account

# libraries to be imported
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

print("email sending started")
begin_time = datetime.datetime.now()

# These id and password are dummy use your own 
fromaddr = "ds2020.ipec@gmail.com"
# This password is not your gmail password, this is app password that is generated via google
password = "Dsipec2020@"
toaddr = ['06akshay2002@gmail.com','2003012005@ipec.org.in','akshay2002singh@gmail.com']


# instance of MIMEMultipart
msg = MIMEMultipart()

    # storing the senders email address
msg['From'] = fromaddr

# storing the subject
msg['Subject'] = "Subject of the Mail"

# string to store the body of the mail
body = "Body_of_the_mail"

# attach the body with the msg instance
msg.attach(MIMEText(body, 'plain'))

# open the file to be sent
filename = "a.txt"
attachment = open("testing_attachment.txt", "rb")

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


print(f"Time Taked = {datetime.datetime.now() - begin_time}")