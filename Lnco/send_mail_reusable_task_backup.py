#!/usr/bin/python

import smtplib
import email.utils
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

SENDER = 'kalyan.gundu@bradsol.com'
SENDERNAME = 'Kalyan'

RECIPIENT = 'kalyan.gundu@bradsol.com'

USERNAME_SMTP = 'AKIATOQHLB3SCULA2HMZ'

PASSWORD_SMTP = 'BKliruk17+I3hqi3g/GYJutwKZYCD4gbKmuP4DGVKCO7'

# CONFIGURATION_SET = ''

HOST = "email-smtp.ap-south-1.amazonaws.com"

PORT = 587

SUBJECT = "test subject"

body_text = "This is a sample body"

msg = MIMEMultipart('alternative')
msg['Subject'] = SUBJECT
msg['From'] = email.utils.formataddr((SENDERNAME, SENDER))
msg['To'] = RECIPIENT

part1 = MIMEText(body_text, 'plain')
# part2 = MIMEText(body_htm, 'html')

msg.attach(part1)
# msg.attach(part2)

try:
    smtp = smtplib.SMTP(HOST, PORT)
    print("Connected to SMTP port")
    smtp.ehlo()
    print("Executed ehlo function")
    smtp.starttls()
    print("Started TLS")
    smtp.login(USERNAME_SMTP, PASSWORD_SMTP)
    print("Login is completed with SMTP credentials")
    smtp.ehlo()
    smtp.sendmail(SENDER, RECIPIENT, str(msg))
    print("sent mail")
except smtplib.SMTPException as e:
    print("Error: unable to send email")
    print(str(e))
