# -*- coding:utf-8 -*-
import smtplib, ssl
import pandas as pd
from getpass import getpass

mails = pd.read_excel("mails.xlsx", engine='openpyxl')
mails = mails[mails.iloc[:,1].notna()]
context = ssl.create_default_context()

def senderMail():
    sender_email = input("From which e-mail address will it be sent?\n")
    password = getpass("What is the password of the e-mail? ")

    servers = sender_email.split("@")
    servers2 = servers[1].split(".")

    if servers2[0] == 'outlook' or servers2[0] == 'hotmail':
        smtp_server = "SMTP.office365.com"
    elif servers2[0] == 'gmail':
        smtp_server = "smtp.gmail.com"
    return sender_email, password,smtp_server

def getMessage():
    subject = input(u"What is the subject of the e-mail?\n")
    body = input(u"What is the content of the e-mail?\n")
    message = f"""\
Subject: {subject}

{body}"""
    return message

def whichColumn():
    column = int(input("What column are the email addresses in? Please enter as a number. "))
    column -=1
    return column

try:
    sender_email, password,smtp_server = senderMail()
    message = getMessage()
    column = whichColumn()
    server = smtplib.SMTP(smtp_server,587)
    server.command_encoding = 'utf-8'
    server.ehlo()
    server.starttls(context=context)
    server.ehlo()
    server.login(sender_email, password)
    for receiver_email in mails.iloc[:,column]:
        if '@' in receiver_email:
            server.sendmail(sender_email, receiver_email, message)
        else:
            pass
except Exception as e:
    print(e)
finally:
    print("E-mails have been sent successfully.")
    server.quit()