# -*- coding: utf-8 -*-
"""
Created on Thu Oct 15 12:15:26 2020

@author: test
"""

import smtplib
import win32com.client as win32
import psutil
import os
import subprocess
from O365 import Account
from secret import gusername
from secret import gpassword
from secret import clientid
from secret import secret

def gmail(recipient, subject, text):
    sender = gusername
    password = gpassword
    
    smtp_server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
    smtp_server.login(sender, password)
    message = "Subject: {}\n\n{}".format(subject, text)
    smtp_server.sendmail(sender, recipient, message.as_string())
    smtp_server.close()

def send_notification(recipient, subject, text):
   outlook = win32.Dispatch('Outlook.Application')
   mail = outlook.CreateItem(0)
   #mail.To = recipients
   for recipient in recipients:
       mail.Recipients.Add(recipient)
   mail.Subject = subject
   mail.Body = text
   #mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional
   mail.Send()

def open_outlook():
    try:
        subprocess.call(['C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE'])
        os.system("C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE")
    except:
        print("Outlook didn't open successfully")

def outlookowa(recipient, subject, text):
    credentials = (clientid,secret)
    account = Account(credentials)
    m = account.new_message()
    m.to.add(recipient)
    m.subject = subject
    m.body = text
    m.send()

def outlookowaattach(recipient, subject, text, emailfile):
    credentials = (clientid,secret)
    account = Account(credentials)
    m = account.new_message()
    m.to.add(recipient)
    m.subject = subject
    m.body = text
    m.attachments.add(emailfile)
    m.send()