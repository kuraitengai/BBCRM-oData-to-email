# -*- coding: utf-8 -*-
"""
Created on Wed Sep  2 11:03:55 2020

@author: test
"""

from tabulate import tabulate
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import odata_queries as o
import utilities as u

recipient = 'email1@email.com'
recipients = ['email1@email.com', 'email2@email.com', 'email3@email.com']
subject = "Assignment Error Check"

message = MIMEMultipart('alternative')

def emailformat(df):
    #EMAIL FORMAT
    text = """
    Here are the Assignments where both members of the household are assigned to a Prospect Manager:
    
    {table}
    
    
    In the case of both household members assigned to the SAME Prospect Manager, determine which should remain active per the original assignment request and remove the other member using an end date that is identical to the start date.
    In the case where the household members are assigned to DIFFERENT Prospect Managers, determine which should remain assigned, usually based on the respective Start Date, and remove the other member using an end date identical to the start date.
    """
    
    html = """
    <table>
    <tbody>
    <p>Here are the Assignments where both members of the household are assigned to a Prospect Manager:</p>
    {table}
    
    <p>In the case of both household members assigned to the SAME Prospect Manager, determine which should remain active per the original assignment request and remove the other member using an end date that is identical to the start date.</p>
    <p>In the case where the household members are assigned to DIFFERENT Prospect Managers, determine which should remain assigned, usually based on the respective Start Date, and remove the other member using an end date identical to the start date.</p>
    </tbody></table>
    """
    
    text = text.format(table=tabulate(df, headers='keys', tablefmt='rst'))
    html = html.format(table=tabulate(df, headers='keys', tablefmt='html'))
    
    part1 = MIMEText(text, 'plain')
    part2 = MIMEText(html, 'html')
    
    message.attach(part1)
    message.attach(part2)
    
    return html

def emailerror():
    text = """
    There was an issue with the Assignment Error Check query. The query location or security could have been changed or the oData password has expired.
    
    """
    
    html = """
    <p>There was an issue with the Assignment Error Check query. The query location or security could have been changed or the oData password has expired.</p>
    """
    
    part1 = MIMEText(text, 'plain')
    part2 = MIMEText(html, 'html')
    
    message.attach(part1)
    message.attach(part2)
    
    return html

#do NOT use all three try/except options. Only use the ONE for the service you are actually using.

#use for Gmail sending
try:
    df = o.assignmenterrors()
    html = emailformat(df)
    u.gmail(recipients, subject, html)
except:
    html = emailerror()
    u.gmail(recipients, subject, html)


#use for OutlookOWA sending
try:
    df = o.assignmenterrors()
    html = emailformat(df)
    u.outlookowa(recipients, subject, html)
except:
    html = emailerror()
    u.outlookowa(recipients, subject, html)

#use for local Outlook sending when Outlook is open on the computer
try:
    df = o.assignmenterrors()
    html = emailformat(df)
    u.send_notification(recipients, subject, html)
except:
    html = emailerror()
    u.send_notification(recipients, subject, html)
