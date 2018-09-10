import win32com.client as win32
import csv
from datetime import date


def send_mail(to: str, subject: str, body: str, cc: str = ""):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.CC = cc
    mail.Subject = subject
    mail.Body = body
    mail.Send()

today = date(2018, 9, 30)
subject = "Email Subject"
body = f"""Todays date {today: %d %B %Y}
Second Line
"""

with open('addresses.csv', newline="") as csvfile:
    row_number = 0
    addresses = csv.reader(csvfile)
    for row in addresses:
        if row_number == 0:
            row_number += 1
            continue
        to, cc = row
        cc = cc.replace(',', ';')
        send_mail(to, subject, body, cc)
