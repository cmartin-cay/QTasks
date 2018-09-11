import win32com.client as win32
import csv
from datetime import date


def send_email(to: str, subject: str, body: str, cc: str = ""):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.CC = cc
    mail.Subject = subject
    mail.Body = body
    mail.Send()

today = date(2018, 9, 30)
subject = "Email Subject"
def prepare_body(name, date):
    body = f"""To {name}
Todays date {date: %d %B %Y}
Second Line
"""
    return body

def read_file(location):
    with open(location, newline="") as csvfile:
        recipient_details = []
        row_number = 0
        addresses = csv.reader(csvfile)
        for row in addresses:
            if row_number == 0:
                row_number += 1
                continue
            recipient_details.append(row)
    return recipient_details

def prepare_email(recipient_details):
    for recipient in recipient_details:
        to, cc, name = recipient
        cc = cc.replace(',', ';')
        body = prepare_body(name, today)
        send_email(to, subject, body, cc)

recipients = read_file('addresses.csv')
prepare_email(recipients)