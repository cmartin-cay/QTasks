import win32com.client
from datetime import datetime, timedelta
from pathlib import Path
import dateutil.parser

path = Path("C:\\Users\\Gubbo\\Desktop")

my_dict = {"colin.martKin@gmail.com": path}

outlook = win32com.client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")

inbox = mapi.GetDefaultFolder(6).Folders.Item("Python")
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)

def get_month_from_email(email):
        email_date = email.CreationTime
        email_date_python = dateutil.parser.parse(str(email_date))
        return email_date_python.strftime("%m %B"), email_date_python.strftime("%d")

today = datetime.now() - timedelta(days=5)
today_messages = messages.Restrict(f"[CreationTime] >= '{today.strftime('%m/%d/%Y %H:%M %p')}'")

def save_message(message):
    # Look up the path associated with message sender
    path = my_dict[message.To]
    month, day = get_month_from_email(message)
    # Create the path object for the mont and date and create the folder
    save_location = (path / month / day)
    save_location.mkdir(parents=True, exist_ok=True)
    # Save the message and attachment
    message.SaveAs((save_location / f"{message.Subject}.msg"))
    for attachment in message.Attachments:
        attachment.SaveAsFile((save_location / attachment.FileName))

for message in today_messages:
    if message.To in my_dict:
        save_message(message)




# for message in messages:
#     path = my_dict[message.To]
#     month, day = get_month_from_email(message)
#     save_location = (path / month / day)
#     save_location.mkdir(parents=True, exist_ok=True)
#     message.SaveAs((save_location / "file.msg"))
#     for attachment in message.Attachments:
#         attachment.SaveAsFile((save_location / attachment.FileName))
