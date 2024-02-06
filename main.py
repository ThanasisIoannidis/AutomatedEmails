



# Second part


import os
from datetime import date
import smtplib
from email.message import EmailMessage
from email.utils import formataddr
import pandas as pd
from pathlib import Path

from dotenv import load_dotenv  # pip install python-dotenv

PORT = 587
EMAIL_SERVER = "smtp-mail.outlook.com"  # Adjust server address, if you are not using @outlook

# setup variables
sender_email = "A.Ioannidis03@outlook.com"
password_email = "Thanas1s121103!"



def send_email(subject, receiver_email, name):
    # Create the base text message.
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = formataddr(("Our Company.", f"{sender_email}"))
    msg["To"] = receiver_email
    msg["BCC"] = sender_email

    msg.set_content(
        f"""\
        Hi {name},
        Our company would like to wish you happy birthday!! (paradeigma)
        """
    )

    with smtplib.SMTP(EMAIL_SERVER, PORT) as server:
        server.starttls()
        server.login(sender_email, password_email)
        server.sendmail(sender_email, receiver_email, msg.as_string())



SHEET_ID = '1EkC1q3brYFUzASM5tTSRfMxqkDEfyFqyT5N73QBiZV0'
SHEET_NAME = 'Sheet1'
url = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={SHEET_NAME}'
df = pd.read_csv(url)
# print(df.head())

# print(df.dob)








# First part

# Imports for Main
from datetime import date
from datetime import datetime
import pandas as pd

SHEET_ID = '1EkC1q3brYFUzASM5tTSRfMxqkDEfyFqyT5N73QBiZV0'
SHEET_NAME = 'Sheet1'
url = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={SHEET_NAME}'
df = pd.read_csv(url)


present_date = date.today()
present = str(present_date) + " 00:00:00"
present_datetime = datetime.strptime(present, '%Y-%m-%d   %H:%M:%S')


for _, row in df.iterrows():
    dob_date = row["dob"]
    name = row["name"]
    email = row["email"]
    dob_datetime = str(dob_date) + " 00:00:00"
    date_obj = datetime.strptime(dob_datetime, '%Y-%m-%d  %H:%M:%S')

    if date_obj == present_datetime:
        print("I did it!")
        send_email(
            subject="Happy Birthday!",
            name=name,
            receiver_email=email,
            )




