from datetime import datetime
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def gemail(name, email):
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    sender_email = "example@gmail.com"
    receiver_email = email
    password = "google passkey f2f"

    subject = "BIRTHDAY WISS"
    body = f"Dear {name},\n\nHAPPY BIRTHDAY!"

    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message.attach(MIMEText(body, "plain"))

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, message.as_string())
            print(f"Email sent successfully to {name} at {email}!")
    except Exception as e:
        print(f"Error sending email to {name}: {e}")

df = pd.read_excel(r"C:\Users\DEvIL'S Kitchen\dob\dob.xlsx")

today = datetime.now()


for index, row in df.iterrows():
    try:
        date = pd.to_datetime(row["date"], dayfirst=True, errors="coerce")

        if date.month == today.month and date.day == today.day:
            name = row["Name"]
            email = row["Email"]
            gemail(name, email)
        else:
            print(f"No match for {name} on today's date.")
    except KeyError as e:
        print(f"Column not found: {e}")
    except Exception as e:
        print(f"Error processing row {index}: {e}")
