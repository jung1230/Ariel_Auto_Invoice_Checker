import pyodbc, pandas as pd, os, smtplib
from email.message import EmailMessage
import json
import shutil
import logging
from datetime import datetime

# Setup logging
log_filename = "daily_report_log.txt"
logging.basicConfig(
    filename=log_filename,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logging.info("Script started")

try:

    # Load secrets
    with open('secrets.json') as f:
        secrets = json.load(f)


    conn_str = (
        f"DRIVER={{ODBC Driver 18 for SQL Server}};"
        f"SERVER={secrets['server']},1433;"
        f"DATABASE={secrets['database']};"
        f"UID={secrets['username']};"
        f"PWD={secrets['password']};"
        "Encrypt=no;"
        "TrustServerCertificate=yes;"
    )

    conn = pyodbc.connect(conn_str)

    query1 = "SELECT * FROM dbo.OrderInformationQuery"
    query2 = "SELECT * FROM dbo.ProductInformationQuery"


    conn = pyodbc.connect(conn_str)
    df1 = pd.read_sql(query1, conn)
    df2 = pd.read_sql(query2, conn)
    # print(df1.head())
    # print("----------")
    # print(df2.head())
    os.makedirs("Output_CSV", exist_ok=True)
    df1.to_csv(f"Output_CSV/OrderInformationQuery.csv", index=False)
    df2.to_csv(f"Output_CSV/ProductInformationQuery.csv", index=False)
    logging.info("Data fetched successfully")
    # --- Email setup ---
    email_sender = secrets['email_sender']
    print(f"Email sender: {email_sender}")
    email_sender_password = secrets['email_sender_password']

    msg = EmailMessage()
    msg["From"] = email_sender
    msg["To"] = "AlanC25@ArielPremium.com"
    print(f"Email recipient: {msg['To']}")
    msg["Subject"] = "Daily SQL Report - " + datetime.now().strftime("%Y-%m-%d")
    Today = datetime.now().strftime("%Y-%m-%d")
    msg.set_content(f"Hi,\n\nPlease find attached {Today}'s report.\n\nBest regards.")

    file_paths = ["Output_CSV/OrderInformationQuery.csv","Output_CSV/ProductInformationQuery.csv"]
    # Attach the file
    for file_path in file_paths:
        with open(file_path, "rb") as f:
            data = f.read()
            msg.add_attachment(data,maintype="text",subtype="csv",filename=os.path.basename(file_path))


    # --- Send via SMTP ---
    smtp = smtplib.SMTP("smtp.gmail.com", 587)    # change to your SMTP server(Gmail or Outlook)
    smtp.ehlo()                                         
    smtp.starttls()                                     
    smtp.ehlo()                                         
    smtp.login(email_sender, email_sender_password)               
    smtp.send_message(msg)     
    logging.info("Email sent successfully")
                      
    smtp.quit()                                       
    f.close()
    conn.close()
    shutil.rmtree("Output_CSV")
except Exception as e:
    logging.error(f"An error occurred: {e}")