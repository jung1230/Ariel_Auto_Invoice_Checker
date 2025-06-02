# Daily SQL Report Automation
This script fetches data from a SQL Server view and emails it as an Excel report daily — without needing Excel installed.

---

##Configuration: secrets.json
Before running the script, create a file named secrets.json in the same directory with the following structure:

```
{
  "server": "your_sql_server:port",
  "database": "your_database_name",
  "username": "your_db_username",
  "password": "your_db_password",
  "email_sender": "your_email@example.com",
  "email_sender_password": "your_email_password_or_app_password"
}
```

### Important Note
If your email account uses two-factor authentication (2FA), you must use an application-specific password instead of your regular login password.
https://support.google.com/accounts/answer/185833?hl=en

---

## Setup & Packaging (Optional but Recommended)
To convert the script into a standalone .exe:

1. Open your terminal or command prompt:

```
pip install pyinstaller
```

2. Build the executable:

```
pyinstaller --onefile main.py
```

3. After the build completes, your .exe file will be located in the dist/ directory:

```
dist/
└── main.exe
```
4. Move secrets.json into the same folder as main.exe.

---

## Automate with Windows Task Scheduler
To run your .exe file automatically every day:

1. Press Win + R, type taskschd.msc, and press Enter.

2. Click Create Basic Task (right pane).

3. Name it: Send Daily SQL Report.

4. Set the trigger to run Daily at your preferred time.

5. Action: Start a program.

6. Browse to select the compiled main.exe file.

7. Finish and test the task.

Your script will now run automatically every day and send the report via email.