# Automated Document Mailer Project

This project automates the process of sending certificates via email using Python. The script reads user details from an Excel file, attaches documents, and sends personalized emails to the recipients.

---

## Features
- Reads recipient details from an Excel file.
- Sends personalized emails to each recipient.
- Attaches documents to emails.
- Logs success or errors for each email.

---

## Requirements

### Libraries to Install
Make sure the following Python libraries are installed:

```bash
pip install pandas openpyxl
```

- **pandas**: For reading data from the Excel file.
- **openpyxl**: For handling Excel (.xlsx) files.

---

## Setting Up 2FA and App Password
Since the script uses your Gmail account to send emails, you need to enable 2FA and generate an App Password.

### Step 1: Enable 2-Factor Authentication (2FA)
1. Go to [Google Account Security Settings](https://myaccount.google.com/security).
2. Under **"Signing in to Google"**, click **2-Step Verification** and enable it.
3. Follow the on-screen instructions to set up 2FA using your phone or an authenticator app.

### Step 2: Generate an App Password
1. Once 2FA is enabled, go back to **"Signing in to Google"** and select **App Passwords**.
2. Choose the app as **Mail** and the device as **Other** (e.g., Python Script).
3. Google will generate a 16-character password (e.g., `abcd-efgh-ijkl-mnop`).
4. Use this App Password in the script instead of your regular email password.

---

## Preparing the Excel File

The Excel file should have the following columns:

1. **Name**: Full name of the recipient.
2. **Email**: Email address of the recipient (e.g., `user@gmail.com`).
3. **Certificate Path**: The relative or absolute file path to the recipient's certificate.

---

## Organizing Certificate Files

1. Store all certificate files in a directory named `documents` (or a similar structure).
2. Ensure the file names match the paths mentioned in the Excel file.
3. Use relative paths in the Excel file for portability (e.g., `documents/user1.pdf`).
4. Example directory structure:

```
automatedmailer/
├── documents/
│   ├── user1.pdf
│   ├── user2.pdf
│   ├── ...
├── script.py
├── recipients.xlsx
```

---

## Running the Script

1. Place the Excel file in the project directory.
2. Update the `data = pd.read_excel('recipients.xlsx')` line in the script with the actual file name.
3. Run the script:

```bash
python script.py
```

4. The script will send emails to all recipients listed in the Excel file.

---

## Notes
- Ensure the Excel file paths and document paths are accurate.
- If the script fails for any recipient, it will log the error for debugging.
- Double-check that your App Password is kept secure and not shared publicly.

Feel free to customize the script or the email body to suit your needs!

