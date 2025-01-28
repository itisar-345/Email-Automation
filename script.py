import pandas as pd 
import smtplib 
from email.mime.text import MIMEText 
from email.mime.multipart import MIMEMultipart 
from email.mime.base import MIMEBase 
from email import encoders 

# Load Excel file containing recipient details
data = pd.read_excel('recipients.xlsx') 

# The Excel file must have columns: 'Name', 'Email', and 'Document'

# Email credentials
sender_email = "your_email" 
sender_password = "app_pass" 
smtp_server = "smtp.gmail.com" 
smtp_port = 587 

# Create an SMTP session
server = smtplib.SMTP(smtp_server, smtp_port) 
server.starttls() 
server.login(sender_email, sender_password) 

# Extract recipient details from each row in the Excel file 
for index, row in data.iterrows():
    name = row['Name'] 
    recipient_email = row['Email'] 
    document_path = row['Document Path'] 

    # Prepare the email
    msg = MIMEMultipart() 
    msg['From'] = sender_email 
    msg['To'] = recipient_email 
    msg['Subject'] = "Your Document" 

    # Email body
    body = f"""\
Dear {name},

Congratulations for attending our Event! 
Please find your document attached.

Best regards,
Your Name
"""
    msg.attach(MIMEText(body, 'plain'))  # Attach the plain text email body to the email

    # Attach the certificate file
    try:
        with open(document_path, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')  
            part.set_payload(attachment.read()) 
            encoders.encode_base64(part) 
            part.add_header(
                'Content-Disposition',
                f"attachment; filename={document_path.split('/')[-1]}" 
            )
            msg.attach(part) 

        # Send the email
        server.send_message(msg) 
        print(f"Email sent to {recipient_email}") 

    except FileNotFoundError:
        print(f"Document not found for {name} at {document_path}")

# Close the SMTP session
server.quit() 
print("All emails have been sent.") 
