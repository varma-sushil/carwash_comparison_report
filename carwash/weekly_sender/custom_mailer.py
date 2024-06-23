import smtplib
from email.message import EmailMessage
from email.utils import formataddr
import os
from weekly_sender import get_week_dates_for_storage,create_storage_directory

# Function to collect all .xlsx files from a directory
def get_excel_files(directory):
    return [os.path.join(directory, file) for file in os.listdir(directory) if file.endswith('.xlsx') or file.endswith(".csv")]

# Function to send email
def send_email(subject, body, to_email, from_email, from_name, smtp_server, smtp_port, 
               smtp_user, smtp_password, attachments,cc_emails=None):
    # Create the email message
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = formataddr((from_name, from_email))
    msg['To'] = to_email
    if cc_emails:
        msg['Cc'] = ', '.join(cc_emails)  # Add CC recipients
    msg.set_content(body)

    # Attach the files
    for file_path in attachments:
        with open(file_path, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(file_path)
            msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

    # Connect to the SMTP server and send the email
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()  # Secure the connection
        server.login(smtp_user, smtp_password)
        server.send_message(msg)
        print(f'Email sent to {to_email}')

# # Configuration
# subject = 'Weekly reports'
# body = 'This is the body of the email.'
# to_email = 'vijaykumarmantheena@gmail.com'
# from_email = 'vijaykumarmanthena@reluconsultancy.in'
# from_name = 'Vijay'
# smtp_server = 'smtp-mail.outlook.com'
# smtp_port = 587
# smtp_user = 'vijaykumarmanthena@reluconsultancy.in'
# smtp_password = '4LdfR8qB062DCxt3'

# path = get_week_dates_for_storage()
# storage_path = create_storage_directory(path)

# # Directory containing Excel files
# directory_path = storage_path
# attachments = get_excel_files(directory_path)

# # Send the email
# send_email(subject, body, to_email, from_email, from_name, smtp_server, smtp_port, smtp_user, smtp_password, attachments,cc_emails=NotImplemented)
