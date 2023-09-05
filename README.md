# Email Automation Script
This Python script automates the process of sending personalized emails with attachments to a list of recipients from an Excel file. It leverages Gmail's SMTP server for email delivery and uses the open-source openpyxl library for managing recipient data. The script is designed to make email campaigns more efficient and trackable.


This Python script automates the process of sending emails with attachments to a list of recipients from an Excel file. It uses the smtplib library to send emails via Gmail's SMTP server and the openpyxl library to manage the Excel data.

Prerequisites
Before running the script, make sure you have the following set up:

Gmail account credentials:

sender_email: Your Gmail email address.
sender_password: App password for Gmail.
Excel file:

excel_file_path: Path to the Excel file containing recipient data.
The Excel file should have a sheet named 'Sheet1' with columns 'Email', 'Name', and 'Status'. 'Status' is used to mark emails as 'completed' to avoid resending.
Email content:

Customize the message_template variable with your desired email content in HTML format. If you prefer plain text, modify the code to use 'plain' instead of 'html'.
Attachment:

Specify the constant attachment file using filename.

Usage
Clone this repository to your local machine.

Install the required libraries if not already installed:

# pip install openpyxl


Certainly, here's a short documentation for the provided code that you can add to your README file on GitHub:

Email Automation Script
This Python script automates the process of sending emails with attachments to a list of recipients from an Excel file. It uses the smtplib library to send emails via Gmail's SMTP server and the openpyxl library to manage the Excel data.

Prerequisites
Before running the script, make sure you have the following set up:

Gmail account credentials:

sender_email: Your Gmail email address.
sender_password: App password for Gmail.
Excel file:

excel_file_path: Path to the Excel file containing recipient data.
The Excel file should have a sheet named 'Sheet1' with columns 'Email', 'Name', and 'Status'. 'Status' is used to mark emails as 'completed' to avoid resending.
Email content:

Customize the message_template variable with your desired email content in HTML format. If you prefer plain text, modify the code to use 'plain' instead of 'html'.
Attachment:

Specify the constant attachment file using filename.
Usage
Clone this repository to your local machine.

Install the required libraries if not already installed:

# pip install openpyxl
Customize the script:

Replace sender_email, sender_password, excel_file_path, message_template, and filename with your own values.
Modify the email content within message_template as needed.
Run the script:
# python send_emails.py

The script will send emails with attachments to the recipients listed in the Excel file. Recipients marked as 'completed' in the 'Status' column will be skipped.

After sending, the 'Status' column in the Excel file will be updated to 'completed' for each sent email to prevent resending.

Notes
Ensure your Gmail account has "Less secure apps" enabled or use an App password for authentication.
Customize the email content and attachment file as per your requirements.
Use the 'Status' column in the Excel file to track the status of sent emails.
