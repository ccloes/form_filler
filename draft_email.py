import csv
import os
import sys
import win32com.client

# Get the path to the CSV file and the email body file from command line arguments
csv_file = sys.argv[1]

# Read the email addresses, subjects, body, and attachment names from the CSV file
to_list = []
subject_list = []
attachment_list = []
with open(csv_file, 'r') as file:
    reader = csv.DictReader(file)
    for row in reader:
        to_list.append(row['EMAIL'])
        subject_list.append(row['SUBJECT'])
        attachment_list.append(row['NAME'])
    body = row['BODY']

# Create the email messages and set the recipients, subjects, body, and attachments
outlook = win32com.client.Dispatch("Outlook.Application")
for to, subject, attachment in zip(to_list, subject_list, attachment_list):
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.Subject = subject
    mail.Body = body
    if attachment:
        # Get the full path of the attachment file
        attachment_path = os.path.abspath(attachment)
        # Attach the file to the email
        mail.Attachments.Add(attachment_path)
    mail.Save()
    print(f"Email draft saved for {to} with subject {subject} and attachment {attachment}.")

