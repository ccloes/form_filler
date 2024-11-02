import csv
import os
import sys
import win32com.client

# Get the path to the CSV file and the email body file from command line arguments
csv_file = sys.argv[1]

# Read the email addresses, subjects, and body from the CSV file
to_list = []
subject_list = []
with open(csv_file, 'r') as file:
    reader = csv.DictReader(file)
    for row in reader:
        to_list.append(row['EMAIL'])
        subject_list.append(row['SUBJECT'])
body = row['BODY']

# Create the email messages and set the recipients, subjects, and body
outlook = win32com.client.Dispatch("Outlook.Application")
for to, subject in zip(to_list, subject_list):
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.Subject = subject
    mail.Body = body
    mail.Save()
    print(f"Email draft saved for {to} with subject {subject}.")

