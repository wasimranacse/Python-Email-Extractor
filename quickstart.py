import os
from google.auth.transport.requests import Request
import pandas as pd
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Set up the Gmail API
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
creds = None
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    with open('token.json', 'w') as token:
        token.write(creds.to_json())

# Set up the Gmail service
service = build('gmail', 'v1', credentials=creds)

# Set the label ID for the label you want to extract emails from
label_id = 'INBOX'

# Retrieve the list of messages in the label
try:
    results = service.users().messages().list(userId='me', labelIds=[label_id]).execute()
    emails = results.get('messages', [])
except HttpError as error:
    print(f'An error occurred: {error}')
    emails = []

# Receive unlimited emails from the label
# emails = []
# next_page_token = None
# while True:
#     results = service.users().messages().list(userId='me', labelIds=[label_id], pageToken=next_page_token).execute()
#     emails.extend(results.get('messages', []))
#     next_page_token = results.get('nextPageToken')
#     if not next_page_token:
#         break

# Extract the email data for each message in the list
sender_info = []
for message in emails:
    msg = service.users().messages().get(userId='me', id=message['id']).execute()
    headers = msg['payload']['headers']
    sender = [header['value'] for header in headers if header['name'] == 'From'][0]
    sender_name, sender_email = sender.split(' <')[0], sender.split(' <')[1][:-1] if '<' in sender else (sender, sender)
    sender_info.append({'name': sender_name, 'email': sender_email})

# Convert the list of emails to a pandas DataFrame
df = pd.DataFrame(sender_info)

# Save the DataFrame to an Excel file
df.to_excel('emails.xlsx', index=False)