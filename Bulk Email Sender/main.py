from __future__ import print_function

import base64
import os.path
from email.mime.text import MIMEText

from googleapiclient import errors
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# --------------------------Variable----------------------------------------------
ID = '19F31OODaONIL959CbfYJ-pHKULUu07kqfoNQapDFIpQ'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/gmail.send', 'https://www.googleapis.com/auth/gmail.labels']
SENDER_EMAIL = '21f1006421@student.onlinedegree.iitm.ac.in'
SUBJECT = 'Test'
MESSAGE = '<br>Hey<br>' \
          '<b></i>hii</i></b>'


# Creating Message
def create_message(sender, to, subject, message_text):
    # Args:
    #   sender: Email address of the sender.
    #   to: Email address of the receiver.
    #   subject: The subject of the email message.
    #   message_text: The text of the email message.
    #
    # Returns:
    #   An object containing a base64url encoded email object.

    mails = ','.join(to)
    message = MIMEText(message_text, 'html')
    message['to'] = mails
    message['from'] = sender
    message['subject'] = subject
    return {'raw': base64.urlsafe_b64encode(message.as_string().encode()).decode()}


# Sending Email
def send_message(service, user_id, message):
    """Send an email message.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    message: Message to be sent.

  Returns:
    Sent Message.
  """
    try:
        message = (service.users().messages().send(userId=user_id, body=message)
                   .execute())
        print('Message Id: %s' % message['id'])
        return message
    except errors.HttpError as error:
        print('An error occurred: %s' % error)


creds = None
# The file token.json stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=8080, prompt="consent", authorization_prompt_message='')
        # Save the credentials for the next run
    with open('token.json', 'w') as token:
        token.write(creds.to_json())

service = build('sheets', 'v4', credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()

# Setting range to get email and status------------------Variable-----------------
range_names = [
    'Sheet2!C2:C', 'Sheet2!E2:E'
]

# Getting information from spreadsheet
result = sheet.values().batchGet(
    spreadsheetId=ID, ranges=range_names, majorDimension='COLUMNS').execute()

EMAIL = []

# Extracting Emails and also checking and updating if email is sent to ser or not
for index in range(len(result['valueRanges'][0]['values'][0])):

    num = len(result['valueRanges'][0]['values'][0])

    # Filling Non updated status with empty field
    if num != len(result['valueRanges'][1]['values'][0]):
        result['valueRanges'][1]['values'][0].append('' * (num - len(result['valueRanges'][1]['values'][0])))

    status = result['valueRanges'][1]['values'][0][index]

    # data to be filled in cell
    data = [['Yes']]

    # Checking if email is sent or not
    if status == 'No' or status == '':
        # Getting cell which need to be updated--------------------Variable-------------
        range_val = f'Sheet2!E{2 + index}'
        EMAIL.append(result['valueRanges'][0]['values'][0][index])
        request = service.spreadsheets().values().update(spreadsheetId=ID, range=range_val,
                                                         valueInputOption='USER_ENTERED',
                                                         body={'values': data}).execute()

service = build('gmail', 'v1', credentials=creds)

if len(EMAIL) != 0:
    # Call the Gmail API
    results = service.users().labels().list(userId='me').execute()

    message_val = create_message(SENDER_EMAIL, EMAIL, SUBJECT, MESSAGE)
    result = send_message(service, SENDER_EMAIL, message_val)

else:
    print('No new emails')

