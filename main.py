#!./bin/python
from __future__ import print_function
import os.path

from typing import List, Any
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.errors import HttpError
import base64
from email.mime.text import MIMEText
import google.auth
from googleapiclient.discovery import build

from dotenv import load_dotenv
from os import getenv as env

load_dotenv()

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly',
          'https://mail.google.com/']

# The ID and range of a sample spreadsheet.

SPREADSHEET_ID = env("SPREADSHEET_ID")
RANGE = env("RANGE")
EMAIL_ADDRESS = env("EMAIL_ADDRESS")
TEXTFILE_PATH = env("TEXTFILE_PATH")
EMAIL_SUBJECT = env("EMAIL_SUBJECT")

def read_message_and_replace_fields(dictionary_fields_to_replace):
    with open(TEXTFILE_PATH) as file:
        text = file.read()
        text_replace_fields = text
    for key in dictionary_fields_to_replace.keys():
        try:
            text_replace_fields = text_replace_fields.replace(key, dictionary_fields_to_replace[key])
        except:
            continue
    return text_replace_fields



class Main:
    def __init__(self):
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
                flow = InstalledAppFlow.from_client_secrets_file(env("PATH_TO_JSON_CREDENTIALS"), SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
        self.creds = creds
        self.values = None

    def build_requests_from_sheet(self) -> List[List[Any]]:
        """
        Get values from a sample spreadsheet.
        Returns values : 2D List of the values in the spreadsheet
        """
        try:
            service = build('sheets', 'v4', credentials=self.creds)

            # Call the Sheets API
            sheet = service.spreadsheets()
            result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                        range=RANGE).execute()
            values = result.get('values', [])
            if not values:
                print('No data found.')
                return
            self.values = values
            return values
        except HttpError as err:
            print(err)

    def send_email_to_values(self):
        column_index = {}
        for i, column_name in enumerate(self.values[0]):
            column_index[column_name] = i
        for row_index in range(1, len(self.values), 1):
            dictionary_fields_to_replace = {}
            for key in column_index.keys():
                dictionary_fields_to_replace[key] = self.values[row_index][column_index[key]]
                print(key, " : ", self.values[row_index][column_index[key]], end=', ')
            print()
            text_email = read_message_and_replace_fields(dictionary_fields_to_replace)
            email_address = dictionary_fields_to_replace["EMAIL"]
            self.gmail_create_draft(email_address, text_email)

    def gmail_create_draft(self, email_address, email_text):
        """Create and insert a draft email.
           Print the returned draft's message and id.
           Returns: Draft object, including draft id and message meta data.
        """
        try:
            # create gmail api client
            service = build('gmail', 'v1', credentials=self.creds)

            message = MIMEText(email_text)
            message['To'] = email_address
            message['From'] = EMAIL_ADDRESS
            message['Subject'] = EMAIL_SUBJECT
            encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

            create_message = {
                'raw': encoded_message
            }
            # pylint: disable=E1101
            send_message = (service.users().messages().send
                            (userId="me", body=create_message).execute())
        except HttpError as error:
            print(F'An error occurred: {error}')
            send_message = None
        return send_message


if __name__ == '__main__':
    main = Main()
    values = main.build_requests_from_sheet()
    main.send_email_to_values()
