import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd

scope = ['https://www.googleapis.com/auth/spreadsheets']


class GoogleSheets:
    def __init__(self):
        self.creds = self.auth()

    def fetch_sheet(self, sheet_id, sheet_name):
        service = build('sheets', 'v4', credentials=self.creds)
        sheet = service.spreadsheets()
        values = sheet.values().get(spreadsheetId=sheet_id, range=sheet_name).execute()
        sheet = values.get('values', [])

        return self.to_df(sheet) if sheet else None

    @staticmethod
    def to_df(sheet):
        columns = sheet.pop(0)
        rows = [row for row in sheet]

        return pd.DataFrame(rows, columns=columns)

    @staticmethod
    def auth():
        creds = None

        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', scope)
                creds = flow.run_local_server(port=0)
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        return creds
