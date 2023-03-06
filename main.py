from __future__ import print_function

import os.path

import google.auth
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import datetime
import time

from pprint import pprint

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = '1gSRMvAb3vydmYpPZWRmHBrSJ_djQ0qVOFD_LIDcdWTg'

SAMPLE_RANGE_NAME = 'Class Data!A2:E'
CATALOGUE_RANGE_NAME = 'catalogue'

TITLE_SERIES_RANGE_NAME = 'titleSeriesNumber'

# https://docs.google.com/spreadsheets/d/1gSRMvAb3vydmYpPZWRmHBrSJ_djQ0qVOFD_LIDcdWTg/edit#gid=0

def creds():
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
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def ReadSheet(range, spreadsheetId=SPREADSHEET_ID):
    service = build('sheets', 'v4', credentials=creds())

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=spreadsheetId,
                                range=range).execute()
    values = result.get('values', [])
    if not values:
        print('No data found.')
        return
    return values

def batchReadSheet(spreadsheet_id):
    """
    Creates the batch_update the user has access to.
    Load pre-authorized user credentials from the environment.
    TODO(developer) - See https://developers.google.com/identity
    for guides on implementing OAuth2 for the application.
        """
    #creds, _ = google.auth.default()
    #creds = creds()
    # pylint: disable=maybe-no-member
    try:
        service = build('sheets', 'v4', credentials=creds())
        current_year = datetime.date.today().year
        current_month = datetime.date.today().month
        months_list = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
        months_list.reverse()
        range_names = []
        # Bunch of tricks in order to exclude future months of current year
        for y in range(current_year, 2016, -1):
            if y == current_year:
                for m in months_list[-current_month:]:
                    range_names.append('%s %s!A:T' % (m, y))
            else:
                for m in months_list:
                    range_names.append('%s %s!A:T' % (m, y))
                
        batch_cyndi = service.spreadsheets().values().batchGet(
            spreadsheetId=spreadsheet_id, ranges=range_names).execute()
        ranges = batch_cyndi.get('valueRanges', [])
        # This returns all of Cyndi's sheets in one huge glob
        return batch_cyndi
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error

def buildCatalogueDict():
    try:
        values = ReadSheet(CATALOGUE_RANGE_NAME)

        # Below dict will hold ALL THE INFO
        # Format: dict of dicts
        # book_title: {series:data, num_in_series:data ...}
        cat_dict = {}

        for listo in values:
            cat_dict.setdefault(listo[0].strip())
            cat_dict[listo[0].strip()] = {
                'series':listo[1].strip(),
                'num_in_series':listo[2].strip(),
                'author':listo[3].strip(),
                'genre':listo[4].strip(),
                'last_featured':listo[5].strip(),
                'roi':listo[6].strip()
            }

    except HttpError as err:
        print(err)
    
    return cat_dict

def populate_rois(batch_cyndi, cat_dict):
    # For each roi in cat_dict
    for title, values in cat_dict.items():
        # Find the title (ie book) column from the cyndi batch dict
        title_column_index = 0
        roi_column_index = 0
        # Want to record which sheet the ROI for the title is found in
        for sheet in batch_cyndi['valueRanges']:
            for row in sheet['values']:
                for column_number in range(len(row)):
                    if row[column_number].lower() == 'book':
                        column_number = title_column_index
                        break

        # If the title from batch matches the title from cat_dict, set the batch roi to the cat_dict roi


    return cat_dict



if __name__ == '__main__':
    #main()
    #tileVsSeriesDict()
    #print(getMostRecentROI('Chance for Love '))
    #pprint(buildCatalogueDict())
    #pprint(batchReadSheet(SPREADSHEET_ID))
    pprint(populate_rois(batchReadSheet(SPREADSHEET_ID), buildCatalogueDict()))