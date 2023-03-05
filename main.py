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
                
        result = service.spreadsheets().values().batchGet(
            spreadsheetId=spreadsheet_id, ranges=range_names).execute()
        ranges = result.get('valueRanges', [])
        # This returns all of Cyndi's sheets in one huge glob
        return result
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

def getMostRecentROI(title):
    # Months list
    months_list = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
    months_list.reverse()
    # Get current year
    year = datetime.date.today().year
    # Iterate through all months for each year down to 2016
    roi = None
    # Break loop when the earliest ROI is found
    while roi == None:
        # No ROI data before year 2017
        while year > 2016:
            # Go backwards through months (so we're checking most recent first) and then years
            # Loops breaks when ROI is found for the title
            for month in months_list:
                sheet_range = '%s %s!A:T' % (month, year)
                error = True
                i = 1
                try:
                    values = ReadSheet(sheet_range)
                    values.pop(0)
                    # Find which column is ROI
                    for r in range(len(values[0])):
                        if values[0][r] == 'ROI':
                            roi_column = r
                            break
                    # Find which column is book
                    for b in range(len(values[0])):
                        if values[0][b] == 'Book':
                            book_column = b
                            break
                    # Grab the ROI from the appropriate row
                    if book_column and roi_column:
                        for row in values:
                            if row[book_column] == title:
                                roi = row[roi_column]
                                break
                # There may be months in the current year that we haven't gotten to yet, have to catch those gracefully
                except Exception as e:
                    print("%s was not found. Exception: %s" % (sheet_range[:-4],e))
            # Decrease the year by one once all the months of current year are checked
            year -= 1
        # If roi is still null after the year loop is exhausted, return an N/A string for ROI
        # This will break us out of the roi == None loop as well as the function
        if not roi:
            roi = "No ROI found for this title"
            return roi
    # If the roi == None loop broke without returning, that means we have a value to return
    return roi



if __name__ == '__main__':
    #main()
    #tileVsSeriesDict()
    #print(getMostRecentROI('Chance for Love '))
    #buildCatalogueDict()
    pprint(batchReadSheet(SPREADSHEET_ID))