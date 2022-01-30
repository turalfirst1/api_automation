import json
from json import loads
import _json
import ast
import pandas as pd
import csv
from Google import Create_Service

DRIVE_CLIENT_SECRET_FILE = 'C:\\Users\\mrtur\\OneDrive\\Desktop\\Google Sheets API\\client_secret.json'
DRIVE_API_NAME = 'drive'
DRIVE_API_VERSION = 'v3'
DRIVE_SCOPES = ['https://www.googleapis.com/auth/drive']

SHEETS_CLIENT_SECRET_FILE = 'C:\\Users\\mrtur\\OneDrive\\Desktop\\Google Sheets API\\client_secret.json'
SHEETS_API_NAME = 'sheets'
SHEETS_API_VERSION = 'v4'
SHEETS_SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

sheet_service = Create_Service(SHEETS_CLIENT_SECRET_FILE, SHEETS_API_NAME, SHEETS_API_VERSION, SHEETS_SCOPES)
drive_service = Create_Service(DRIVE_CLIENT_SECRET_FILE, DRIVE_API_NAME, DRIVE_API_VERSION, DRIVE_SCOPES)

# 1ST STEP - CREATE FOLDER

file_metadata = {
    'name': 'VAT-REPORTS-2',
    'mimeType': 'application/vnd.google-apps.folder'
    #'parents': []
}

folder2 = drive_service.files().create(body = file_metadata).execute()
folderId = folder2['id']

# 2ND STEP - CREATE GOOGLE SHEETS FILE WITH 3 SHEETS

sheet_body = {
    'properties': {
        'title': 'AO-VAT-AUTOMATED-RESULTS-ULTIMATE-TEST',
        'locale': 'en_US',
        'autoRecalc': 'ON_CHANGE'
    },
    'sheets': [{
        'properties': {
            'title': 'ALL-TOGETHER'
        }}
    ]
}

sheets_file2 = sheet_service.spreadsheets().create(
    body = sheet_body
).execute()

worksheet_name = 'ALL-TOGETHER'
# 3RD STEP - MOVE FILE TO THE CREATED FOLDER

drive_service.files().update(
    fileId = sheets_file2['spreadsheetId'],
    addParents = folderId,
    removeParents = 'root'
).execute()

# 4TH STEP - IN CASE FETCHING DATA FROM A CSV FILE, READ THE FILE, THEN FETCH HEADERS AND ROWS SEPARATELY AND ASSIGN
# THEM TO RESPECTIVE LISTS

file = open('airflow_sandbox_vat_automation_repl.csv', encoding='utf8') # DON'T FORGET TO SPECIFY ENCODING ARGUMENT!
csvreader = csv.reader(file)
header = next(csvreader)
headers = []
for head in header:
    headers.append([head])

rows = []
for row in csvreader:
    rows.append(row)

spreadsheet_id = sheets_file2['spreadsheetId']
myspreadsheets = sheet_service.spreadsheets().get(spreadsheetId = spreadsheet_id).execute()

cell_range_header = '!A1'
value_range_body_headers = {
    'majorDimension': 'COLUMNS',
    'values': headers
}

cell_range = '!A2'

sheet_service.spreadsheets().values().update(
            spreadsheetId = spreadsheet_id,
            valueInputOption = 'USER_ENTERED',
            range = worksheet_name + cell_range_header,
            body = value_range_body_headers
        ).execute()

value_range_body = {
        'majorDimension': 'ROWS',
        'values': rows
    }

sheet_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        valueInputOption='USER_ENTERED',
        range= worksheet_name + cell_range,
        body=value_range_body
    ).execute()

# Adding a column:

value_range_body_headers = {
    'majorDimension': 'COLUMNS',
    'values': [['check_results']]
}

sheet_service.spreadsheets().values().update(
            spreadsheetId = spreadsheet_id,
            valueInputOption = 'USER_ENTERED',
            range = worksheet_name + '!W1',
            body = value_range_body_headers
        ).execute()

## Formatting:

request_body = {
    'requests': [
        {
            "addConditionalFormatRule":{
                "rule":{
                    "ranges":[{
                        "sheetId": 1386722256,
                        "startRowIndex": 1,
                        "endRowIndex":100000,
                        "startColumnIndex": 2,
                        "endColumnIndex":3
                    }],
                    "booleanRule":{
                        "condition":{
                            "type": "TEXT_CONTAINS",
                            "values": [
                                {
                                    "userEnteredValue" : 'VAT20R'
                                }
                            ]
                        },
                        "format": {
                            "backgroundColor":{
                                'red': 1,
                                'green': 0,
                                'blue': 0
                            }
                        }
                    }
                }
            }
        },
        {
            "addConditionalFormatRule":{
                "rule":{
                    "ranges":[{
                        "sheetId": 1386722256,
                        "startRowIndex": 1,
                        "endRowIndex":100000,
                        "startColumnIndex": 2,
                        "endColumnIndex":3
                    }],
                    "booleanRule":{
                        "condition":{
                            "type": "TEXT_CONTAINS",
                            "values": [
                                {
                                    "userEnteredValue" : 'VAT20'
                                }
                            ]
                        },
                        "format": {
                            "backgroundColor":{
                                'red': 1,
                                'green': 0,
                                'blue': 0
                            }
                        }
                    }
                }
            }
        },
        {
            "addConditionalFormatRule":{
                "rule":{
                    "ranges":[{
                        "sheetId": 1386722256,
                        "startRowIndex": 1,
                        "endRowIndex":100000,
                        "startColumnIndex": 2,
                        "endColumnIndex":3
                    }],
                    "booleanRule":{
                        "condition":{
                            "type": "TEXT_CONTAINS",
                            "values": [
                                {
                                    "userEnteredValue" : 'NOVAT'
                                }
                            ]
                        },
                        "format": {
                            "backgroundColor":{
                                'red': 1,
                                'green': 0,
                                'blue': 0
                            }
                        }
                    }
                }
            }
        }

    ]
}
sheet_service.spreadsheets().batchUpdate(
    spreadsheetId=spreadsheet_id,
    body=request_body
).execute()