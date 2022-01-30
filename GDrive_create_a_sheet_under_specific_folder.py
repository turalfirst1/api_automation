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
    'name': 'VAT-REPORTS',
    'mimeType': 'application/vnd.google-apps.folder'
    #'parents': []
}

folder1 = drive_service.files().create(body = file_metadata).execute()
folderId = folder1['id']

# 2ND STEP - CREATE GOOGLE SHEETS FILE WITH 3 SHEETS

sheet_body = {
    'properties': {
        'title': 'AO-VAT-AUTOMATED-RESULTS',
        'locale': 'en_US',
        'autoRecalc': 'ON_CHANGE'
    },
    'sheets': [{
        'properties': {
            'title': 'AUS'
        }},
        {
            'properties': {
                'title': 'EST'
            }},
{
        'properties': {
            'title': 'HUN'
        }}
    ]
}

sheets_file1 = sheet_service.spreadsheets().create(
    body = sheet_body
).execute()

# 3RD STEP - MOVE FILE TO THE CREATED FOLDER

drive_service.files().update(
    fileId = sheets_file1['spreadsheetId'],
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

aus_rows = rows[:9679]

est_rows = rows[9679:19358]

hun_rows = rows[19358:]

spreadsheet_id = sheets_file1['spreadsheetId']
myspreadsheets = sheet_service.spreadsheets().get(spreadsheetId = spreadsheet_id).execute()

worksheet_names = ['AUS', 'EST', 'HUN']
cell_range = '!A2'

# 5TH STEP - ADD HEADERS TO SHEETS

cell_range_header = '!A1'
value_range_body_headers = {
    'majorDimension': 'COLUMNS',
    'values': headers
}
for worksheet_name in worksheet_names:
    sheet_service.spreadsheets().values().update(
            spreadsheetId = spreadsheet_id,
            valueInputOption = 'USER_ENTERED',
            range = worksheet_name + cell_range_header,
            body = value_range_body_headers
        ).execute()

# 6TH STEP: POPULATING Sheets with proper data
dictionary = {
    'AUS': aus_rows,
    'EST': est_rows,
    'HUN': hun_rows
}


for key, value in dictionary.items():
    value_range_body = {
        'majorDimension': 'ROWS',
        'values': value
    }
    sheet_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        valueInputOption='USER_ENTERED',
        range= key + cell_range,
        body=value_range_body
    ).execute()

# sheet_service.spreadsheets().values().clear(
#         spreadsheetId = spreadsheet_id,
#         range = 'AUS!A2:10000000'
#     ).execute()
