import json
from json import loads
import _json
import ast
import pandas as pd
import csv
from Google import Create_Service

SHEETS_CLIENT_SECRET_FILE = 'C:\\Users\\mrtur\\OneDrive\\Desktop\\Google Sheets API\\client_secret.json'
SHEETS_API_NAME = 'sheets'
SHEETS_API_VERSION = 'v4'
SHEETS_SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

sheet_service = Create_Service(SHEETS_CLIENT_SECRET_FILE, SHEETS_API_NAME, SHEETS_API_VERSION, SHEETS_SCOPES)
spreadsheet_id = "10Ed6KGdMXy6XomaIcNsh41fT5KRJigilvuqu7gxullA"
sheets = ['AUS','EST','HUN']
sheet_ids = []
for sheet in sheets:
    response = sheet_service.spreadsheets().get(
    spreadsheetId = spreadsheet_id,
    ranges = sheet).execute()
    sheet_ids.append(response['sheets'][0]['properties']['sheetId'])

# EST Controls check:

vat_prod_posting_temp_cont = sheet_service.spreadsheets().values().get(
    spreadsheetId = spreadsheet_id,
    majorDimension = 'ROWS',
    range = 'EST!C2:C100000').execute()

request_body = {
    'requests': [
        {
            "addConditionalFormatRule":{
                "rule":{
                    "ranges":[{
                        "sheetId": 244573353,
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
                        "sheetId": 244573353,
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
                        "sheetId": 244573353,
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



