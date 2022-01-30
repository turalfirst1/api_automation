from Google import Create_Service

SHEETS_CLIENT_SECRET_FILE = 'C:\\Users\\mrtur\\OneDrive\\Desktop\\Google Sheets API\\client_secret.json'
SHEETS_API_NAME = 'sheets'
SHEETS_API_VERSION = 'v4'
SHEETS_SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

sheet_service = Create_Service(SHEETS_CLIENT_SECRET_FILE, SHEETS_API_NAME, SHEETS_API_VERSION, SHEETS_SCOPES)

"""
BLANK SPREADSHEETS FILE
"""

"""
dict_keys(['spreadsheetId', 'properties', 'sheets', 'spreadsheetUrl'])
"""

sheets_file1 = sheet_service.spreadsheets().create().execute()
print(sheets_file1['spreadsheetUrl'])
print(sheets_file1['sheets'][0]['properties']['gridProperties'])
print(sheets_file1['spreadsheetId'])
print(sheets_file1['properties'])
print(sheets_file1['properties']['gridProperties'])
"""
ADVANCED EXAMPLE 
"""

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

sheets_file2 = sheet_service.spreadsheets().create(
    body = sheet_body
).execute()

