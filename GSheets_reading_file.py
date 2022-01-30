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

"""
Example 1. Get method (single cell range of values)
"""
spreadsheet_id = "10Ed6KGdMXy6XomaIcNsh41fT5KRJigilvuqu7gxullA"
response = sheet_service.spreadsheets().values().get(
    spreadsheetId = spreadsheet_id,
    majorDimension = 'ROWS',
    range = "AUS!A1:Z5"
).execute()

print(response)
print(response.keys())
print(response['range'])
print(response['majorDimension'])
print(response['values'])

columns = response['values'][0]
data = response['values'][1:]
df = pd.DataFrame(data, columns = columns)
print(df)

df2 = df.set_index('vat_enty_no_temp_cont')
print(df2)

"""
Example 2. Get method 
"""

response = sheet_service.spreadsheets().values().get(
    spreadsheetId = spreadsheet_id,
    majorDimension = 'ROWS',
    range = "AUS"
).execute()

response['values'][:10]
columns = response['values'][0]
data = response['values'][1:]
df = pd.DataFrame(data, columns = columns)

"""
Example 3. batchGet Method
"""

valueRanges_body = [
    "AUS!A1:D4",
    "EST!A1:D7",
    "HUN!A1:D8"
]

response1 = sheet_service.spreadsheets().values().batchGet(
    spreadsheetId = spreadsheet_id,
    majorDimension = 'ROWS',
    ranges = valueRanges_body
).execute()

print(response1.keys())
print(response1["valueRanges"])

dataset = {}
for item in response1["valueRanges"]:
    dataset[item['range']] = item['values']

print(dataset)
print(dataset['AUS!A1:D4'])


df = {}
for indx, k in enumerate(dataset):
    columns = dataset[k]
    data = dataset[k][1:]
    df[indx] = pd.DataFrame(data, columns = columns)
pd.set_option("display.max_columns", None)

df[0]

