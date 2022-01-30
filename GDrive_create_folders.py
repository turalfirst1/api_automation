from Google import Create_Service

DRIVE_CLIENT_SECRET_FILE = 'C:\\Users\\mrtur\\OneDrive\\Desktop\\Google Sheets API\\client_secret.json'
DRIVE_API_NAME = 'drive'
DRIVE_API_VERSION = 'v3'
DRIVE_SCOPES = ['https://www.googleapis.com/auth/drive']

drive_service = Create_Service(DRIVE_CLIENT_SECRET_FILE, DRIVE_API_NAME, DRIVE_API_VERSION, DRIVE_SCOPES)

file_metadata = {
    'name': 'VAT-REPORTS-2',
    'mimeType': 'application/vnd.google-apps.folder'
    #'parents': []
}

folder1 = drive_service.files().create(body = file_metadata).execute()
folder_id = folder1['id']

print(folder1)