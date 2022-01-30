from googleapiclient.http import MediaFileUpload
from Google import Create_Service

DRIVE_CLIENT_SECRET_FILE = 'C:\\Users\\mrtur\\OneDrive\\Desktop\\Google Sheets API\\client_secret.json'
DRIVE_API_NAME = 'drive'
DRIVE_API_VERSION = 'v3'
DRIVE_SCOPES = ['https://www.googleapis.com/auth/drive']

drive_service = Create_Service(DRIVE_CLIENT_SECRET_FILE, DRIVE_API_NAME, DRIVE_API_VERSION, DRIVE_SCOPES)

folder_id = '1dNKub4lGepsF48eixoaXYtMVwbfMomx8'

file_name = ['airflow_sandbox_vat_automation_repl.csv']
mime_type = ['text/csv']
file_medatada = {
    'name': file_name,
    'parents': [folder_id]
}

drive_service.files().create(
    body = file_medatada,
    fields = 'id'
).execute()