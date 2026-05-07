import io
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
import config

SCOPES = ['https://www.googleapis.com/auth/drive']

def get_drive_service():
    creds = service_account.Credentials.from_service_account_info(
        config.SERVICE_ACCOUNT_INFO, scopes=SCOPES)
    service = build('drive', 'v3', credentials=creds)
    return service

XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
GOOGLE_SHEET_MIME = 'application/vnd.google-apps.spreadsheet'


def download_template(file_id=config.DRIVE_TEMPLATE_FILE_ID):
    """Downloads the Excel template from Google Drive into a BytesIO object.
    Handles both native .xlsx (binary) and Google Sheets (export to .xlsx).
    `supportsAllDrives=True` ensures files in Shared Drives are accessible."""
    service = get_drive_service()

    meta = service.files().get(
        fileId=file_id,
        fields='mimeType,name',
        supportsAllDrives=True,
    ).execute()
    mime = meta.get('mimeType', '')

    if mime == GOOGLE_SHEET_MIME:
        request = service.files().export_media(fileId=file_id, mimeType=XLSX_MIME)
    else:
        request = service.files().get_media(fileId=file_id, supportsAllDrives=True)

    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _status, done = downloader.next_chunk()

    fh.seek(0)
    return fh

def upload_output(file_path, filename, folder_id=config.DRIVE_OUTPUT_FOLDER_ID):
    """Uploads the completed Excel file to the specified Google Drive folder."""
    service = get_drive_service()
    
    file_metadata = {
        'name': filename,
        'parents': [folder_id]
    }
    media = MediaFileUpload(file_path,
                            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            resumable=True)
                            
    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id',
        supportsAllDrives=True,
    ).execute()
    return file.get('id')
