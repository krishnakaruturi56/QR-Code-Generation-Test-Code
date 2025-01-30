# QR-Code-Generation-Test-Code


•  Generate excel sheet with structured details for each row in your Excel sheet.
•  Save these excel sheet to a directory.
•  Generate QR codes that encode the URLs of these images.
•  Save these QR codes to another directory.
•  Save the updated Excel file with the paths to these images and QR codes.
Code


Number: 348028083296ID: evident-zone-447715-d8

import os
import pandas as pd
import qrcode
import logging
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# Disable the file_cache warning
import warnings
warnings.filterwarnings('ignore', 'file_cache is only supported with oauth2client<4.0.0')

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname=s - %(message=s')

# Google Drive API setup
SCOPES = ['https://www.googleapis.com/auth/drive.file']
SERVICE_ACCOUNT_FILE = r'Z:\\MTech Quality\\Krishna\\Pythonfiles\\gauge-447714-a62babc8413d.json'

try:
    credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('drive', 'v3', credentials=credentials)
except Exception as e:
    logging.error(f"Error setting up Google Drive API: {e}")
    raise

# Function to upload file to Google Drive and set permissions
def upload_to_drive(file_path):
    try:
        file_metadata = {'name': os.path.basename(file_path)}
        media = MediaFileUpload(file_path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()

        # Set file permissions to be viewable by anyone with the link
        permission = {
            'type': 'anyone',
            'role': 'reader',
        }
        service.permissions().create(fileId=file['id'], body=permission).execute()

        file_link = f"https://drive.google.com/uc?id={file['id']}"
        return file_link
    except Exception as e:
        logging.error(f"Error uploading file to Google Drive: {e}")
        return None

# Load the main Excel file
file_path = r'Z:\\MTech Quality\\Krishna\\GaugeMasterlist.xlsx'
sheet_name = 'Gauge Master 2024'
try:
    df = pd.read_excel(file_path, sheet_name=sheet_name)
except Exception as e:
    logging.error(f"Error reading Excel file: {e}")
    raise

# Create directories for storing individual Excel files and QR codes
individual_excel_dir = 'individual_excels'
qr_code_dir = 'qr_codes'
os.makedirs(individual_excel_dir, exist_ok=True)
os.makedirs(qr_code_dir, exist_ok=True)

# Generate individual Excel files, upload to Google Drive, and create QR codes
df['QR Code'] = ''
for index, row in df.iterrows():
    details_df = pd.DataFrame([row])
    individual_file_name = f"GaugeDetails_{row['Gauge ID No. ']}.xlsx"
    individual_file_path = os.path.join(individual_excel_dir, individual_file_name)

    # Save individual Excel file
    try:
        details_df.to_excel(individual_file_path, index=False)
        logging.info(f"Individual Excel file saved to: {individual_file_path}")
    except Exception as e:
        logging.error(f"Error saving individual Excel file: {e}")
        continue

    # Upload individual Excel file to Google Drive
    upload_url = upload_to_drive(individual_file_path)
    if not upload_url:
        logging.error(f"Failed to upload file to Google Drive: {individual_file_path}")
        continue

    # Generate QR code with the file link
    try:
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )
        qr.add_data(upload_url)
        qr.make(fit=True)
        qr_img = qr.make_image(fill='black', back_color='white')
        qr_code_file = os.path.join(qr_code_dir, f"QR_{row['Gauge ID No. ']}.png")
        qr_img.save(qr_code_file)
        df.at[index, 'QR Code'] = qr_code_file
    except Exception as e:
        logging.error(f"Error generating QR code: {e}")

# Save the main updated Excel file with QR codes inserted as images
new_file_name = 'GaugeMasterlist_with_QR_Codes.xlsx'
new_file_path = os.path.join(os.path.expanduser('~'), 'Documents', new_file_name)

# Ensure the directory exists
os.makedirs(os.path.dirname(new_file_path), exist_ok=True)

try:
    with pd.ExcelWriter(new_file_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        for index, row in df.iterrows():
            qr_code_file = row['QR Code']
            if os.path.exists(qr_code_file):
                worksheet.insert_image(index + 1, len(df.columns), qr_code_file, {'x_scale': 0.5, 'y_scale': 0.5})

    logging.info(f"Final updated Excel file saved to: {new_file_path}")
except Exception as e:
    logging.error(f"Error saving main Excel file: {e}")

logging.info(f"Individual Excel files saved to: {individual_excel_dir}")
logging.info(f"QR codes saved to: {qr_code_dir}")
