"""
Google Sheets Service - Xử lý kết nối và thao tác với Google Sheets API
"""

import os
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import pickle
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Phạm vi quyền truy cập
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


class GoogleSheetsService:
    """Class để quản lý kết nối và thao tác với Google Sheets"""
    
    def __init__(self):
        self.spreadsheet_id = os.getenv('SPREADSHEET_ID')
        self.credentials_file = os.getenv('CREDENTIALS_FILE', 'credentials.json')
        self.service = None
        self.creds = None
        
    def authenticate(self):
        """
        Xác thực với Google Sheets API
        Sử dụng OAuth 2.0 flow cho user authentication
        """
        creds = None
        
        # Token file lưu access và refresh tokens
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        
        # Nếu không có credentials hợp lệ, yêu cầu đăng nhập
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                if not os.path.exists(self.credentials_file):
                    raise FileNotFoundError(
                        f"Không tìm thấy file credentials: {self.credentials_file}\n"
                        "Vui lòng tải file credentials từ Google Cloud Console"
                    )
                
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.credentials_file, SCOPES)
                creds = flow.run_local_server(port=0)
            
            # Lưu credentials cho lần chạy tiếp theo
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
        
        self.creds = creds
        self.service = build('sheets', 'v4', credentials=creds)
        return True
    
    def get_spreadsheet_info(self):
        """Lấy thông tin về spreadsheet"""
        try:
            sheet_metadata = self.service.spreadsheets().get(
                spreadsheetId=self.spreadsheet_id
            ).execute()
            
            title = sheet_metadata.get('properties', {}).get('title', 'Unknown')
            sheets = sheet_metadata.get('sheets', [])
            sheet_names = [sheet.get('properties', {}).get('title', '') for sheet in sheets]
            
            return {
                'title': title,
                'sheets': sheet_names,
                'url': f"https://docs.google.com/spreadsheets/d/{self.spreadsheet_id}"
            }
        except HttpError as error:
            raise Exception(f"Lỗi khi lấy thông tin spreadsheet: {error}")
    
    def read_data(self, range_name):
        """
        Đọc dữ liệu từ sheet

        Args:
            range_name: Phạm vi đọc (ví dụ: 'Sheet1!A1:D10' hoặc 'Sheet1')

        Returns:
            List of lists chứa dữ liệu
        """
        try:
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range=range_name
            ).execute()

            values = result.get('values', [])
            return values
        except HttpError as error:
            error_details = str(error)
            if 'Unable to parse range' in error_details:
                raise Exception(f"Lỗi format range: '{range_name}'. Vui lòng kiểm tra lại tên sheet và format range")
            raise Exception(f"Lỗi khi đọc dữ liệu: {error}")
    
    def write_data(self, range_name, values):
        """
        Ghi dữ liệu vào sheet (ghi đè dữ liệu cũ)
        
        Args:
            range_name: Phạm vi ghi (ví dụ: 'Sheet1!A1')
            values: List of lists chứa dữ liệu cần ghi
        
        Returns:
            Số lượng cells đã cập nhật
        """
        try:
            body = {
                'values': values
            }
            result = self.service.spreadsheets().values().update(
                spreadsheetId=self.spreadsheet_id,
                range=range_name,
                valueInputOption='USER_ENTERED',
                body=body
            ).execute()
            
            return result.get('updatedCells', 0)
        except HttpError as error:
            raise Exception(f"Lỗi khi ghi dữ liệu: {error}")
    
    def append_data(self, range_name, values):
        """
        Thêm dữ liệu vào cuối sheet

        Args:
            range_name: Phạm vi thêm (ví dụ: 'Sheet1!A1' hoặc 'Sheet1')
            values: List of lists chứa dữ liệu cần thêm

        Returns:
            Số lượng rows đã thêm
        """
        try:
            body = {
                'values': values
            }
            result = self.service.spreadsheets().values().append(
                spreadsheetId=self.spreadsheet_id,
                range=range_name,
                valueInputOption='USER_ENTERED',
                insertDataOption='INSERT_ROWS',
                body=body
            ).execute()

            return result.get('updates', {}).get('updatedRows', 0)
        except HttpError as error:
            # Thêm thông tin chi tiết về lỗi
            error_details = str(error)
            if 'Unable to parse range' in error_details:
                raise Exception(f"Lỗi format range: '{range_name}'. Vui lòng sử dụng format như 'Sheet1!A1' hoặc 'Sheet1'")
            raise Exception(f"Lỗi khi thêm dữ liệu: {error}")
    
    def clear_data(self, range_name):
        """
        Xóa dữ liệu trong một phạm vi
        
        Args:
            range_name: Phạm vi xóa (ví dụ: 'Sheet1!A1:D10')
        
        Returns:
            True nếu thành công
        """
        try:
            self.service.spreadsheets().values().clear(
                spreadsheetId=self.spreadsheet_id,
                range=range_name
            ).execute()
            
            return True
        except HttpError as error:
            raise Exception(f"Lỗi khi xóa dữ liệu: {error}")
    
    def batch_update(self, data_list):
        """
        Cập nhật nhiều phạm vi cùng lúc
        
        Args:
            data_list: List of dicts với format {'range': 'Sheet1!A1', 'values': [[...]]}
        
        Returns:
            Tổng số cells đã cập nhật
        """
        try:
            batch_data = []
            for item in data_list:
                batch_data.append({
                    'range': item['range'],
                    'values': item['values']
                })
            
            body = {
                'valueInputOption': 'USER_ENTERED',
                'data': batch_data
            }
            
            result = self.service.spreadsheets().values().batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body=body
            ).execute()
            
            return result.get('totalUpdatedCells', 0)
        except HttpError as error:
            raise Exception(f"Lỗi khi batch update: {error}")

