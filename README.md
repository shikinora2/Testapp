# 🔗 Google Sheets API Test Application

Ứng dụng Python với giao diện GUI để test kết nối và thao tác với Google Sheets API.

## ✨ Tính năng

- ✅ Xác thực với Google Sheets API
- 📖 Đọc dữ liệu từ sheet
- ✍️ Ghi dữ liệu vào sheet (overwrite)
- ➕ Thêm dữ liệu vào cuối sheet (append)
- 🔄 Cập nhật cell cụ thể
- 📦 Batch update nhiều cells cùng lúc
- 🗑️ Xóa dữ liệu
- ℹ️ Xem thông tin spreadsheet
- 🎨 Giao diện đẹp mắt với Tkinter

## 📋 Yêu cầu

- Python 3.7 trở lên
- Tài khoản Google
- Google Cloud Project với Sheets API được bật

## 🚀 Hướng dẫn cài đặt

### Bước 1: Clone hoặc tải project

```bash
cd d:\Code\Testapp
```

### Bước 2: Cài đặt dependencies

```bash
pip install -r requirements.txt
```

### Bước 3: Tạo Google Cloud Project và lấy credentials

#### 3.1. Tạo Project trên Google Cloud Console

1. Truy cập [Google Cloud Console](https://console.cloud.google.com/)
2. Nhấn **"Select a project"** → **"New Project"**
3. Đặt tên project (ví dụ: "Google Sheets Test")
4. Nhấn **"Create"**

#### 3.2. Bật Google Sheets API

1. Trong project vừa tạo, vào menu **"APIs & Services"** → **"Library"**
2. Tìm kiếm **"Google Sheets API"**
3. Nhấn vào **"Google Sheets API"**
4. Nhấn **"Enable"**

#### 3.3. Tạo OAuth 2.0 Credentials

1. Vào **"APIs & Services"** → **"Credentials"**
2. Nhấn **"Create Credentials"** → **"OAuth client ID"**
3. Nếu chưa có OAuth consent screen:
   - Nhấn **"Configure Consent Screen"**
   - Chọn **"External"** → **"Create"**
   - Điền thông tin:
     - App name: `Google Sheets Test App`
     - User support email: email của bạn
     - Developer contact: email của bạn
   - Nhấn **"Save and Continue"**
   - Ở phần **Scopes**, nhấn **"Add or Remove Scopes"**
   - Tìm và chọn: `https://www.googleapis.com/auth/spreadsheets`
   - Nhấn **"Update"** → **"Save and Continue"**
   - Ở phần **Test users**, nhấn **"Add Users"**
   - Thêm email Google của bạn
   - Nhấn **"Save and Continue"**

4. Quay lại **"Credentials"** → **"Create Credentials"** → **"OAuth client ID"**
5. Chọn **Application type**: **"Desktop app"**
6. Đặt tên: `Google Sheets Desktop Client`
7. Nhấn **"Create"**
8. Nhấn **"Download JSON"** để tải file credentials
9. Đổi tên file thành `credentials.json` và copy vào thư mục project

### Bước 4: Tạo Google Spreadsheet

1. Truy cập [Google Sheets](https://sheets.google.com/)
2. Tạo một spreadsheet mới
3. Copy **Spreadsheet ID** từ URL:
   ```
   https://docs.google.com/spreadsheets/d/1ABC123xyz456/edit
                                        ^^^^^^^^^^^^^^^^
                                        Đây là Spreadsheet ID
   ```

### Bước 5: Cấu hình file .env

1. Copy file `.env.example` thành `.env`:
   ```bash
   copy .env.example .env
   ```

2. Mở file `.env` và điền thông tin:
   ```env
   SPREADSHEET_ID=1ABC123xyz456
   CREDENTIALS_FILE=credentials.json
   ```

## 🎮 Cách sử dụng

### Chạy ứng dụng

```bash
python app_gui.py
```

### Các bước test

1. **Xác thực & Kết nối**
   - Nhấn nút **"🔐 Xác thực & Kết nối"**
   - Trình duyệt sẽ mở ra, đăng nhập bằng tài khoản Google
   - Cho phép ứng dụng truy cập Google Sheets
   - Sau khi xác thực thành công, các nút khác sẽ được kích hoạt

2. **Xem thông tin Spreadsheet**
   - Nhấn **"ℹ️ Thông tin Spreadsheet"**
   - Xem tên spreadsheet, danh sách sheets, và URL

3. **Ghi dữ liệu mẫu**
   - Nhấn **"✍️ Ghi dữ liệu mẫu"**
   - Ứng dụng sẽ ghi headers và 4 dòng dữ liệu mẫu vào Sheet1

4. **Đọc dữ liệu**
   - Nhập phạm vi vào ô **"Phạm vi (Range)"** (ví dụ: `Sheet1!A1:E10`)
   - Nhấn **"📖 Đọc dữ liệu"**
   - Dữ liệu sẽ hiển thị trong khung kết quả

5. **Thêm dữ liệu**
   - Nhấn **"➕ Thêm dữ liệu"**
   - 2 dòng mới sẽ được thêm vào cuối sheet

6. **Cập nhật cell**
   - Nhấn **"🔄 Cập nhật cell"**
   - Cell D2 sẽ được cập nhật giá trị mới

7. **Batch Update**
   - Nhấn **"📦 Batch Update"**
   - Nhiều cells sẽ được cập nhật cùng lúc

8. **Xóa dữ liệu**
   - Nhập phạm vi cần xóa
   - Nhấn **"🗑️ Xóa dữ liệu"**
   - Xác nhận để xóa

## 📁 Cấu trúc project

```
Testapp/
├── app_gui.py                 # File chính - Giao diện GUI
├── google_sheets_service.py   # Service xử lý Google Sheets API
├── requirements.txt           # Dependencies
├── .env.example              # File cấu hình mẫu
├── .env                      # File cấu hình (tự tạo)
├── credentials.json          # OAuth credentials (tự tải)
├── token.json               # Token xác thực (tự động tạo)
├── .gitignore               # Git ignore
└── README.md                # Hướng dẫn này
```

## 🔧 Các API methods được sử dụng

### 1. `authenticate()`
Xác thực với Google Sheets API sử dụng OAuth 2.0

### 2. `get_spreadsheet_info()`
Lấy thông tin về spreadsheet (tên, danh sách sheets)

### 3. `read_data(range_name)`
Đọc dữ liệu từ một phạm vi cụ thể
- **Tham số**: `range_name` (ví dụ: `'Sheet1!A1:D10'`)
- **Trả về**: List of lists chứa dữ liệu

### 4. `write_data(range_name, values)`
Ghi dữ liệu vào sheet (ghi đè dữ liệu cũ)
- **Tham số**: 
  - `range_name`: Phạm vi ghi
  - `values`: List of lists chứa dữ liệu
- **Trả về**: Số cells đã cập nhật

### 5. `append_data(range_name, values)`
Thêm dữ liệu vào cuối sheet
- **Tham số**:
  - `range_name`: Phạm vi (ví dụ: `'Sheet1!A:D'`)
  - `values`: List of lists chứa dữ liệu
- **Trả về**: Số rows đã thêm

### 6. `clear_data(range_name)`
Xóa dữ liệu trong một phạm vi
- **Tham số**: `range_name`
- **Trả về**: True nếu thành công

### 7. `batch_update(data_list)`
Cập nhật nhiều phạm vi cùng lúc
- **Tham số**: List of dicts `[{'range': '...', 'values': [[...]]}]`
- **Trả về**: Tổng số cells đã cập nhật

## 🎯 Ví dụ sử dụng Range

- `Sheet1!A1:D10` - Đọc từ A1 đến D10 trong Sheet1
- `Sheet1!A:D` - Toàn bộ cột A đến D
- `Sheet1!1:5` - Toàn bộ dòng 1 đến 5
- `Sheet1!A1` - Chỉ cell A1
- `'My Sheet'!A1:B2` - Sheet có tên chứa khoảng trắng

## ⚠️ Lưu ý

1. **Lần đầu chạy**: Trình duyệt sẽ mở để xác thực. Sau đó token sẽ được lưu trong `token.json`
2. **Bảo mật**: Không commit file `credentials.json`, `.env`, và `token.json` lên Git
3. **Quyền truy cập**: Đảm bảo email test user đã được thêm vào OAuth consent screen
4. **Rate limits**: Google Sheets API có giới hạn số lượng requests. Tránh gọi quá nhiều trong thời gian ngắn

## 🐛 Xử lý lỗi thường gặp

### Lỗi: "File credentials.json not found"
- **Nguyên nhân**: Chưa tải file credentials
- **Giải pháp**: Làm theo Bước 3.3 để tải credentials.json

### Lỗi: "Access blocked: This app's request is invalid"
- **Nguyên nhân**: Chưa thêm email vào test users
- **Giải pháp**: Vào OAuth consent screen → Test users → Add users

### Lỗi: "The caller does not have permission"
- **Nguyên nhân**: Chưa bật Google Sheets API
- **Giải pháp**: Làm theo Bước 3.2

### Lỗi: "Spreadsheet not found"
- **Nguyên nhân**: Sai Spreadsheet ID hoặc không có quyền truy cập
- **Giải pháp**: Kiểm tra lại ID trong file .env

## 📚 Tài liệu tham khảo

- [Google Sheets API Documentation](https://developers.google.com/sheets/api)
- [Python Quickstart](https://developers.google.com/sheets/api/quickstart/python)
- [API Reference](https://developers.google.com/sheets/api/reference/rest)

## 📝 License

MIT License - Tự do sử dụng cho mục đích học tập và thương mại.

---

**Chúc bạn test thành công! 🎉**

