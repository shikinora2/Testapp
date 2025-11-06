"""
Google Sheets Test App - Giao diện GUI với Tkinter
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
from google_sheets_service import GoogleSheetsService
from datetime import datetime


class GoogleSheetsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Google Sheets API Test Application")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        self.service = GoogleSheetsService()
        self.is_authenticated = False
        
        self.setup_ui()
        
    def setup_ui(self):
        """Thiết lập giao diện người dùng"""
        
        # Header
        header_frame = tk.Frame(self.root, bg="#4285f4", height=60)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(
            header_frame,
            text="🔗 Google Sheets API Test Application",
            font=("Arial", 16, "bold"),
            bg="#4285f4",
            fg="white"
        )
        title_label.pack(pady=15)
        
        # Main container
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Connection section
        conn_frame = tk.LabelFrame(main_frame, text="Kết nối", font=("Arial", 10, "bold"), padx=10, pady=10)
        conn_frame.pack(fill=tk.X, pady=(0, 10))

        # Spreadsheet URL/ID input
        url_frame = tk.Frame(conn_frame)
        url_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(url_frame, text="Google Sheet URL hoặc ID:", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=5)

        self.sheet_url_entry = tk.Entry(url_frame, font=("Arial", 9), width=60)
        self.sheet_url_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # Load spreadsheet ID from .env if exists
        if self.service.spreadsheet_id:
            self.sheet_url_entry.insert(0, self.service.spreadsheet_id)

        self.update_sheet_button = tk.Button(
            url_frame,
            text="📝 Cập nhật",
            command=self.update_spreadsheet_id,
            bg="#4285f4",
            fg="white",
            font=("Arial", 9, "bold"),
            padx=10,
            pady=5,
            cursor="hand2"
        )
        self.update_sheet_button.pack(side=tk.LEFT, padx=5)

        # Help text
        help_text = tk.Label(
            url_frame,
            text="💡",
            font=("Arial", 9),
            fg="gray",
            cursor="hand2"
        )
        help_text.pack(side=tk.LEFT, padx=2)
        help_text.bind("<Button-1>", lambda e: self.show_url_help())

        # Buttons frame
        buttons_frame = tk.Frame(conn_frame)
        buttons_frame.pack(fill=tk.X)

        self.auth_button = tk.Button(
            buttons_frame,
            text="🔐 Xác thực & Kết nối",
            command=self.authenticate,
            bg="#34a853",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=20,
            pady=10,
            cursor="hand2"
        )
        self.auth_button.pack(side=tk.LEFT, padx=5)

        self.info_button = tk.Button(
            buttons_frame,
            text="ℹ️ Thông tin Spreadsheet",
            command=self.get_info,
            bg="#fbbc04",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=20,
            pady=10,
            cursor="hand2",
            state=tk.DISABLED
        )
        self.info_button.pack(side=tk.LEFT, padx=5)

        self.status_label = tk.Label(
            buttons_frame,
            text="⚪ Chưa kết nối",
            font=("Arial", 9),
            fg="gray"
        )
        self.status_label.pack(side=tk.LEFT, padx=20)
        
        # Operations section
        ops_frame = tk.LabelFrame(main_frame, text="Thao tác với dữ liệu", font=("Arial", 10, "bold"), padx=10, pady=10)
        ops_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Row 1
        row1_frame = tk.Frame(ops_frame)
        row1_frame.pack(fill=tk.X, pady=5)
        
        self.read_button = tk.Button(
            row1_frame,
            text="📖 Đọc dữ liệu",
            command=self.read_data,
            bg="#4285f4",
            fg="white",
            font=("Arial", 9, "bold"),
            padx=15,
            pady=8,
            cursor="hand2",
            state=tk.DISABLED,
            width=20
        )
        self.read_button.pack(side=tk.LEFT, padx=5)
        
        self.write_button = tk.Button(
            row1_frame,
            text="✍️ Ghi dữ liệu mẫu",
            command=self.write_sample_data,
            bg="#34a853",
            fg="white",
            font=("Arial", 9, "bold"),
            padx=15,
            pady=8,
            cursor="hand2",
            state=tk.DISABLED,
            width=20
        )
        self.write_button.pack(side=tk.LEFT, padx=5)
        
        self.append_button = tk.Button(
            row1_frame,
            text="➕ Thêm dữ liệu",
            command=self.append_data,
            bg="#ea4335",
            fg="white",
            font=("Arial", 9, "bold"),
            padx=15,
            pady=8,
            cursor="hand2",
            state=tk.DISABLED,
            width=20
        )
        self.append_button.pack(side=tk.LEFT, padx=5)
        
        # Row 2
        row2_frame = tk.Frame(ops_frame)
        row2_frame.pack(fill=tk.X, pady=5)
        
        self.update_button = tk.Button(
            row2_frame,
            text="🔄 Cập nhật cell",
            command=self.update_cell,
            bg="#fbbc04",
            fg="white",
            font=("Arial", 9, "bold"),
            padx=15,
            pady=8,
            cursor="hand2",
            state=tk.DISABLED,
            width=20
        )
        self.update_button.pack(side=tk.LEFT, padx=5)
        
        self.batch_button = tk.Button(
            row2_frame,
            text="📦 Batch Update",
            command=self.batch_update,
            bg="#9c27b0",
            fg="white",
            font=("Arial", 9, "bold"),
            padx=15,
            pady=8,
            cursor="hand2",
            state=tk.DISABLED,
            width=20
        )
        self.batch_button.pack(side=tk.LEFT, padx=5)
        
        self.clear_button = tk.Button(
            row2_frame,
            text="🗑️ Xóa dữ liệu",
            command=self.clear_data,
            bg="#f44336",
            fg="white",
            font=("Arial", 9, "bold"),
            padx=15,
            pady=8,
            cursor="hand2",
            state=tk.DISABLED,
            width=20
        )
        self.clear_button.pack(side=tk.LEFT, padx=5)
        
        # Range input
        range_frame = tk.Frame(ops_frame)
        range_frame.pack(fill=tk.X, pady=10)

        tk.Label(range_frame, text="Tên Sheet:", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=5)

        self.sheet_name_entry = tk.Entry(range_frame, font=("Arial", 9), width=20)
        self.sheet_name_entry.insert(0, "Sheet1")
        self.sheet_name_entry.pack(side=tk.LEFT, padx=5)

        self.get_sheets_button = tk.Button(
            range_frame,
            text="📋 Lấy danh sách",
            command=self.list_sheets,
            bg="#9c27b0",
            fg="white",
            font=("Arial", 8),
            padx=8,
            pady=3,
            cursor="hand2",
            state=tk.DISABLED
        )
        self.get_sheets_button.pack(side=tk.LEFT, padx=5)

        tk.Label(range_frame, text="Phạm vi:", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=(10, 5))

        self.range_entry = tk.Entry(range_frame, font=("Arial", 9), width=20)
        self.range_entry.insert(0, "A1:E10")
        self.range_entry.pack(side=tk.LEFT, padx=5)

        tk.Label(range_frame, text="(Ví dụ: A1:E10 hoặc để trống)", font=("Arial", 8), fg="gray").pack(side=tk.LEFT)
        
        # Output section
        output_frame = tk.LabelFrame(main_frame, text="Kết quả", font=("Arial", 10, "bold"), padx=10, pady=10)
        output_frame.pack(fill=tk.BOTH, expand=True)
        
        self.output_text = scrolledtext.ScrolledText(
            output_frame,
            font=("Consolas", 9),
            wrap=tk.WORD,
            bg="#f5f5f5",
            fg="#333333"
        )
        self.output_text.pack(fill=tk.BOTH, expand=True)
        
        # Clear output button
        clear_output_btn = tk.Button(
            output_frame,
            text="🧹 Xóa log",
            command=self.clear_output,
            bg="#607d8b",
            fg="white",
            font=("Arial", 8),
            cursor="hand2"
        )
        clear_output_btn.pack(pady=5)
        
    def log(self, message, level="INFO"):
        """Ghi log vào output text"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        if level == "ERROR":
            prefix = "❌"
        elif level == "SUCCESS":
            prefix = "✅"
        elif level == "INFO":
            prefix = "ℹ️"
        else:
            prefix = "📝"
        
        log_message = f"[{timestamp}] {prefix} {message}\n"
        self.output_text.insert(tk.END, log_message)
        self.output_text.see(tk.END)
        self.root.update()
        
    def clear_output(self):
        """Xóa nội dung output"""
        self.output_text.delete(1.0, tk.END)

    def extract_spreadsheet_id(self, url_or_id):
        """
        Trích xuất Spreadsheet ID từ URL hoặc trả về ID nếu đã là ID

        Args:
            url_or_id: URL đầy đủ hoặc Spreadsheet ID

        Returns:
            Spreadsheet ID
        """
        import re

        # Nếu là URL đầy đủ
        if 'docs.google.com/spreadsheets' in url_or_id:
            # Pattern: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/...
            match = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', url_or_id)
            if match:
                return match.group(1)

        # Nếu đã là ID (hoặc không match pattern)
        return url_or_id.strip()

    def update_spreadsheet_id(self):
        """Cập nhật Spreadsheet ID từ input"""
        try:
            url_or_id = self.sheet_url_entry.get().strip()

            if not url_or_id:
                messagebox.showwarning("Cảnh báo", "Vui lòng nhập URL hoặc ID của Google Sheet!")
                return

            # Trích xuất ID
            spreadsheet_id = self.extract_spreadsheet_id(url_or_id)

            # Cập nhật service
            self.service.spreadsheet_id = spreadsheet_id

            # Cập nhật lại entry với ID đã trích xuất
            self.sheet_url_entry.delete(0, tk.END)
            self.sheet_url_entry.insert(0, spreadsheet_id)

            self.log(f"Đã cập nhật Spreadsheet ID: {spreadsheet_id}", "SUCCESS")

            # Nếu đã xác thực, reset trạng thái để xác thực lại với sheet mới
            if self.is_authenticated:
                self.log("Vui lòng xác thực lại để kết nối với sheet mới", "INFO")
                self.is_authenticated = False
                self.status_label.config(text="⚪ Chưa kết nối", fg="gray")
                self.auth_button.config(text="🔐 Xác thực & Kết nối", bg="#34a853", state=tk.NORMAL)

                # Disable các nút khác
                self.info_button.config(state=tk.DISABLED)
                self.read_button.config(state=tk.DISABLED)
                self.write_button.config(state=tk.DISABLED)
                self.append_button.config(state=tk.DISABLED)
                self.update_button.config(state=tk.DISABLED)
                self.batch_button.config(state=tk.DISABLED)
                self.clear_button.config(state=tk.DISABLED)
                self.get_sheets_button.config(state=tk.DISABLED)

        except Exception as e:
            self.log(f"Lỗi khi cập nhật Spreadsheet ID: {str(e)}", "ERROR")
            messagebox.showerror("Lỗi", f"Không thể cập nhật Spreadsheet ID:\n{str(e)}")

    def show_url_help(self):
        """Hiển thị hướng dẫn về URL/ID"""
        help_message = """📋 HƯỚNG DẪN NHẬP GOOGLE SHEET

Bạn có thể nhập một trong hai dạng:

1️⃣ URL đầy đủ:
https://docs.google.com/spreadsheets/d/1ABC123xyz456/edit

2️⃣ Chỉ Spreadsheet ID:
1ABC123xyz456

💡 Cách lấy URL/ID:
- Mở Google Sheet của bạn
- Copy URL từ thanh địa chỉ trình duyệt
- Hoặc chỉ copy phần ID giữa "/d/" và "/edit"

Ví dụ:
https://docs.google.com/spreadsheets/d/1ABC123xyz456/edit
                                        ^^^^^^^^^^^^^^^^
                                        Đây là ID
"""
        messagebox.showinfo("Hướng dẫn", help_message)
        
    def enable_buttons(self):
        """Kích hoạt các nút sau khi xác thực"""
        self.info_button.config(state=tk.NORMAL)
        self.read_button.config(state=tk.NORMAL)
        self.write_button.config(state=tk.NORMAL)
        self.append_button.config(state=tk.NORMAL)
        self.update_button.config(state=tk.NORMAL)
        self.batch_button.config(state=tk.NORMAL)
        self.clear_button.config(state=tk.NORMAL)
        self.get_sheets_button.config(state=tk.NORMAL)
        
    def authenticate(self):
        """Xác thực với Google Sheets API"""
        def auth_thread():
            try:
                self.log("Đang xác thực với Google Sheets API...")
                self.auth_button.config(state=tk.DISABLED, text="Đang xác thực...")
                
                self.service.authenticate()
                
                self.is_authenticated = True
                self.status_label.config(text="🟢 Đã kết nối", fg="green")
                self.log("Xác thực thành công!", "SUCCESS")
                self.auth_button.config(text="✅ Đã kết nối", bg="#34a853")
                
                self.enable_buttons()
                
            except Exception as e:
                self.log(f"Lỗi xác thực: {str(e)}", "ERROR")
                self.auth_button.config(state=tk.NORMAL, text="🔐 Xác thực & Kết nối")
                self.status_label.config(text="🔴 Lỗi kết nối", fg="red")
        
        threading.Thread(target=auth_thread, daemon=True).start()

    def get_info(self):
        """Lấy thông tin spreadsheet"""
        def info_thread():
            try:
                self.log("Đang lấy thông tin spreadsheet...")
                info = self.service.get_spreadsheet_info()

                self.log(f"Tên: {info['title']}", "SUCCESS")
                self.log(f"Sheets: {', '.join(info['sheets'])}", "INFO")
                self.log(f"URL: {info['url']}", "INFO")

            except Exception as e:
                self.log(f"Lỗi: {str(e)}", "ERROR")

        threading.Thread(target=info_thread, daemon=True).start()

    def list_sheets(self):
        """Lấy danh sách các sheets và cho phép chọn"""
        def list_thread():
            try:
                self.log("Đang lấy danh sách sheets...")
                info = self.service.get_spreadsheet_info()
                sheets = info['sheets']

                if not sheets:
                    self.log("Không tìm thấy sheet nào", "INFO")
                    return

                self.log(f"Tìm thấy {len(sheets)} sheet(s):", "SUCCESS")
                for i, sheet in enumerate(sheets, 1):
                    self.log(f"  {i}. {sheet}")

                # Tạo dialog để chọn sheet
                self.root.after(0, lambda: self.show_sheet_selector(sheets))

            except Exception as e:
                self.log(f"Lỗi: {str(e)}", "ERROR")

        threading.Thread(target=list_thread, daemon=True).start()

    def show_sheet_selector(self, sheets):
        """Hiển thị dialog chọn sheet"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Chọn Sheet")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()

        tk.Label(dialog, text="Chọn sheet để làm việc:", font=("Arial", 10, "bold")).pack(pady=10)

        # Listbox
        listbox_frame = tk.Frame(dialog)
        listbox_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        scrollbar = tk.Scrollbar(listbox_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        listbox = tk.Listbox(listbox_frame, font=("Arial", 10), yscrollcommand=scrollbar.set)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)

        for sheet in sheets:
            listbox.insert(tk.END, sheet)

        # Select current sheet if exists
        current_sheet = self.sheet_name_entry.get().strip()
        if current_sheet in sheets:
            listbox.selection_set(sheets.index(current_sheet))

        def on_select():
            selection = listbox.curselection()
            if selection:
                selected_sheet = listbox.get(selection[0])
                self.sheet_name_entry.delete(0, tk.END)
                self.sheet_name_entry.insert(0, selected_sheet)
                self.log(f"Đã chọn sheet: {selected_sheet}", "SUCCESS")
                dialog.destroy()

        def on_double_click(event):
            on_select()

        listbox.bind('<Double-Button-1>', on_double_click)

        # Buttons
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text="Chọn", command=on_select, bg="#4285f4", fg="white",
                 font=("Arial", 9, "bold"), padx=20, pady=5, cursor="hand2").pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Hủy", command=dialog.destroy, bg="#666", fg="white",
                 font=("Arial", 9, "bold"), padx=20, pady=5, cursor="hand2").pack(side=tk.LEFT, padx=5)

    def get_full_range(self):
        """Tạo full range từ sheet name và range"""
        sheet_name = self.sheet_name_entry.get().strip()
        range_part = self.range_entry.get().strip()

        if not sheet_name:
            sheet_name = "Sheet1"

        if range_part:
            return f"{sheet_name}!{range_part}"
        else:
            return sheet_name

    def read_data(self):
        """Đọc dữ liệu từ sheet"""
        def read_thread():
            try:
                range_name = self.get_full_range()
                self.log(f"Đang đọc dữ liệu từ {range_name}...")

                data = self.service.read_data(range_name)

                if not data:
                    self.log("Không có dữ liệu trong phạm vi này", "INFO")
                    return

                self.log(f"Đọc được {len(data)} dòng:", "SUCCESS")
                self.log("─" * 80)

                for i, row in enumerate(data, 1):
                    self.log(f"Dòng {i}: {' | '.join(str(cell) for cell in row)}")

                self.log("─" * 80)

            except Exception as e:
                self.log(f"Lỗi: {str(e)}", "ERROR")

        threading.Thread(target=read_thread, daemon=True).start()

    def write_sample_data(self):
        """Ghi dữ liệu mẫu vào sheet"""
        def write_thread():
            try:
                self.log("Đang ghi dữ liệu mẫu...")

                sheet_name = self.sheet_name_entry.get().strip() or "Sheet1"

                # Headers
                headers = [['ID', 'Họ tên', 'Email', 'Tuổi', 'Thành phố']]
                cells_updated = self.service.write_data(f'{sheet_name}!A1:E1', headers)
                self.log(f"Đã ghi header ({cells_updated} cells)", "SUCCESS")

                # Sample data
                sample_data = [
                    ['1', 'Nguyễn Văn A', 'nguyenvana@email.com', '25', 'Hà Nội'],
                    ['2', 'Trần Thị B', 'tranthib@email.com', '30', 'TP.HCM'],
                    ['3', 'Lê Văn C', 'levanc@email.com', '28', 'Đà Nẵng'],
                    ['4', 'Phạm Thị D', 'phamthid@email.com', '22', 'Cần Thơ'],
                ]

                cells_updated = self.service.write_data(f'{sheet_name}!A2:E5', sample_data)
                self.log(f"Đã ghi {len(sample_data)} dòng dữ liệu ({cells_updated} cells)", "SUCCESS")

                self.log("✅ Hoàn thành ghi dữ liệu mẫu!", "SUCCESS")

            except Exception as e:
                self.log(f"Lỗi: {str(e)}", "ERROR")

        threading.Thread(target=write_thread, daemon=True).start()

    def append_data(self):
        """Thêm dữ liệu mới vào cuối sheet"""
        def append_thread():
            try:
                self.log("Đang thêm dữ liệu mới...")

                sheet_name = self.sheet_name_entry.get().strip() or "Sheet1"
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                new_data = [
                    ['5', f'Người dùng mới {timestamp}', 'newuser@email.com', '27', 'Hải Phòng'],
                    ['6', f'Test User {timestamp}', 'testuser@email.com', '24', 'Huế'],
                ]

                # Sử dụng format đơn giản cho append
                rows_added = self.service.append_data(f'{sheet_name}!A1', new_data)
                self.log(f"Đã thêm {rows_added} dòng mới", "SUCCESS")

            except Exception as e:
                self.log(f"Lỗi: {str(e)}", "ERROR")

        threading.Thread(target=append_thread, daemon=True).start()

    def update_cell(self):
        """Cập nhật một cell cụ thể"""
        def update_thread():
            try:
                self.log("Đang cập nhật cell...")

                sheet_name = self.sheet_name_entry.get().strip() or "Sheet1"
                # Update tuổi của người đầu tiên
                cells_updated = self.service.write_data(f'{sheet_name}!D2', [['26']])
                self.log(f"Đã cập nhật cell D2 thành '26' ({cells_updated} cells)", "SUCCESS")

            except Exception as e:
                self.log(f"Lỗi: {str(e)}", "ERROR")

        threading.Thread(target=update_thread, daemon=True).start()

    def batch_update(self):
        """Cập nhật nhiều vị trí cùng lúc"""
        def batch_thread():
            try:
                self.log("Đang thực hiện batch update...")

                sheet_name = self.sheet_name_entry.get().strip() or "Sheet1"
                batch_data = [
                    {'range': f'{sheet_name}!E2', 'values': [['Hà Nội (Updated)']]},
                    {'range': f'{sheet_name}!E3', 'values': [['TP.HCM (Updated)']]},
                    {'range': f'{sheet_name}!E4', 'values': [['Đà Nẵng (Updated)']]},
                ]

                cells_updated = self.service.batch_update(batch_data)
                self.log(f"Đã cập nhật {cells_updated} cells", "SUCCESS")

            except Exception as e:
                self.log(f"Lỗi: {str(e)}", "ERROR")

        threading.Thread(target=batch_thread, daemon=True).start()

    def clear_data(self):
        """Xóa dữ liệu"""
        def clear_thread():
            try:
                result = messagebox.askyesno(
                    "Xác nhận",
                    "Bạn có chắc muốn xóa dữ liệu trong phạm vi này?"
                )

                if not result:
                    return

                range_name = self.get_full_range()
                self.log(f"Đang xóa dữ liệu từ {range_name}...")

                self.service.clear_data(range_name)
                self.log(f"Đã xóa dữ liệu từ {range_name}", "SUCCESS")

            except Exception as e:
                self.log(f"Lỗi: {str(e)}", "ERROR")

        threading.Thread(target=clear_thread, daemon=True).start()


def main():
    root = tk.Tk()
    app = GoogleSheetsApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

