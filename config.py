# ============================================================
# CONFIG.PY - Cấu hình chung cho toàn bộ ứng dụng
# File này chứa đường dẫn thư mục, file database,
# và danh sách tên cột để auto-detect khi đọc file Excel/CSV
# ============================================================

import os

# Thư mục gốc của project (nơi chứa file app.py)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Thư mục lưu file upload tạm thời (sẽ bị xóa mỗi lần chạy mới)
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'upload')

# Thư mục lưu file output (CSV + ZIP kết quả)
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'output')

# Thư mục lưu database lịch sử
HISTORY_FOLDER = os.path.join(BASE_DIR, 'history')

# File database SQLite lưu lịch sử (hỗ trợ hàng triệu records)
HISTORY_DB = os.path.join(HISTORY_FOLDER, 'history.db')

# File log (không dùng nữa, giữ lại cho tương thích)
LOG_FILE = os.path.join(HISTORY_FOLDER, 'ProcessLog.xlsx')

# Định dạng file được phép upload
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}

# ============================================================
# AUTO DETECT TÊN CỘT
# Khi đọc file Excel/CSV, tool sẽ tự nhận diện tên cột
# dù người dùng đặt tên khác nhau (không phân biệt hoa thường)
# VD: "nid", "NID", "Nid", "CustomerNID" đều được hiểu là cột NID
# ============================================================

# Các biến thể tên cột NID
NID_VARIANTS = ['nid', 'customernid', 'customer_nid']

# Các biến thể tên cột Result (kết quả xử lý)
RESULT_VARIANTS = ['result', 'it_result', 'it result', 'd2cresult', 'd2c_result', 'd2c result']

# Các biến thể tên cột Số điện thoại
PHONE_VARIANTS = ['phonenumber', 'phone_number', 'newphonenumber', 'new_phone_number']

# Các biến thể tên cột Ticket Pega
TICKET_VARIANTS = ['ticketpega', 'ticket_pega', 'ticket pega', 'ticket']

# ============================================================
# CẤU HÌNH ĐỘ DÀI CHUẨN CỦA PHONE VÀ NID
# Nếu file Excel lưu dạng Number (mất số 0 đầu),
# tool sẽ tự thêm số 0 cho đủ độ dài chuẩn.
# Đặt = 0 nếu KHÔNG muốn tool tự thêm (giữ nguyên file gốc)
# ============================================================

# Phone Việt Nam chuẩn 10 số (bắt đầu bằng 0)
# VD: 961338170 (9 số) → 0961338170 (10 số)
PHONE_STANDARD_LENGTH = 10

# NID chuẩn 12 số (bắt đầu bằng 0)
# VD: 25200006000 (11 số) → 025200006000 (12 số)
# Đặt = 0 nếu NID của hệ thống bạn không cố định độ dài
NID_STANDARD_LENGTH = 0

# Các biến thể tên cột Username
USERNAME_VARIANTS = ['newusername', 'new_username', 'username']
