# ============================================================
# DATABASE.PY - Quản lý database SQLite
# File này chứa tất cả hàm tương tác với database:
# - Tạo bảng, tạo index
# - Thêm dữ liệu lịch sử (history)
# - Thêm log xử lý (process_log)
# - Tìm kiếm lịch sử
# - Kiểm tra ticket trùng
# - Cập nhật kết quả (update result)
# ============================================================

import sqlite3
import os
from config import HISTORY_DB, HISTORY_FOLDER


def get_db():
    """
    Kết nối đến database SQLite.
    Tự động tạo thư mục nếu chưa có.
    """
    os.makedirs(HISTORY_FOLDER, exist_ok=True)
    conn = sqlite3.connect(HISTORY_DB)
    conn.row_factory = sqlite3.Row  # Cho phép truy cập cột bằng tên
    return conn


def init_db():
    """
    Khởi tạo database: tạo bảng và index nếu chưa có.
    Được gọi 1 lần khi app khởi động.
    """
    conn = get_db()
    cursor = conn.cursor()

    # ----- BẢNG HISTORY -----
    # Lưu toàn bộ lịch sử xử lý phone (mỗi dòng = 1 NID đã xử lý)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            NewPhoneNumber TEXT DEFAULT '',
            NID TEXT DEFAULT '',
            NewUsername TEXT DEFAULT '',
            TicketPega TEXT DEFAULT '',
            TicketXuLy TEXT DEFAULT '',
            DateUpdate TEXT DEFAULT '',
            D2CResult TEXT DEFAULT '',
            ITResult TEXT DEFAULT ''
        )
    ''')

    # ----- BẢNG PROCESS LOG -----
    # Lưu log mỗi lần chạy tool (thành công/lỗi, thời gian, file nào)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS process_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            DateTime TEXT DEFAULT '',
            File TEXT DEFAULT '',
            Status TEXT DEFAULT '',
            Error TEXT DEFAULT '',
            Duration TEXT DEFAULT ''
        )
    ''')

    # ----- INDEX -----
    # Tạo index để tìm kiếm nhanh (quan trọng khi dữ liệu lên triệu records)
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_history_nid ON history(NID)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_history_ticket_xu_ly ON history(TicketXuLy)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_history_phone ON history(NewPhoneNumber)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_history_ticket_pega ON history(TicketPega)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_history_date ON history(DateUpdate)')

    conn.commit()
    conn.close()


# ============================================================
# KIỂM TRA TRÙNG TÊN TICKET
# Dùng khi Process mới - nếu tên đã có thì không cho chạy lại
# ============================================================

def check_ticket_exists(ticket_name):
    """
    Kiểm tra tên ticket đã tồn tại trong history chưa.
    Trả về True nếu đã có (không cho xử lý trùng).
    """
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT COUNT(*) FROM history WHERE TicketXuLy = ?', (ticket_name,))
    count = cursor.fetchone()[0]
    conn.close()
    return count > 0


# ============================================================
# LƯU DỮ LIỆU MỚI
# ============================================================

def save_history_records(records):
    """
    Thêm nhiều dòng mới vào bảng history.
    Dùng khi Process xong 1 cặp file D2C + Result.
    records: list of dict với keys: NewPhoneNumber, NID, NewUsername, 
             TicketPega, TicketXuLy, DateUpdate, D2CResult, ITResult
    """
    conn = get_db()
    cursor = conn.cursor()
    cursor.executemany('''
        INSERT INTO history (NewPhoneNumber, NID, NewUsername, TicketPega, TicketXuLy, DateUpdate, D2CResult, ITResult)
        VALUES (:NewPhoneNumber, :NID, :NewUsername, :TicketPega, :TicketXuLy, :DateUpdate, :D2CResult, :ITResult)
    ''', records)
    conn.commit()
    conn.close()


def save_log_records(records):
    """
    Thêm log xử lý vào bảng process_log.
    records: list of dict với keys: DateTime, File, Status, Error, Duration
    """
    conn = get_db()
    cursor = conn.cursor()
    cursor.executemany('''
        INSERT INTO process_log (DateTime, File, Status, Error, Duration)
        VALUES (:DateTime, :File, :Status, :Error, :Duration)
    ''', records)
    conn.commit()
    conn.close()


# ============================================================
# TÌM KIẾM LỊCH SỬ (dùng cho trang History)
# Hỗ trợ tìm theo NID, Phone, Ticket, Date, Result
# Có phân trang để không load hết triệu records 1 lúc
# ============================================================

def search_history(search='', page=1, per_page=50):
    """
    Tìm kiếm lịch sử với phân trang.
    search: từ khóa tìm (NID, phone, ticket, date,...)
    page: trang hiện tại
    per_page: số dòng mỗi trang
    Trả về: (data, total) - danh sách records và tổng số kết quả
    """
    conn = get_db()
    cursor = conn.cursor()
    offset = (page - 1) * per_page

    if search:
        like = f'%{search}%'
        # Đếm tổng số kết quả tìm được
        cursor.execute('''
            SELECT COUNT(*) FROM history 
            WHERE NID LIKE ? OR NewPhoneNumber LIKE ? OR TicketPega LIKE ? 
            OR TicketXuLy LIKE ? OR DateUpdate LIKE ? OR ITResult LIKE ?
        ''', (like, like, like, like, like, like))
        total = cursor.fetchone()[0]

        # Lấy dữ liệu trang hiện tại (sắp xếp mới nhất trước)
        cursor.execute('''
            SELECT * FROM history 
            WHERE NID LIKE ? OR NewPhoneNumber LIKE ? OR TicketPega LIKE ? 
            OR TicketXuLy LIKE ? OR DateUpdate LIKE ? OR ITResult LIKE ?
            ORDER BY id DESC
            LIMIT ? OFFSET ?
        ''', (like, like, like, like, like, like, per_page, offset))
    else:
        # Không tìm kiếm - lấy tất cả
        cursor.execute('SELECT COUNT(*) FROM history')
        total = cursor.fetchone()[0]
        cursor.execute('SELECT * FROM history ORDER BY id DESC LIMIT ? OFFSET ?', (per_page, offset))

    # Chuyển kết quả sang list of dict
    rows = cursor.fetchall()
    data = []
    for row in rows:
        data.append({
            'id': row['id'],
            'NewPhoneNumber': row['NewPhoneNumber'],
            'NID': row['NID'],
            'NewUsername': row['NewUsername'],
            'TicketPega': row['TicketPega'],
            'Ticket xử lý': row['TicketXuLy'],
            'Date Update': row['DateUpdate'],
            'D2C Result': row['D2CResult'],
            'IT Result': row['ITResult']
        })

    conn.close()
    return data, total


# ============================================================
# LẤY LOG XỬ LÝ (dùng cho trang Process Log)
# ============================================================

def get_log(search='', page=1, per_page=50):
    """
    Lấy danh sách log xử lý với phân trang và tìm kiếm.
    search: từ khóa tìm (ticket, NID, thao tác,...)
    """
    conn = get_db()
    cursor = conn.cursor()
    offset = (page - 1) * per_page

    if search:
        like = f'%{search}%'
        cursor.execute('''
            SELECT COUNT(*) FROM process_log
            WHERE DateTime LIKE ? OR File LIKE ? OR Status LIKE ? OR Error LIKE ?
        ''', (like, like, like, like))
        total = cursor.fetchone()[0]

        cursor.execute('''
            SELECT * FROM process_log
            WHERE DateTime LIKE ? OR File LIKE ? OR Status LIKE ? OR Error LIKE ?
            ORDER BY id DESC LIMIT ? OFFSET ?
        ''', (like, like, like, like, per_page, offset))
    else:
        cursor.execute('SELECT COUNT(*) FROM process_log')
        total = cursor.fetchone()[0]
        cursor.execute('SELECT * FROM process_log ORDER BY id DESC LIMIT ? OFFSET ?', (per_page, offset))

    rows = cursor.fetchall()
    data = []
    for row in rows:
        data.append({
            'id': row['id'],
            'DateTime': row['DateTime'],
            'File': row['File'],
            'Status': row['Status'],
            'Error': row['Error'],
            'Duration': row['Duration']
        })

    conn.close()
    return data, total


# ============================================================
# QUẢN LÝ TICKET (dùng cho tính năng Update Result)
# ============================================================

def get_all_ticket_names():
    """
    Lấy danh sách tất cả tên ticket đã xử lý (không trùng).
    Dùng cho dropdown chọn ticket khi Update.
    """
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT DISTINCT TicketXuLy FROM history WHERE TicketXuLy != "" ORDER BY id DESC')
    rows = cursor.fetchall()
    tickets = [row['TicketXuLy'] for row in rows]
    conn.close()
    return tickets


def get_records_by_ticket(ticket_name):
    """
    Lấy tất cả records thuộc 1 ticket cụ thể.
    Dùng để preview dữ liệu và export lại CSV.
    """
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM history WHERE TicketXuLy = ? ORDER BY id', (ticket_name,))
    rows = cursor.fetchall()
    data = []
    for row in rows:
        data.append({
            'id': row['id'],
            'NewPhoneNumber': row['NewPhoneNumber'],
            'NID': row['NID'],
            'NewUsername': row['NewUsername'],
            'TicketPega': row['TicketPega'],
            'TicketXuLy': row['TicketXuLy'],
            'DateUpdate': row['DateUpdate'],
            'D2CResult': row['D2CResult'],
            'ITResult': row['ITResult']
        })
    conn.close()
    return data


# ============================================================
# CẬP NHẬT KẾT QUẢ (UPDATE RESULT)
# Khi dev trả kết quả sai, dùng hàm này để update lại
# IT Result đúng cho từng NID trong ticket đã chạy trước đó
# ============================================================

def update_it_result_by_ticket(ticket_name, nid_result_map, new_date):
    """
    Cập nhật IT Result cho các records thuộc ticket_name.
    
    Params:
        ticket_name: tên ticket cần update
        nid_result_map: dict {NID: result_mới} từ file Result mới
        new_date: ngày cập nhật mới
    
    Trả về: (updated_count, not_found_count)
        - updated_count: số NID đã update thành công
        - not_found_count: số NID không tìm thấy trong file Result mới
    """
    conn = get_db()
    cursor = conn.cursor()
    updated_count = 0
    not_found_count = 0

    # Lấy tất cả records của ticket này
    cursor.execute('SELECT id, NID FROM history WHERE TicketXuLy = ?', (ticket_name,))
    rows = cursor.fetchall()

    for row in rows:
        nid = row['NID']
        if nid in nid_result_map:
            # Tìm thấy NID trong file Result mới → cập nhật kết quả
            cursor.execute('''
                UPDATE history SET ITResult = ?, DateUpdate = ? WHERE id = ?
            ''', (nid_result_map[nid], new_date, row['id']))
            updated_count += 1
        else:
            # Không tìm thấy NID trong file Result mới
            cursor.execute('''
                UPDATE history SET ITResult = 'NOT FOUND', DateUpdate = ? WHERE id = ?
            ''', (new_date, row['id']))
            not_found_count += 1

    conn.commit()
    conn.close()
    return updated_count, not_found_count


# ============================================================
# CRUD - SỬA / XÓA RECORD TRONG HISTORY (chỉnh tay)
# ============================================================

def update_history_record(record_id, data):
    """
    Cập nhật 1 record trong history theo ID.
    data: dict chứa các field cần update
    VD: {'ITResult': 'Success', 'NewPhoneNumber': '0961338170'}
    """
    conn = get_db()
    cursor = conn.cursor()

    # Chỉ update các field được truyền vào
    allowed_fields = [
        'NewPhoneNumber', 'NID', 'NewUsername', 'TicketPega',
        'TicketXuLy', 'DateUpdate', 'D2CResult', 'ITResult'
    ]
    set_parts = []
    values = []
    for key, val in data.items():
        if key in allowed_fields:
            set_parts.append(f'{key} = ?')
            values.append(val)

    if not set_parts:
        conn.close()
        return False

    values.append(record_id)
    sql = f"UPDATE history SET {', '.join(set_parts)} WHERE id = ?"
    cursor.execute(sql, values)
    conn.commit()
    affected = cursor.rowcount
    conn.close()
    return affected > 0


def delete_history_record(record_id):
    """Xóa 1 record trong history theo ID."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM history WHERE id = ?', (record_id,))
    conn.commit()
    affected = cursor.rowcount
    conn.close()
    return affected > 0


def delete_history_by_ticket(ticket_name):
    """Xóa toàn bộ records của 1 ticket (khi cần xóa sạch để chạy lại)."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM history WHERE TicketXuLy = ?', (ticket_name,))
    conn.commit()
    affected = cursor.rowcount
    conn.close()
    return affected


def get_history_record_by_id(record_id):
    """Lấy 1 record theo ID."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM history WHERE id = ?', (record_id,))
    row = cursor.fetchone()
    conn.close()
    if row:
        return {
            'id': row['id'],
            'NewPhoneNumber': row['NewPhoneNumber'],
            'NID': row['NID'],
            'NewUsername': row['NewUsername'],
            'TicketPega': row['TicketPega'],
            'TicketXuLy': row['TicketXuLy'],
            'DateUpdate': row['DateUpdate'],
            'D2CResult': row['D2CResult'],
            'ITResult': row['ITResult']
        }
    return None


def delete_history_by_tickets(ticket_names):
    """
    Xóa nhiều ticket cùng lúc.
    ticket_names: list các tên ticket cần xóa
    Trả về: tổng số records đã xóa
    """
    conn = get_db()
    cursor = conn.cursor()
    total = 0
    for name in ticket_names:
        cursor.execute('DELETE FROM history WHERE TicketXuLy = ?', (name,))
        total += cursor.rowcount
    conn.commit()
    conn.close()
    return total


# ============================================================
# QUẢN LÝ MẪU PHẢN HỒI (RESPONSE TEMPLATES)
# CRUD: Thêm / Sửa / Xóa / Lấy danh sách mẫu phản hồi
# Lưu vào SQLite để quản lý từ giao diện web
# ============================================================

def init_responses_db():
    """Tạo bảng response_templates nếu chưa có. Seed dữ liệu mẫu nếu bảng trống."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS response_templates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            flow TEXT NOT NULL,
            content TEXT NOT NULL,
            tags TEXT DEFAULT '',
            copy_count INTEGER DEFAULT 0,
            created_at TEXT DEFAULT '',
            updated_at TEXT DEFAULT ''
        )
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_responses_flow ON response_templates(flow)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_responses_tags ON response_templates(tags)')
    conn.commit()

    # Seed dữ liệu mẫu nếu bảng trống
    cursor.execute('SELECT COUNT(*) FROM response_templates')
    if cursor.fetchone()[0] == 0:
        _seed_default_responses(cursor)
        conn.commit()

    conn.close()


def _seed_default_responses(cursor):
    """Thêm các mẫu phản hồi mặc định."""
    from datetime import datetime
    now = datetime.now().strftime('%d/%m/%Y %H:%M')
    defaults = [
        ("Báo khách hàng check lại", "FEOL", "Dear Team ITSD,\n\nNhờ Team báo khách hàng check lại giúp em.\n\nThanks Team.", "check,retry"),
        ("KYC hết hạn - Yêu cầu update KYC", "FEOL", "Dear Team ITSD,\n\nNhờ Team báo khách hàng update KYC rồi check lại giúp em - Kyc khách hàng hết hạn.\n\nThanks Team.", "kyc"),
        ("KYC hết hạn - Quét NFC lại", "FEOL", "Dear Team ITSD,\n\nNhờ Team báo khách hàng update KYC rồi quét NFC lại giúp em\n\nThanks Team.", "kyc,nfc"),
        ("Quên tên đăng nhập + mật khẩu", "FEOL", "Dear Team ITSD,\n\nNhờ Team báo khách hàng thực hiện cả 2 tính năng quên tên đăng nhập và quên mật khẩu củng lúc. Sau đó cho em xin video thao tác nếu có lỗi\n\nThanks Team.", "login,password"),
        ("NID đăng ký với Phone 3 số đuôi", "FEOL", "Dear Team ITSD,\n\nFEOL kiểm tra NID khách hàng đang đăng ký với Phone có 3 sô đuôi là: \n\nThanks Team!", "nid,phone"),
        ("Xác minh danh tính thất bại - Điểm thấp", "FEOL", "Dear Team ITSD,\n\nTrường Hợp khách hàng xác minh danh tính thất bại. Do điểm xác minh danh tính khách hàng thấp. Trường hợp cần lấy Logs vui lòng cho em xin ngày thực hiên.\n\nThanks Team!", "ekyc,logs"),
        ("App loading lâu - Chờ chọn tài khoản", "FEOL", "Dear Team ITSD,\n\nNhờ team báo khách hàng thử lại giúp em. Hiện tại App FEOL đang loading lâu. khách hàng vui lòng đợi và chọn số tài khoản giúp em.\n\nThanks Team!", "loading,retry"),
        ("Quên TĐN + MK rồi login lại", "FEOL", "Dear Team ITSD,\n\nNhờ anh báo khách hàng thực hiện cả 2 tính năng quên tên đăng nhập và quên mật khẩu cùng 1 lúc. Sau đó thử login lại giúp em.\n\nThanks Team.", "login,password"),
        ("SĐT không khớp - Cancel hồ sơ", "FEOL", "Dear Team ITSD,\n\nSố điện thoại lên hồ sơ không khớp với FEOL. Vui lòng cancel lên lại / Sau 15 ngày SLA sẽ tự động cancel.\n\nThanks Team.", "phone,cancel"),
        ("Yêu cầu cung cấp thông tin lấy Logs", "FEOL", "Dear Anh/Chị,\n\nĐể hỗ trợ lấy logs chính xác và kiểm tra cho khách hàng, vui lòng cung cấp thông tin cho FEOL sau:\n1 Thời gian xảy ra lỗi\n2 Màn hình/flow mà người dùng đang thao tác khi gặp lỗi (hoặc các bước để tái hiện nếu có).\n3 Screenshot hoặc video của lỗi.\n\nLý do: FEOL là ứng dụng facing người dùng nên từ phía hệ thống không thể xác định chính xác màn hình hoặc luồng thao tác mà người dùng đang thực hiện. Nếu chỉ cung cấp khoảng thời gian quá dài, việc quét toàn bộ logs sẽ phát sinh chi phí xử lý lớn và khó đảm bảo lấy đúng bản ghi liên quan. Việc cung cấp thời gian và màn hình xảy ra lỗi sẽ giúp team xác định đúng logs cần kiểm tra và rút ngắn thời gian điều tra.\n\nCảm ơn Anh/Chị!", "logs,info"),
        ("UBank - CIF Number existed in OCB", "UBank", "Dear anh/chị,\n\nOwner: VPBank\nIssue: Khách hàng thực hiện ký eSign TopUp, hệ thống báo lỗi UBank.\nRoot Cause: VPBank trả về mã lỗi \"CIF Number existed in OCB\".\nStatus: Null\n\nHiện tại, đối với các trường hợp thông tin giữa T24 và OCB profile không trùng khớp, khách hàng tạm thời chưa thể thực hiện giải ngân qua UBank.\n\nTeam đang phối hợp các bên liên quan để xây dựng quy trình xử lý ngoại lệ cho các trường hợp này. Sau khi có hướng dẫn chính thức, các case cần hỗ trợ tiếp tục vui lòng gửi request lại để team kiểm tra và hỗ trợ xử lý.\n\nTrong thời gian chờ cập nhật quy trình mới, IT xin phép tạm thời đóng ticket hiện tại cho đến khi có thông báo tiếp theo.\n\nCảm ơn anh/chị.", "esign,cif,topup"),
        ("UBank - FAIL_OCB_DIFFERENCE_PID_OR_PHONE", "UBank", "Dear anh/chị,\n\nOwner: VPBank\nIssue: Khách hàng thực hiện ký eSign TopUp, hệ thống báo lỗi UBank.\nRoot Cause: VPBank trả về mã lỗi \"FAIL_OCB_DIFFERENCE_PID_OR_PHONE\".\nStatus: Null\n\nHiện tại, đối với các trường hợp thông tin giữa T24 và OCB profile không trùng khớp, khách hàng tạm thời chưa thể thực hiện giải ngân qua UBank.\n\nTeam đang phối hợp các bên liên quan để xây dựng quy trình xử lý ngoại lệ cho các trường hợp này. Sau khi có hướng dẫn chính thức, các case cần hỗ trợ tiếp tục vui lòng gửi request lại để team kiểm tra và hỗ trợ xử lý.\n\nTrong thời gian chờ cập nhật quy trình mới, IT xin phép tạm thời đóng ticket hiện tại cho đến khi có thông báo tiếp theo.\n\nCảm ơn anh/chị.", "esign,pid,phone,topup"),
        ("UBank - Trạng thái INITIAL (eSign/Đăng ký TK)", "UBank", "Dear anh/chị,\n\nIssue: Khách hàng thực hiện ký esign \"báo có lỗi xãy ra\"./ Đăng ký tài khoản ubank\nRoot Cause: VPBank trả về trạng thái tài khoản khách hàng là \"INITIAL\", dẫn đến hệ thống không thể thực hiện xác thực thông tin tài khoản UBank. Nguyên nhân được xác định là do dữ liệu khách hàng giữa hệ thống T24 và OCB Profile chưa đồng bộ hoặc không khớp nhau.\n\nAction Required:\nNhờ CS hỗ trợ kiểm tra trạng thái tài khoản UBank của khách hàng trên hệ thống T24.\nTrường hợp tài khoản vẫn ở trạng thái Active trên T24, vui lòng hướng dẫn khách hàng đến quầy giao dịch VPBank để kiểm tra và cập nhật lại thông tin tài khoản (NID, SĐT, MAIL).\nTrường hợp tài khoản vẫn ở trạng thái Inactive trên T24 hiện tại chưa thể hỗ trợ được. Vui lòng check thêm thông tin với CS. FEOL sẽ hỗ trợ cung cấp thông tin tài khoản Ubank của khách hàng nếu cs cần kiểm tra.\n\nCảm ơn anh/chị.", "esign,initial,t24"),
        ("UBank - Trạng thái Fail (eSign/Đăng ký TK)", "UBank", "Dear anh/chị,\n\nIssue: Khách hàng thực hiện ký esign \"báo có lỗi xãy ra\"./ Đăng ký tài khoản ubank\nRoot Cause: VPBank trả về trạng thái tài khoản khách hàng là \"Fail\", dẫn đến hệ thống không thể thực hiện xác thực thông tin tài khoản UBank. Nguyên nhân được xác định là do dữ liệu khách hàng giữa hệ thống T24 và OCB Profile chưa đồng bộ hoặc không khớp nhau.\n\nAction Required:\nNhờ CS hỗ trợ kiểm tra trạng thái tài khoản UBank của khách hàng trên hệ thống T24.\nTrường hợp tài khoản vẫn ở trạng thái Active trên T24, vui lòng hướng dẫn khách hàng đến quầy giao dịch VPBank để kiểm tra và cập nhật lại thông tin tài khoản (NID, SĐT, MAIL).\nTrường hợp tài khoản vẫn ở trạng thái Inactive trên T24 hiện tại chưa thể hỗ trợ được. Vui lòng check thêm thông tin với CS. FEOL sẽ hỗ trợ cung cấp thông tin tài khoản Ubank của khách hàng nếu cs cần kiểm tra.\n\nCảm ơn anh/chị.", "esign,fail,t24"),
        ("UBank - Phone number already exists", "UBank", "Dear anh/chị,\n\nDear team,\nOwner: VPBank\nIssue: Khách hàng mở TK Ubank\nRoot Cause: VPBank trả về mã lỗi \"Phone number already exists.\".\nStatus: Null\nAction: Hướng dẫn khách hàng ra quầy VPBank để kiểm tra thông tin.\nETA: Null\n\nBest Regards", "phone,ubank,register"),
        ("UBank - Email format is invalid", "UBank", "Dear anh/chị,\n\nOwner: Vpbank\nIssue: Lỗi không ký được esign \nRoot cause: Vpbank báo lỗi Email format is invalid\nStatus: null\nAction: Nhờ khách hàng ra quầy Vpbank kiểm tra lại email ở feol và Vpbank. Định dạng mail phải là @gmail.com\nETA: null\n\nBest Regards", "esign,email"),
    ]
    for title, flow, content, tags in defaults:
        cursor.execute('''
            INSERT INTO response_templates (title, flow, content, tags, copy_count, created_at, updated_at)
            VALUES (?, ?, ?, ?, 0, ?, ?)
        ''', (title, flow, content, tags, now, now))


def get_all_responses(flow='', search=''):
    """
    Lấy danh sách mẫu phản hồi, có thể lọc theo luồng và tìm kiếm.
    Sắp xếp theo copy_count giảm dần (hay dùng nhất lên đầu).
    """
    conn = get_db()
    cursor = conn.cursor()

    conditions = []
    params = []

    if flow:
        conditions.append('flow = ?')
        params.append(flow)
    if search:
        like = f'%{search}%'
        conditions.append('(title LIKE ? OR content LIKE ? OR tags LIKE ?)')
        params.extend([like, like, like])

    where = f"WHERE {' AND '.join(conditions)}" if conditions else ""
    cursor.execute(f'SELECT * FROM response_templates {where} ORDER BY copy_count DESC, id DESC', params)

    rows = cursor.fetchall()
    data = []
    for row in rows:
        data.append({
            'id': row['id'],
            'title': row['title'],
            'flow': row['flow'],
            'content': row['content'],
            'tags': row['tags'],
            'copy_count': row['copy_count'],
            'created_at': row['created_at'],
            'updated_at': row['updated_at']
        })
    conn.close()
    return data


def get_all_flows():
    """Lấy danh sách các luồng (flow) duy nhất."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT DISTINCT flow FROM response_templates ORDER BY flow')
    rows = cursor.fetchall()
    flows = [row['flow'] for row in rows]
    conn.close()
    return flows


def add_response(title, flow, content, tags=''):
    """Thêm mẫu phản hồi mới."""
    from datetime import datetime
    now = datetime.now().strftime('%d/%m/%Y %H:%M')
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO response_templates (title, flow, content, tags, copy_count, created_at, updated_at)
        VALUES (?, ?, ?, ?, 0, ?, ?)
    ''', (title, flow, content, tags, now, now))
    conn.commit()
    new_id = cursor.lastrowid
    conn.close()
    return new_id


def update_response(response_id, title, flow, content, tags=''):
    """Cập nhật mẫu phản hồi theo ID."""
    from datetime import datetime
    now = datetime.now().strftime('%d/%m/%Y %H:%M')
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE response_templates 
        SET title = ?, flow = ?, content = ?, tags = ?, updated_at = ?
        WHERE id = ?
    ''', (title, flow, content, tags, now, response_id))
    conn.commit()
    affected = cursor.rowcount
    conn.close()
    return affected > 0


def delete_response(response_id):
    """Xóa mẫu phản hồi theo ID."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM response_templates WHERE id = ?', (response_id,))
    conn.commit()
    affected = cursor.rowcount
    conn.close()
    return affected > 0


def increment_copy_count(response_id):
    """Tăng số lần copy (dùng khi user click Copy)."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('UPDATE response_templates SET copy_count = copy_count + 1 WHERE id = ?', (response_id,))
    conn.commit()
    conn.close()


def get_response_by_id(response_id):
    """Lấy 1 mẫu phản hồi theo ID."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM response_templates WHERE id = ?', (response_id,))
    row = cursor.fetchone()
    conn.close()
    if row:
        return {
            'id': row['id'],
            'title': row['title'],
            'flow': row['flow'],
            'content': row['content'],
            'tags': row['tags'],
            'copy_count': row['copy_count'],
            'created_at': row['created_at'],
            'updated_at': row['updated_at']
        }
    return None


def export_all_responses():
    """Export toàn bộ mẫu phản hồi (dùng cho backup/chia sẻ)."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM response_templates ORDER BY flow, id')
    rows = cursor.fetchall()
    data = []
    for row in rows:
        data.append({
            'id': row['id'],
            'title': row['title'],
            'flow': row['flow'],
            'content': row['content'],
            'tags': row['tags'],
            'copy_count': row['copy_count'],
            'created_at': row['created_at'],
            'updated_at': row['updated_at']
        })
    conn.close()
    return data


def import_responses(records):
    """
    Import mẫu phản hồi từ file backup.
    records: list of dict có keys: title, flow, content, tags
    Trả về: số records đã import
    """
    from datetime import datetime
    now = datetime.now().strftime('%d/%m/%Y %H:%M')
    conn = get_db()
    cursor = conn.cursor()
    count = 0
    for r in records:
        cursor.execute('''
            INSERT INTO response_templates (title, flow, content, tags, copy_count, created_at, updated_at)
            VALUES (?, ?, ?, ?, 0, ?, ?)
        ''', (r.get('title', ''), r.get('flow', ''), r.get('content', ''), r.get('tags', ''), now, now))
        count += 1
    conn.commit()
    conn.close()
    return count
