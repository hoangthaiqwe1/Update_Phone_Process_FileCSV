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
    """Tạo bảng response_templates nếu chưa có. Không tự thêm data mẫu."""
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
    conn.close()


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
    Import mẫu phản hồi từ file backup hoặc CSV.
    Kiểm tra trùng: nếu title HOẶC content đã tồn tại → bỏ qua.
    records: list of dict có keys: title, flow, content, tags
    Trả về: (imported_count, skipped_count)
    """
    from datetime import datetime
    now = datetime.now().strftime('%d/%m/%Y %H:%M')
    conn = get_db()
    cursor = conn.cursor()
    imported = 0
    skipped = 0
    for r in records:
        title = r.get('title', '').strip()
        content = r.get('content', '').strip()
        if not title or not content:
            skipped += 1
            continue
        # Kiểm tra trùng title hoặc content
        cursor.execute('''
            SELECT COUNT(*) FROM response_templates
            WHERE title = ? OR content = ?
        ''', (title, content))
        if cursor.fetchone()[0] > 0:
            skipped += 1
            continue
        # Không trùng → insert
        cursor.execute('''
            INSERT INTO response_templates (title, flow, content, tags, copy_count, created_at, updated_at)
            VALUES (?, ?, ?, ?, 0, ?, ?)
        ''', (title, r.get('flow', ''), content, r.get('tags', ''), now, now))
        imported += 1
    conn.commit()
    conn.close()
    return imported, skipped


# ============================================================
# QUẢN LÝ API CATALOG (Giao diện API kèm hình ảnh)
# CRUD: Thêm / Sửa / Xóa / Lấy danh sách API
# ============================================================

def init_api_catalog_db():
    """Tạo bảng api_catalog và api_images nếu chưa có."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS api_catalog (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            method TEXT NOT NULL DEFAULT 'GET',
            endpoint TEXT NOT NULL,
            description TEXT DEFAULT '',
            request_body TEXT DEFAULT '',
            response_body TEXT DEFAULT '',
            image_path TEXT DEFAULT '',
            category TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            sort_order INTEGER DEFAULT 0,
            created_at TEXT DEFAULT '',
            updated_at TEXT DEFAULT ''
        )
    ''')
    # Bảng lưu nhiều ảnh cho mỗi API
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS api_images (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            api_id INTEGER NOT NULL,
            image_path TEXT NOT NULL,
            caption TEXT DEFAULT '',
            sort_order INTEGER DEFAULT 0,
            FOREIGN KEY (api_id) REFERENCES api_catalog(id) ON DELETE CASCADE
        )
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_api_catalog_category ON api_catalog(category)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_api_catalog_method ON api_catalog(method)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_api_images_api_id ON api_images(api_id)')

    # Migrate: thêm cột mới nếu bảng đã tồn tại trước đó
    try:
        cursor.execute('ALTER TABLE api_catalog ADD COLUMN notes TEXT DEFAULT ""')
    except:
        pass
    try:
        cursor.execute('ALTER TABLE api_catalog ADD COLUMN sort_order INTEGER DEFAULT 0')
    except:
        pass

    conn.commit()
    conn.close()


def get_all_api_entries(category='', search=''):
    """
    Lấy danh sách API, có thể lọc theo category và tìm kiếm.
    Sắp xếp theo sort_order, rồi theo tên.
    """
    conn = get_db()
    cursor = conn.cursor()

    conditions = []
    params = []

    if category:
        conditions.append('category = ?')
        params.append(category)
    if search:
        like = f'%{search}%'
        conditions.append('(name LIKE ? OR endpoint LIKE ? OR description LIKE ? OR notes LIKE ?)')
        params.extend([like, like, like, like])

    where = f"WHERE {' AND '.join(conditions)}" if conditions else ""
    cursor.execute(f'SELECT * FROM api_catalog {where} ORDER BY sort_order ASC, category, name', params)

    rows = cursor.fetchall()
    data = []
    for row in rows:
        entry = {
            'id': row['id'],
            'name': row['name'],
            'method': row['method'],
            'endpoint': row['endpoint'],
            'description': row['description'],
            'request_body': row['request_body'],
            'response_body': row['response_body'],
            'image_path': row['image_path'],
            'category': row['category'],
            'notes': row['notes'] if 'notes' in row.keys() else '',
            'sort_order': row['sort_order'] if 'sort_order' in row.keys() else 0,
            'created_at': row['created_at'],
            'updated_at': row['updated_at'],
            'images': []
        }
        data.append(entry)

    # Load all images cho các API
    if data:
        api_ids = [d['id'] for d in data]
        placeholders = ','.join(['?' for _ in api_ids])
        cursor.execute(f'SELECT * FROM api_images WHERE api_id IN ({placeholders}) ORDER BY sort_order ASC, id ASC', api_ids)
        img_rows = cursor.fetchall()
        # Group images by api_id
        img_map = {}
        for img in img_rows:
            aid = img['api_id']
            if aid not in img_map:
                img_map[aid] = []
            img_map[aid].append({
                'id': img['id'],
                'image_path': img['image_path'],
                'caption': img['caption'],
                'sort_order': img['sort_order']
            })
        for entry in data:
            entry['images'] = img_map.get(entry['id'], [])

    conn.close()
    return data


def get_all_api_categories():
    """Lấy danh sách các category duy nhất."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT DISTINCT category FROM api_catalog WHERE category != "" ORDER BY category')
    rows = cursor.fetchall()
    categories = [row['category'] for row in rows]
    conn.close()
    return categories


def add_api_entry(name, method, endpoint, description, request_body, response_body, image_path, category, notes='', sort_order=0):
    """Thêm API entry mới."""
    from datetime import datetime
    now = datetime.now().strftime('%d/%m/%Y %H:%M')
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO api_catalog (name, method, endpoint, description, request_body, response_body, image_path, category, notes, sort_order, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (name, method, endpoint, description, request_body, response_body, image_path, category, notes, sort_order, now, now))
    new_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return new_id


def get_api_entry_by_id(entry_id):
    """Lấy 1 API entry theo ID kèm danh sách ảnh."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM api_catalog WHERE id = ?', (entry_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        return None

    entry = {
        'id': row['id'],
        'name': row['name'],
        'method': row['method'],
        'endpoint': row['endpoint'],
        'description': row['description'],
        'request_body': row['request_body'],
        'response_body': row['response_body'],
        'image_path': row['image_path'],
        'category': row['category'],
        'notes': row['notes'] if 'notes' in row.keys() else '',
        'sort_order': row['sort_order'] if 'sort_order' in row.keys() else 0,
        'created_at': row['created_at'],
        'updated_at': row['updated_at'],
        'images': []
    }

    # Load images
    cursor.execute('SELECT * FROM api_images WHERE api_id = ? ORDER BY sort_order ASC, id ASC', (entry_id,))
    img_rows = cursor.fetchall()
    for img in img_rows:
        entry['images'].append({
            'id': img['id'],
            'image_path': img['image_path'],
            'caption': img['caption'],
            'sort_order': img['sort_order']
        })

    conn.close()
    return entry


def update_api_entry(entry_id, name, method, endpoint, description, request_body, response_body, image_path, category, notes='', sort_order=0):
    """Cập nhật API entry theo ID."""
    from datetime import datetime
    now = datetime.now().strftime('%d/%m/%Y %H:%M')
    conn = get_db()
    cursor = conn.cursor()
    if image_path is not None:
        cursor.execute('''
            UPDATE api_catalog
            SET name = ?, method = ?, endpoint = ?, description = ?, request_body = ?,
                response_body = ?, image_path = ?, category = ?, notes = ?, sort_order = ?, updated_at = ?
            WHERE id = ?
        ''', (name, method, endpoint, description, request_body, response_body, image_path, category, notes, sort_order, now, entry_id))
    else:
        cursor.execute('''
            UPDATE api_catalog
            SET name = ?, method = ?, endpoint = ?, description = ?, request_body = ?,
                response_body = ?, category = ?, notes = ?, sort_order = ?, updated_at = ?
            WHERE id = ?
        ''', (name, method, endpoint, description, request_body, response_body, category, notes, sort_order, now, entry_id))
    conn.commit()
    affected = cursor.rowcount
    conn.close()
    return affected > 0


def delete_api_entry(entry_id):
    """Xóa API entry + xóa tất cả ảnh liên quan trong DB."""
    conn = get_db()
    cursor = conn.cursor()
    # Xóa ảnh phụ
    cursor.execute('DELETE FROM api_images WHERE api_id = ?', (entry_id,))
    # Xóa entry
    cursor.execute('DELETE FROM api_catalog WHERE id = ?', (entry_id,))
    conn.commit()
    affected = cursor.rowcount
    conn.close()
    return affected > 0


def add_api_image(api_id, image_path, caption='', sort_order=0):
    """Thêm 1 ảnh cho API entry."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO api_images (api_id, image_path, caption, sort_order)
        VALUES (?, ?, ?, ?)
    ''', (api_id, image_path, caption, sort_order))
    new_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return new_id


def delete_api_image(image_id):
    """Xóa 1 ảnh theo ID."""
    conn = get_db()
    cursor = conn.cursor()
    # Lấy path trước khi xóa
    cursor.execute('SELECT image_path FROM api_images WHERE id = ?', (image_id,))
    row = cursor.fetchone()
    path = row['image_path'] if row else None
    cursor.execute('DELETE FROM api_images WHERE id = ?', (image_id,))
    conn.commit()
    conn.close()
    return path


def get_images_by_api_id(api_id):
    """Lấy danh sách ảnh của 1 API."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM api_images WHERE api_id = ? ORDER BY sort_order ASC, id ASC', (api_id,))
    rows = cursor.fetchall()
    data = []
    for row in rows:
        data.append({
            'id': row['id'],
            'image_path': row['image_path'],
            'caption': row['caption'],
            'sort_order': row['sort_order']
        })
    conn.close()
    return data


def update_api_sort_order(entry_id, sort_order):
    """Cập nhật thứ tự hiển thị cho 1 API."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('UPDATE api_catalog SET sort_order = ? WHERE id = ?', (sort_order, entry_id))
    conn.commit()
    conn.close()


def export_all_api_entries():
    """Export toàn bộ API catalog (không kèm file ảnh, chỉ metadata)."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM api_catalog ORDER BY sort_order ASC, category, name')
    rows = cursor.fetchall()
    data = []
    for row in rows:
        entry = {
            'name': row['name'],
            'method': row['method'],
            'endpoint': row['endpoint'],
            'description': row['description'],
            'request_body': row['request_body'],
            'response_body': row['response_body'],
            'category': row['category'],
            'notes': row['notes'] if 'notes' in row.keys() else '',
            'sort_order': row['sort_order'] if 'sort_order' in row.keys() else 0,
        }
        data.append(entry)
    conn.close()
    return data


def import_api_entries(records):
    """
    Import API entries từ file JSON backup.
    Kiểm tra trùng theo endpoint + method.
    Trả về: (imported_count, skipped_count)
    """
    from datetime import datetime
    now = datetime.now().strftime('%d/%m/%Y %H:%M')
    conn = get_db()
    cursor = conn.cursor()
    imported = 0
    skipped = 0
    for r in records:
        name = r.get('name', '').strip()
        endpoint = r.get('endpoint', '').strip()
        method = r.get('method', 'GET').strip().upper()
        if not name or not endpoint:
            skipped += 1
            continue
        # Kiểm tra trùng endpoint + method
        cursor.execute('SELECT COUNT(*) FROM api_catalog WHERE endpoint = ? AND method = ?', (endpoint, method))
        if cursor.fetchone()[0] > 0:
            skipped += 1
            continue
        cursor.execute('''
            INSERT INTO api_catalog (name, method, endpoint, description, request_body, response_body, image_path, category, notes, sort_order, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, '', ?, ?, ?, ?, ?)
        ''', (name, method, endpoint, r.get('description', ''), r.get('request_body', ''),
              r.get('response_body', ''), r.get('category', ''), r.get('notes', ''),
              r.get('sort_order', 0), now, now))
        imported += 1
    conn.commit()
    conn.close()
    return imported, skipped
