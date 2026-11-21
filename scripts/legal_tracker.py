# -*- coding: utf-8 -*-
import os
import json
import argparse
import sys
import io
import shutil
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
import openpyxl
from openpyxl.utils import get_column_letter

sys.path.append(os.path.dirname(os.path.abspath(__file__)))
import excel_styles

# Đảm bảo in ký tự Unicode tiếng Việt không bị lỗi trên Windows console
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# Đường dẫn mặc định đến file Excel trong cùng dự án AI branding
DEFAULT_EXCEL_PATH = r"d:\AI AGENT THUY\AI branding\reports\Danh_sach_Van_ban_Phap_luat_XNK.xlsx"
ONEDRIVE_SYNC_DIRS = [
    r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\Marketing FB 2026",
    r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\Phòng Marketing - Eureka - Marketing - Eureka\thuanh"
]

COLUMNS = [
    "STT",
    "Mã Số Văn Bản",
    "Tên Văn Bản",
    "Loại Văn Bản",
    "Cơ Quan Ban Hành",
    "Ngày Ban Hành",
    "Ngày Có Hiệu Lực",
    "Ngày Hết Hiệu Lực",
    "Văn Bản Bị Thay Thế",
    "Phân Tích Sơ Bộ",
    "So Sánh Thay Đổi",
    "Ngày Cập Nhật"
]

VALID_TYPES = ["Luật", "Nghị định", "Thông tư", "Công văn", "Quyết định", "Nghị quyết", "Văn bản hợp nhất", "QCVN", "TCVN", "Khác"]

def sanitize_text(val):
    """Làm sạch các ký tự lỗi font Unicode (nếu có)"""
    if not isinstance(val, str):
        return val
    replacements = {
        '\u0b00': 'ư',
        't\u0b00': 'tư',
        'cଁng': 'cường',
        'vật tଁ': 'vật tư',
        'chất l\u0b00ợng': 'chất lượng',
        'l\u0b00u': 'lưu',
        'n\u0b00ớc': 'nước',
        'sଁ': 'sơ',
        'kỹ thuế': 'kỹ thuật',
        'yêu cần': 'yêu cầu',
        'tr\u0b00ớc': 'trước',
        'đ\u0b00a': 'đưa',
        'tଁ nhân': 'tư nhân',
        'uy tính': 'uy tín',
        '\u0b30': 'r'
    }
    for k, v in replacements.items():
        val = val.replace(k, v)
    return val

def parse_date(date_str):
    """Hỗ trợ phân tích nhiều định dạng ngày phổ biến"""
    if date_str is None:
        return None
    if isinstance(date_str, datetime):
        return date_str.date()
    if isinstance(date_str, date):
        return date_str
    
    date_str_clean = str(date_str).strip()
    if date_str_clean in ["", "None", "Chưa xác định", "-", "N/A", "Còn hiệu lực", "Đang áp dụng", "Chưa xác định"]:
        return None
    
    part = date_str_clean.split()[0]
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(part, fmt).date()
        except ValueError:
            continue
    return None

def validate_doc(doc, index):
    """Thẩm định tính đầy đủ và hợp lệ của dữ liệu đầu vào trước khi lưu"""
    warnings = []
    ma_so = str(doc.get("ma_so", "")).strip()
    if not ma_so:
        return False, ["Mã số văn bản bị trống."]
    
    ten = str(doc.get("ten", "")).strip()
    if not ten or len(ten) < 5:
        warnings.append(f"Tên văn bản '{ten}' quá ngắn hoặc chưa chính xác.")
        
    co_quan = str(doc.get("co_quan", "")).strip()
    if not co_quan:
        warnings.append("Cơ quan ban hành bị trống.")
        
    phan_tich = str(doc.get("phan_tich", "")).strip()
    if not phan_tich or len(phan_tich) < 15:
        warnings.append("Nội dung phân tích sơ bộ quá ngắn hoặc thiếu thông tin chuyên môn.")
        
    so_sanh = str(doc.get("so_sanh", "")).strip()
    if not so_sanh or len(so_sanh) < 10:
        warnings.append("Nội dung so sánh thay đổi chưa đầy đủ.")

    return True, warnings

def format_excel_file(file_path):
    """Định dạng Excel đẹp, chuyên nghiệp bằng openpyxl"""
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    
    # Kích hoạt hiển thị đường lưới (grid lines)
    ws.views.sheetView[0].showGridLines = True
    
    # Định dạng Header
    for col_idx, col_name in enumerate(COLUMNS, 1):
        cell = ws.cell(row=1, column=col_idx)
        excel_styles.apply_header_style(cell)
    
    ws.row_dimensions[1].height = 28
    
    # Định dạng các dòng dữ liệu
    max_row = ws.max_row
    for row in range(2, max_row + 1):
        # Thiết lập STT tăng dần tự động
        ws.cell(row=row, column=1, value=row - 1)
        
        is_zebra = (row % 2 == 1)
        ws.row_dimensions[row].height = 24
        
        for col in range(1, len(COLUMNS) + 1):
            cell = ws.cell(row=row, column=col)
            
            # Căn lề và in đậm
            align_left = col not in [1, 2, 4, 6, 7, 8, 9, 12] # Tên văn bản, Cơ quan, Phân tích, So sánh
            bold = (col == 2) # Mã số văn bản
            
            excel_styles.apply_body_style(cell, is_zebra=is_zebra, align_left=align_left, bold=bold)
                
    # Thiết lập độ rộng cột tối ưu
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val_str = str(cell.value or '')
            if cell.row == 1:
                # Tiêu đề cột
                max_len = max(max_len, len(val_str) + 4)
            else:
                max_len = max(max_len, len(val_str))
        
        # Giới hạn độ rộng của các cột chứa nhiều text
        col_idx = col[0].column
        col_name = COLUMNS[col_idx - 1]
        
        if col_name in ["Tên Văn Bản"]:
            ws.column_dimensions[col_letter].width = 35
        elif col_name in ["Phân Tích Sơ Bộ", "So Sánh Thay Đổi"]:
            ws.column_dimensions[col_letter].width = 50
        elif col_name in ["Cơ Quan Ban Hành"]:
            ws.column_dimensions[col_letter].width = 20
        else:
            # Các cột ngắn tự động căn theo nội dung
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 25)
            
    tmp_path = file_path + ".tmp.xlsx"
    wb.save(tmp_path)
    wb.close()
    if os.path.exists(tmp_path):
        os.replace(tmp_path, file_path)

def process_tracker(json_data, excel_path):
    print(f"Bắt đầu xử lý tệp Excel tại: {excel_path}")
    
    existing_rows = []
    # 1. Tải file Excel cũ hoặc khởi tạo mới
    if os.path.exists(excel_path):
        try:
            wb = openpyxl.load_workbook(excel_path)
            ws = wb.active
            headers = [cell.value for cell in ws[1]]
            col_map = {col_name: idx for idx, col_name in enumerate(headers)}
            
            for r in range(2, ws.max_row + 1):
                row_data = {}
                for col in COLUMNS:
                    if col in col_map:
                        val = ws.cell(row=r, column=col_map[col] + 1).value
                        row_data[col] = sanitize_text(val)
                    else:
                        row_data[col] = None
                if any(row_data.values()):
                    existing_rows.append(row_data)
        except Exception as e:
            print(f"Lỗi khi đọc file Excel hiện tại: {e}. Tiến hành tạo file mới.")
            existing_rows = []
            
    existing_dict = {str(r.get("Mã Số Văn Bản", "")).strip(): r for r in existing_rows if r.get("Mã Số Văn Bản")}
    
    today = datetime.now().date()
    
    added_count = 0
    updated_count = 0
    replaced_codes = set()
    
    # 2. Xử lý các văn bản mới hoặc cập nhật
    for idx, doc in enumerate(json_data, 1):
        is_valid, warnings = validate_doc(doc, idx)
        ma_so = str(doc.get("ma_so", "")).strip()
        if not is_valid:
            print(f"⚠️ Bỏ qua mục #{idx} do lỗi dữ liệu: {', '.join(warnings)}")
            continue
            
        if warnings:
            print(f"⚠️ Cảnh báo kiểm định mục {ma_so}: {', '.join(warnings)}")
            
        ten = doc.get("ten", "")
        loai = doc.get("loai", "")
        co_quan = doc.get("co_quan", "")
        ngay_ban_hanh = doc.get("ngay_ban_hanh", "")
        ngay_hieu_luc = doc.get("ngay_hieu_luc", "")
        ngay_het_hieu_luc = doc.get("ngay_het_hieu_luc", "")
        van_ban_bi_thay_the = str(doc.get("van_ban_bi_thay_the", "")).strip()
        phan_tich = doc.get("phan_tich", "")
        so_sanh = doc.get("so_sanh", "")
        
        # Ghi nhận mã văn bản bị thay thế để xóa sau
        if van_ban_bi_thay_the and van_ban_bi_thay_the not in ["", "None", "N/A", "-"]:
            replaced_codes.add(van_ban_bi_thay_the)
            
        row_data = {
            "STT": None,
            "Mã Số Văn Bản": ma_so,
            "Tên Văn Bản": ten,
            "Loại Văn Bản": loai,
            "Cơ Quan Ban Hành": co_quan,
            "Ngày Ban Hành": ngay_ban_hanh,
            "Ngày Có Hiệu Lực": ngay_hieu_luc,
            "Ngày Hết Hiệu Lực": ngay_het_hieu_luc if ngay_het_hieu_luc else "Chưa xác định",
            "Văn Bản Bị Thay Thế": van_ban_bi_thay_the if van_ban_bi_thay_the else "-",
            "Phân Tích Sơ Bộ": phan_tich,
            "So Sánh Thay Đổi": so_sanh,
            "Ngày Cập Nhật": today.strftime("%Y-%m-%d")
        }
        
        if ma_so in existing_dict:
            existing_dict[ma_so] = row_data
            updated_count += 1
        else:
            existing_dict[ma_so] = row_data
            added_count += 1
        
    all_rows = list(existing_dict.values())
        
    # 3. Thực hiện tự động xóa các văn bản hết hiệu lực hoặc có hiệu lực quá 3 tháng (Theo wf04)
    rows_to_keep = []
    deleted_expired_count = 0
    deleted_replaced_count = 0
    deleted_older_3m_count = 0
    cutoff_date = today - relativedelta(months=3)
    
    for row in all_rows:
        ma_so = str(row.get("Mã Số Văn Bản", "")).strip()
        ngay_het_hieu_luc_str = row.get("Ngày Hết Hiệu Lực")
        ngay_co_hieu_luc_str = row.get("Ngày Có Hiệu Lực")
        
        # A. Kiểm tra nếu nằm trong danh sách bị thay thế bởi văn bản mới
        if ma_so in replaced_codes:
            print(f"❌ Xóa văn bản {ma_so} do bị thay thế bởi văn bản mới nhập.")
            deleted_replaced_count += 1
            continue
            
        # B. Kiểm tra nếu ngày hết hiệu lực đã qua hoặc có trạng thái hết hiệu lực
        ngay_het_hieu_luc = parse_date(ngay_het_hieu_luc_str)
        is_expired = False
        if ngay_het_hieu_luc_str and any(term in str(ngay_het_hieu_luc_str).lower() for term in ["hết hiệu lực", "bị thay thế", "bãi bỏ"]):
            is_expired = True
        elif ngay_het_hieu_luc and ngay_het_hieu_luc <= today:
            is_expired = True
            
        if is_expired:
            print(f"❌ Xóa văn bản {ma_so} do đã hết hiệu lực từ ngày {ngay_het_hieu_luc_str}.")
            deleted_expired_count += 1
            continue
            
        # C. Kiểm tra nếu ngày có hiệu lực đã quá 3 tháng so với hiện tại
        ngay_co_hieu_luc = parse_date(ngay_co_hieu_luc_str)
        if ngay_co_hieu_luc and ngay_co_hieu_luc < cutoff_date:
            print(f"❌ Xóa văn bản {ma_so} do đã có hiệu lực quá 3 tháng (từ {ngay_co_hieu_luc_str} < mốc {cutoff_date.strftime('%d/%m/%Y')}).")
            deleted_older_3m_count += 1
            continue
            
        rows_to_keep.append(row)
        
    # Đảm bảo thư mục reports tồn tại
    os.makedirs(os.path.dirname(excel_path), exist_ok=True)
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(COLUMNS)
    
    for idx, row in enumerate(rows_to_keep, 1):
        row["STT"] = idx
        row_vals = [row.get(col, "") for col in COLUMNS]
        ws.append(row_vals)
        
    # Kỹ thuật lưu nguyên tử atomic_save (tmp.xlsx -> os.replace) theo Quy tắc 10 GEMINI.md
    tmp_path = excel_path + ".tmp.xlsx"
    wb.save(tmp_path)
    wb.close()
    if os.path.exists(tmp_path):
        os.replace(tmp_path, excel_path)
    
    # Định dạng lại Excel cho chuyên nghiệp
    format_excel_file(excel_path)
    
    # Tự động đồng bộ sang các thư mục OneDrive
    sync_msgs = []
    for sync_dir in ONEDRIVE_SYNC_DIRS:
        if os.path.exists(sync_dir):
            try:
                dest_path = os.path.join(sync_dir, os.path.basename(excel_path))
                shutil.copy2(excel_path, dest_path)
                sync_msgs.append(f"☁️ Đồng bộ tức thì sang OneDrive: {dest_path}")
            except Exception as e:
                sync_msgs.append(f"⚠️ Lỗi khi đồng bộ sang OneDrive ({sync_dir}): {e}")
        else:
            sync_msgs.append(f"ℹ️ Thư mục OneDrive chưa tồn tại, bỏ qua: {sync_dir}")
    
    print("\n--- BÁO CÁO CẬP NHẬT PHÁP LUẬT ---")
    print(f"📍 Tổng số văn bản mới thêm vào: {added_count}")
    print(f"📍 Số văn bản được cập nhật thông tin: {updated_count}")
    print(f"📍 Số văn bản bị xóa do hết hiệu lực: {deleted_expired_count}")
    print(f"📍 Số văn bản bị xóa do có hiệu lực quá 3 tháng: {deleted_older_3m_count}")
    print(f"📍 Số văn bản bị xóa do bị thay thế: {deleted_replaced_count}")
    print(f"📍 Tổng số dòng dữ liệu hiện tại trong file: {len(rows_to_keep)}")
    print(f"💾 File Excel đã được lưu tại: {excel_path}")
    for msg in sync_msgs:
        print(f"{msg}")
    print("----------------------------------")

def main():
    parser = argparse.ArgumentParser(description="Quản lý và cập nhật tệp Excel theo dõi văn bản pháp luật XNK.")
    parser.add_argument("-i", "--input", required=True, help="Đường dẫn đến file JSON chứa danh sách văn bản mới.")
    parser.add_argument("-o", "--output", default=DEFAULT_EXCEL_PATH, help="Đường dẫn lưu file Excel kết quả.")
    args = parser.parse_args()
    
    if not os.path.exists(args.input):
        print(f"Lỗi: Không tìm thấy file JSON dữ liệu đầu vào tại: {args.input}")
        return
        
    try:
        with open(args.input, "r", encoding="utf-8") as f:
            json_data = json.load(f)
    except Exception as e:
        print(f"Lỗi khi đọc file JSON đầu vào: {e}")
        return
        
    process_tracker(json_data, args.output)

if __name__ == "__main__":
    main()
