# -*- coding: utf-8 -*-
import os
import sys
import io
import json
import argparse
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

# Đảm bảo in tiếng Việt trên console không bị lỗi
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# Import styles từ module dùng chung
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
import excel_styles

DEFAULT_INPUT_JSON = r"d:\AI AGENT THUY\AI branding\scratch\weekly_calendar_input.json"
DEFAULT_OUTPUT_EXCEL = r"d:\AI AGENT THUY\AI branding\reports\plan1\Weekly_Content_Calendar.xlsx"

COLUMNS = [
    "STT",
    "Ngày đăng",
    "Nền tảng",
    "Pillar",
    "Persona Giả định",
    "Định dạng",
    "Góc tiếp cận",
    "Option A (Chủ đề chính)",
    "Option B (Dự phòng 1)",
    "Option C (Dự phòng 2)",
    "Nguồn SSoT / Cốt truyện",
    "Trạng thái Duyệt"
]

DEFAULT_DATA = [
    {
        "Ngày đăng": "Thứ Hai (15/06)",
        "Nền tảng": "LinkedIn",
        "Pillar": "P1-AI",
        "Persona Giả định": "Anh Dũng (Founder)",
        "Định dạng": "Slide/Carousel",
        "Góc tiếp cận": "Góc 3: AI Vận hành",
        "Option A (Chủ đề chính)": "Đừng May Rủi Khi Khai Hải Quan: Quy Trình AI check HS Code",
        "Option B (Dự phòng 1)": "Hướng dẫn xây dựng AI kiểm tra HS Code tự động cho doanh nghiệp",
        "Option C (Dự phòng 2)": "Tự động hóa khâu tra cứu HS Code nhập khẩu bằng AI",
        "Nguồn SSoT / Cốt truyện": "reports/SEO_Keyword_Trends_Eureka.xlsx",
        "Trạng thái Duyệt": "Chờ duyệt"
    },
    {
        "Ngày đăng": "Thứ Ba (16/06)",
        "Nền tảng": "Facebook",
        "Pillar": "P2-XNK",
        "Persona Giả định": "Anh Minh (Chủ shop)",
        "Định dạng": "Bài viết + Ảnh",
        "Góc tiếp cận": "Góc 2: Giải pháp",
        "Option A (Chủ đề chính)": "Đêm Muộn Mếu Máo Vì Bị Phạt 50 Triệu Mã HS",
        "Option B (Dự phòng 1)": "Bài học xương máu: Nhập khẩu sai mã HS phụ tùng và cái kết bị phạt 50 triệu",
        "Option C (Dự phòng 2)": "3 mẹo đơn giản tránh bị phạt mã HS khi thông quan chính ngạch",
        "Nguồn SSoT / Cốt truyện": "SEO Eureka Col 4",
        "Trạng thái Duyệt": "Chờ duyệt"
    },
    {
        "Ngày đăng": "Thứ Tư (17/06)",
        "Nền tảng": "TikTok",
        "Pillar": "P2-XNK",
        "Persona Giả định": "Anh Minh (Chủ shop)",
        "Định dạng": "Clip ngắn < 60s",
        "Góc tiếp cận": "Góc 1: Cảnh báo",
        "Option A (Chủ đề chính)": "Cảnh báo ghi sai nhãn phụ theo NĐ 37",
        "Option B (Dự phòng 1)": "Coi chừng bị tịch thu hàng do nhãn phụ sai quy định Nghị định 37",
        "Option C (Dự phòng 2)": "Lỗi ghi nhãn phụ hải quan hay bắt lỗi doanh nghiệp XNK nhất",
        "Nguồn SSoT / Cốt truyện": "NĐ 37/2020/NĐ-CP",
        "Trạng thái Duyệt": "Chờ duyệt"
    },
    {
        "Ngày đăng": "Thứ Năm (18/06)",
        "Nền tảng": "Facebook",
        "Pillar": "P4-LIFE",
        "Persona Giả định": "Anh Minh (Chủ shop)",
        "Định dạng": "Kể chuyện + Ảnh",
        "Góc tiếp cận": "Góc 4: Đơn giản hóa",
        "Option A (Chủ đề chính)": "Nhật ký một ngày xử lý sự cố kẹt cảng Cát Lái",
        "Option B (Dự phòng 1)": "Chuyện cứu container trễ giờ tàu chạy tại Cát Lái và niềm vui của người gác cửa chính ngạch",
        "Option C (Dự phòng 2)": "Phía sau những chuyến hàng cập bến: Một ngày lăn lộn cùng khách hàng tại cảng Cát Lái",
        "Nguồn SSoT / Cốt truyện": "daily_notes.md L42",
        "Trạng thái Duyệt": "Chờ duyệt"
    }
]

def load_data(json_path):
    """Tải dữ liệu từ tệp JSON nếu có, ngược lại dùng default mock data"""
    if json_path and os.path.exists(json_path):
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                print(f"Đã tải thành công dữ liệu từ: {json_path}")
                return data
        except Exception as e:
            print(f"Lỗi khi đọc file JSON đầu vào ({e}). Sử dụng dữ liệu mặc định.")
    else:
        print(f"Không tìm thấy file JSON đầu vào tại {json_path}. Sử dụng dữ liệu mặc định.")
    return DEFAULT_DATA

def create_calendar_excel(input_json, output_excel, week_title="TUẦN 24/2026"):
    print(f"Bắt đầu khởi tạo file Lịch Excel tại: {output_excel}")
    
    # Tạo thư mục đầu ra nếu chưa tồn tại
    os.makedirs(os.path.dirname(output_excel), exist_ok=True)
    
    # Load data
    data = load_data(input_json)
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Weekly Calendar"
    
    # Bật hiển thị đường lưới (grid lines)
    ws.views.sheetView[0].showGridLines = True
    
    # 2. Ghi tiêu đề lớn cho Sheet
    ws.merge_cells("A1:L1")
    title_cell = ws["A1"]
    title_cell.value = f"LỊCH BIÊN TẬP TUẦN - HỆ THỐNG PBAN ({week_title})"
    title_cell.font = excel_styles.TITLE_FONT
    title_cell.alignment = excel_styles.ALIGN_LEFT
    ws.row_dimensions[1].height = 35
    
    # Ghi chú nhỏ
    ws.merge_cells("A2:L2")
    note_cell = ws["A2"]
    note_cell.value = "Hướng dẫn: Nhấp vào cột 'Trạng thái Duyệt' để chọn Option được duyệt thông qua dropdown."
    note_cell.font = excel_styles.NOTE_FONT
    note_cell.alignment = excel_styles.ALIGN_LEFT
    ws.row_dimensions[2].height = 20
    
    # 3. Ghi Header (Dòng 4)
    ws.row_dimensions[4].height = 30
    for col_idx, col_name in enumerate(COLUMNS, 1):
        cell = ws.cell(row=4, column=col_idx, value=col_name)
        excel_styles.apply_header_style(cell)
        
    # 4. Ghi dữ liệu
    start_row = 5
    for idx, item in enumerate(data):
        row_idx = start_row + idx
        ws.row_dimensions[row_idx].height = 35 # Chiều cao dòng rộng rãi
        
        # Điền STT
        ws.cell(row=row_idx, column=1, value=idx + 1)
        
        # Điền các trường còn lại
        for col_idx in range(2, len(COLUMNS) + 1):
            col_name = COLUMNS[col_idx - 1]
            val = item.get(col_name, "")
            ws.cell(row=row_idx, column=col_idx, value=val)
            
        # Áp dụng định dạng cho từng ô trong hàng
        is_zebra = (idx % 2 == 1)
        for col_idx in range(1, len(COLUMNS) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            
            # Phân loại Alignment & Bold cho cột
            align_left = col_idx not in [1, 2, 3, 4, 7, 12] # STT, Ngày, Kênh, Pillar, Góc, Trạng thái
            bold = col_idx in [2, 3, 4] # Ngày đăng, Nền tảng, Pillar
            
            excel_styles.apply_body_style(cell, is_zebra=is_zebra, align_left=align_left, bold=bold)
                
    # 5. Thiết lập Data Validation (Dropdown) cho cột Trạng thái Duyệt (Cột 12 = L)
    dv = DataValidation(
        type="list", 
        formula1='"Chờ duyệt,Duyệt Option A,Duyệt Option B,Duyệt Option C,Không duyệt"', 
        allow_blank=True
    )
    dv.prompt = 'Chọn trạng thái phê duyệt cho bài viết này'
    dv.promptTitle = 'Trạng thái Duyệt'
    
    # Áp dụng cho các ô chứa dữ liệu
    end_row = start_row + len(data) - 1
    if end_row >= start_row:
        validation_range = f"L{start_row}:L{end_row}"
        ws.add_data_validation(dv)
        dv.add(validation_range)
    
    # 6. Thiết lập độ rộng cột tối ưu
    column_widths = {
        "A": 8,   # STT
        "B": 18,  # Ngày đăng
        "C": 12,  # Nền tảng
        "D": 10,  # Pillar
        "E": 20,  # Persona
        "F": 16,  # Định dạng
        "G": 18,  # Góc tiếp cận
        "H": 40,  # Option A
        "I": 40,  # Option B
        "J": 40,  # Option C
        "K": 25,  # SSoT
        "L": 20   # Trạng thái Duyệt
    }
    
    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width
        
    wb.save(output_excel)
    print(f"Đã lưu lịch thành công tại: {output_excel}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Tạo lịch tuần Excel thương hiệu cá nhân PBAN.")
    parser.options = parser.add_argument("--input", default=DEFAULT_INPUT_JSON, help="Đường dẫn đến file JSON chứa dữ liệu lịch biên tập.")
    parser.options = parser.add_argument("--output", default=DEFAULT_OUTPUT_EXCEL, help="Đường dẫn lưu file Excel kết quả.")
    parser.options = parser.add_argument("--week", default="TUẦN 24/2026", help="Tiêu đề tuần đăng nội dung.")
    
    args = parser.parse_args()
    create_calendar_excel(args.input, args.output, args.week)
