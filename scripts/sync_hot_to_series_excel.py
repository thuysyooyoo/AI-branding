# -*- coding: utf-8 -*-
"""
Script đồng bộ và hoàn thiện sổ cái Master Series YouTube trên OneDrive:
`C:\\Users\\thuyn\\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\\BRANDING\\Ke_Hoach_Phat_Song_YouTube_Series.xlsx`

Tuân thủ nghiêm ngặt:
1. RULE 10 (Permanent Zero Backup Rule): atomic_save qua file .tmp.xlsx rồi os.replace.
2. RULE 22: Cập nhật 100% đúng cấu trúc 15 cột của 4 Sheet chuyên mục (CM1, CM2, CM3, CM4).
3. Sửa lỗi lệch cột ở dòng CM2-CD19 (nếu có).
4. Thêm 5 chủ đề nóng W45 (HOT37 -> HOT41) vào các sheet tương ứng với trạng thái `🟢 ĐÃ HOÀN THÀNH KỊCH BẢN` và công thức `=HYPERLINK(...)` 1-click.
5. Cập nhật bảng tổng quan tại Sheet `TONG_QUAN_LICH_PHAT_SONG`.
"""

import os
import shutil
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

ONEDRIVE_DIR = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING"
ONEDRIVE_YT_DIR = os.path.join(ONEDRIVE_DIR, "youtube")
EXCEL_PATH = os.path.join(ONEDRIVE_DIR, "Ke_Hoach_Phat_Song_YouTube_Series.xlsx")

# Định nghĩa các style chuẩn của Master File
font_title = Font(name="Arial", size=10, bold=False, color="000000")
font_bold_center = Font(name="Arial", size=10, bold=True, color="000000")
font_status_done = Font(name="Arial", size=10, bold=True, color="000E6251")
fill_status_done = PatternFill(start_color="D4EFDF", end_color="D4EFDF", fill_type="solid")
font_link = Font(name="Arial", size=10, bold=False, color="0000FF", underline="single")
font_stale_retained = Font(name="Arial", size=10, bold=True, color="145A32")

align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

thin_border = Border(
    left=Side(style='thin', color='D3D3D3'),
    right=Side(style='thin', color='D3D3D3'),
    top=Side(style='thin', color='D3D3D3'),
    bottom=Side(style='thin', color='D3D3D3')
)

def atomic_save_workbook(wb, filepath):
    """Ghi file an toàn tuyệt đối không tạo backup (Rule 10)."""
    tmp_path = filepath + ".tmp.xlsx"
    if os.path.exists(tmp_path):
        try:
            os.remove(tmp_path)
        except Exception:
            pass
    wb.save(tmp_path)
    wb.close()
    os.replace(tmp_path, filepath)
    print(f"✅ Đã ghi đè an toàn vào: {filepath}")

def fix_cm2_cd19_if_needed(ws):
    """Sửa lỗi lệch cột của dòng CM2-CD19 (row 17) nếu trước đó bị copy nhầm từ calendar."""
    target_row = None
    for r in range(5, ws.max_row + 1):
        if ws.cell(r, 2).value == "CM2-CD19":
            target_row = r
            break
            
    if target_row:
        print(f"🔧 Đang chuẩn hóa cấu trúc 15 cột cho CM2-CD19 tại dòng {target_row}...")
        ws.cell(target_row, 1, 13).alignment = align_center
        ws.cell(target_row, 1).font = font_bold_center
        ws.cell(target_row, 2, "CM2-CD19").alignment = align_center
        ws.cell(target_row, 2).font = font_bold_center
        ws.cell(target_row, 3, "Công Văn 22560 Về Nhập Khẩu Ủy Thác Từ Trung Quốc: 3 Nguyên Tắc Cốt Tử Về Người Khai Hải Quan & Tính Hợp Lệ Của C/O Mẫu E Để Hưởng Thuế ACFTA!").alignment = align_left
        ws.cell(target_row, 3).font = font_title
        ws.cell(target_row, 4, "Video Đơn lẻ (7-9p)").alignment = align_center
        ws.cell(target_row, 4).font = font_title
        ws.cell(target_row, 5, "7 - 9 phút").alignment = align_center
        ws.cell(target_row, 5).font = font_title
        ws.cell(target_row, 6, "Công văn 22560/CHQ-GSQL; Nghị định 08/2015/NĐ-CP; Thông tư 38/2015/TT-BTC; Hiệp định ACFTA").alignment = align_left
        ws.cell(target_row, 6).font = font_title
        ws.cell(target_row, 7, "(0:00 - 0:45) Nhập khẩu ủy thác qua bên thứ ba bị bác C/O Form E và truy thu thuế ACFTA hàng trăm triệu đồng... Thúy chia sẻ 3 nguyên tắc sống còn khi làm thủ tục ủy thác để không bao giờ bị bác C/O.").alignment = align_left
        ws.cell(target_row, 7).font = font_title
        ws.cell(target_row, 8, "1. Bản chất Công văn 22560 về người đứng tên trên C/O và tờ khai hải quan.\n2. Điều kiện hợp lệ để hưởng thuế suất ưu đãi đặc biệt ACFTA khi ủy thác nhập khẩu.\n3. Các lỗi sai chí mạng doanh nghiệp thường mắc phải khi làm chứng từ qua bên thứ ba.\n4. Bộ hồ sơ ủy thác mẫu và quy trình kiểm soát rủi ro.").alignment = align_left
        ws.cell(target_row, 8).font = font_title
        ws.cell(target_row, 9, "Bàn gỗ mộc, ấm trà nóng, bộ hồ sơ ủy thác nhập khẩu, mẫu C/O Form E thực tế, screencast tờ khai VNACCS.").alignment = align_left
        ws.cell(target_row, 9).font = font_title
        ws.cell(target_row, 10, "Không áp dụng (Tập đơn lẻ)").alignment = align_left
        ws.cell(target_row, 10).font = font_title
        ws.cell(target_row, 11, "Làm xuất nhập khẩu ủy thác, Thúy luôn tự dặn dò đội ngũ: Sự cẩn trọng trong từng điều khoản hợp đồng và tên người khai trên chứng từ chính là chiếc lá chắn bảo vệ an toàn thuế cho khách hàng.").alignment = align_left
        ws.cell(target_row, 11).font = font_title
        ws.cell(target_row, 12, "🟢 ĐÃ HOÀN THÀNH KỊCH BẢN").alignment = align_center
        ws.cell(target_row, 12).font = font_status_done
        ws.cell(target_row, 12).fill = fill_status_done
        ws.cell(target_row, 13, f'=HYPERLINK("{os.path.join(ONEDRIVE_YT_DIR, "CM2-CD19_Cong_van_22560_Nhap_khau_uy_thac_va_CO_Form_E.md")}", "👉 Xem Kịch Bản Chi Tiết")').alignment = align_left
        ws.cell(target_row, 13).font = font_link
        ws.cell(target_row, 14, "W46/2026").alignment = align_center
        ws.cell(target_row, 14).font = font_bold_center
        ws.cell(target_row, 15, "🛡️ Bảo lưu vĩnh viễn (Đã hoàn thành)").alignment = align_center
        ws.cell(target_row, 15).font = font_stale_retained
        
        for c in range(1, 16):
            ws.cell(target_row, c).border = thin_border

def append_hot_topic(ws, item_data):
    """Bổ sung một chủ đề HOT vào sheet tương ứng."""
    code = item_data["code"]
    
    # Kiểm tra xem code đã tồn tại trong sheet chưa
    for r in range(5, ws.max_row + 1):
        if ws.cell(r, 2).value == code:
            print(f"⚠️ Mã {code} đã tồn tại tại dòng {r}. Cập nhật dữ liệu...")
            row_idx = r
            break
    else:
        row_idx = ws.max_row + 1
        print(f"➕ Thêm mới mã {code} vào dòng {row_idx}...")
        
    stt = row_idx - 4  # Dòng 5 là STT 1
    
    ws.cell(row_idx, 1, stt).alignment = align_center
    ws.cell(row_idx, 1).font = font_bold_center
    
    ws.cell(row_idx, 2, code).alignment = align_center
    ws.cell(row_idx, 2).font = font_bold_center
    
    ws.cell(row_idx, 3, item_data["title"]).alignment = align_left
    ws.cell(row_idx, 3).font = font_title
    
    ws.cell(row_idx, 4, item_data.get("format", "Video Đơn lẻ (8-10p)")).alignment = align_center
    ws.cell(row_idx, 4).font = font_title
    
    ws.cell(row_idx, 5, item_data.get("duration", "8 - 10 phút")).alignment = align_center
    ws.cell(row_idx, 5).font = font_title
    
    ws.cell(row_idx, 6, item_data["sources"]).alignment = align_left
    ws.cell(row_idx, 6).font = font_title
    
    ws.cell(row_idx, 7, item_data["hook"]).alignment = align_left
    ws.cell(row_idx, 7).font = font_title
    
    ws.cell(row_idx, 8, item_data["core"]).alignment = align_left
    ws.cell(row_idx, 8).font = font_title
    
    ws.cell(row_idx, 9, item_data["visual"]).alignment = align_left
    ws.cell(row_idx, 9).font = font_title
    
    ws.cell(row_idx, 10, item_data.get("cliffhanger", "Không áp dụng (Tập đơn lẻ)")).alignment = align_left
    ws.cell(row_idx, 10).font = font_title
    
    ws.cell(row_idx, 11, item_data["outro"]).alignment = align_left
    ws.cell(row_idx, 11).font = font_title
    
    ws.cell(row_idx, 12, "🟢 ĐÃ HOÀN THÀNH KỊCH BẢN").alignment = align_center
    ws.cell(row_idx, 12).font = font_status_done
    ws.cell(row_idx, 12).fill = fill_status_done
    
    file_path = os.path.join(ONEDRIVE_YT_DIR, item_data["filename"])
    formula = f'=HYPERLINK("{file_path}", "👉 Xem Kịch Bản Chi Tiết")'
    ws.cell(row_idx, 13, formula).alignment = align_left
    ws.cell(row_idx, 13).font = font_link
    
    ws.cell(row_idx, 14, item_data.get("week", "W45/2026")).alignment = align_center
    ws.cell(row_idx, 14).font = font_bold_center
    
    ws.cell(row_idx, 15, "🛡️ Bảo lưu vĩnh viễn (Đã hoàn thành)").alignment = align_center
    ws.cell(row_idx, 15).font = font_stale_retained
    
    for c in range(1, 16):
        ws.cell(row_idx, c).border = thin_border
        
    ws.row_dimensions[row_idx].height = 42

def main():
    if not os.path.exists(EXCEL_PATH):
        print(f"❌ Không tìm thấy file Master Excel tại: {EXCEL_PATH}")
        return

    print(f"📂 Đang tải workbook: {EXCEL_PATH}")
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=False)

    # 1. Sửa lỗi row 17 của CM2 nếu cần
    if "CM2_DICH_LUAT_TIENG_NGUOI" in wb.sheetnames:
        fix_cm2_cd19_if_needed(wb["CM2_DICH_LUAT_TIENG_NGUOI"])

    # 2. Dữ liệu 5 chủ đề nóng
    DATA_HOT = [
        # CM1
        {
            "sheet": "CM1_VAN_BAN_NONG",
            "code": "CM1-HOT37",
            "title": "Kỳ Án Kiểm Tra Chuyên Ngành: Một Lô Hàng Cõng 3 Tầng Giấy Phép & Đột Phá Thông Tư 104/2026 Thống Nhất Cơ Quan Hải Quan Làm Đầu Mối Duy Nhất",
            "format": "Video Đơn lẻ (8-10p)",
            "duration": "8 - 10 phút",
            "sources": "Thông tư số 104/2026/TT-BTC; Nghị định số 08/2015/NĐ-CP; Quyết định số 38/QĐ-TTg của Thủ tướng Chính phủ; Cổng thông tin Tổng cục Hải quan (customs.gov.vn).",
            "hook": "(0:00 - 0:45) Lô hàng máy đo khoảng cách laser gắn camera nhập khẩu về cảng bị dừng chờ kiểm tra chuyên ngành cùng lúc 3 bộ quản lý: Bộ Thông tin & Truyền thông, Bộ Khoa học & Công nghệ, và Bộ Công Thương! Tiền lưu cont lưu bãi vọt lên 40 triệu đồng. Tại sao một mặt hàng lại phải cõng 3 tầng giấy phép chồng chéo và Thông tư 104/2026 đã giải quyết căn bệnh nan giải này như thế nào?",
            "core": "1. Bóc tách nguyên nhân một mặt hàng tích hợp công nghệ cao bị nhiều Bộ quản lý chồng chéo (tính năng đo laser, bộ thu phát wifi/bluetooth, nhãn hàng hóa).\n2. Đột phá từ Thông tư 104/2026/TT-BTC: Thống nhất cơ quan Hải quan làm đầu mối duy nhất tại cửa khẩu tiếp nhận hồ sơ và kiểm tra theo cơ chế Một cửa quốc gia.\n3. Nguyên tắc quản lý rủi ro & thừa nhận lẫn nhau: Giảm tỷ lệ kiểm tra thực tế đối với doanh nghiệp tuân thủ tốt pháp luật.\n4. 4 Bước doanh nghiệp cần chủ động chuẩn bị hồ sơ kỹ thuật, catalog và giấy chứng nhận hợp quy để thông quan ngay trong 24 giờ.",
            "visual": "Bàn gỗ mộc, ấm trà nóng, ánh sáng vàng ấm 3500K. Đạo cụ cuốn Luật Hải quan và Thông tư 104/2026/TT-BTC. Screencast sơ đồ luồng kiểm tra chuyên ngành và giao diện Cổng Một cửa Quốc gia vnsw.gov.vn.",
            "cliffhanger": "Không áp dụng (Tập đơn lẻ)",
            "outro": "Với bản thân Thúy, mười lăm năm lăn lộn với từng tờ khai, từng biên bản lấy mẫu hiện trường, mình luôn tự dặn lòng một điều rất giản dị: Cơ chế chính sách có thể thông thoáng hơn, cổng một cửa có thể hiện đại hơn, nhưng sự chủ động nghiên cứu và tính kỷ luật tuân thủ của người làm nghề mới là chiếc chìa khóa vạn năng để bảo vệ dòng tiền cho doanh nghiệp.",
            "filename": "CM1-HOT37_Thong_tu_104_Kiem_tra_chuyen_nganh.md",
            "week": "W45/2026"
        },
        {
            "sheet": "CM1_VAN_BAN_NONG",
            "code": "CM1-HOT38",
            "title": "Một Cửa Quốc Gia Giai Đoạn 3 & Nghị Định 336/2026: 4 Bước Đồng Bộ Chữ Ký Số & Kiểm Tra Tính Hợp Lệ C/O Điện Tử Trên Cổng ASEAN Tránh Bị Truy Thu Thuế MFN",
            "format": "Video Đơn lẻ (8-10p)",
            "duration": "8 - 10 phút",
            "sources": "Nghị định số 336/2026/NĐ-CP của Chính phủ; Thông tư số 33/2023/TT-BTC về xác định xuất xứ hàng hóa; Luật Giao dịch điện tử số 20/2023/QH15; Cổng thông tin Một cửa quốc gia (vnsw.gov.vn).",
            "hook": "(0:00 - 0:45) Hàng nhập khẩu về cảng, đối tác xuất khẩu đã cấp C/O điện tử Form D, Form E đầy đủ nhưng khi truyền tờ khai hải quan VNACCS lại bị báo lỗi không tìm thấy chứng từ xuất xứ trên hệ thống Một cửa ASEAN! Doanh nghiệp đứng trước nguy cơ phải nộp thuế theo mức MFN chênh lệch hàng trăm triệu đồng hoặc nợ thuế để chờ xác minh. Đâu là nguyên nhân đằng sau và Nghị định 336/2026 giai đoạn 3 yêu cầu những gì?",
            "core": "1. Điểm mới cốt lõi của Nghị định 336/2026/NĐ-CP triển khai Một cửa quốc gia Giai đoạn 3: Chấm dứt hoàn toàn việc nộp file scan C/O PDF thủ công, bắt buộc đồng bộ mã định danh điện tử doanh nghiệp tập trung.\n2. 3 Nguyên nhân phổ biến khiến dữ liệu C/O điện tử không khớp giữa cổng ASEAN Single Window và hệ thống VNACCS (lệch mã số thuế người nhập khẩu, sai định dạng số tham chiếu Reference No, chưa đồng bộ chữ ký số công cộng).\n3. Quy trình 4 bước kiểm tra tính hợp lệ và trạng thái trao đổi C/O điện tử trực tiếp trên Cổng Một cửa Quốc gia vnsw.gov.vn.\n4. Cách xử lý tình huống khẩn cấp khi tàu đã cập cảng nhưng C/O điện tử chưa sang kịp: Thủ tục xin bảo lãnh thuế hoặc cam kết nộp bổ sung chứng từ theo Thông tư 33/2023/TT-BTC.",
            "visual": "Bàn gỗ mộc, ấm trà nóng, laptop mở sẵn Cổng Một cửa Quốc gia vnsw.gov.vn. Screencast trực quan 4 bước tra cứu trạng thái C/O điện tử trên Cổng ASEAN Single Window, highlight các trường dữ liệu quan trọng như Số tham chiếu Reference No, Mã số thuế nhà xuất khẩu, Mã phân loại HS 6 số.",
            "cliffhanger": "Không áp dụng (Tập đơn lẻ)",
            "outro": "Với bản thân Thúy, sau bao nhiêu lần cùng anh em hiện trường hồi hộp chờ từng dòng dữ liệu nhảy trên màn hình máy tính, mình luôn tự dặn lòng một điều: Trong kỷ nguyên số hóa hải quan, dữ liệu chính xác đến từng dấu chấm dấu phẩy chính là sinh mệnh pháp lý của doanh nghiệp. Nhanh một phút kiểm tra trước khi truyền tờ khai sẽ cứu được cả trăm triệu đồng tiền phạt và những tháng ngày khiếu nại mệt mỏi.",
            "filename": "CM1-HOT38_Nghi_dinh_336_Mot_cua_quoc_gia_Giai_doan_3.md",
            "week": "W45/2026"
        },
        {
            "sheet": "CM1_VAN_BAN_NONG",
            "code": "CM1-HOT40",
            "title": "Cửa Khẩu Thông Minh Lạng Sơn Vận Hành Xe Tự Hành AGV 24/7: Cẩm Nang 4 Bước Khai Báo Trước (Pre-clearance) Theo Thông Tư 128/2026 Để Thông Quan Hàng Xuyên Đêm",
            "format": "Video Đơn lẻ (8-10p)",
            "duration": "8 - 10 phút",
            "sources": "Thông tư số 128/2026/TT-BTC của Bộ Tài chính; Đề án thí điểm xây dựng Cửa khẩu thông minh tại cặp cửa khẩu Hữu Nghị - Hữu Nghị Quan (Lạng Sơn); Cổng thông tin Cục Hải quan Lạng Sơn; Cổng Một cửa Quốc gia.",
            "hook": "(0:00 - 0:45) Nếu các bạn đã từng làm hàng xuất khẩu nông sản như sầu riêng, thanh long, mít sang thị trường Trung Quốc qua cửa khẩu Lạng Sơn, chắc chắn không thể nào quên được cảnh tượng hàng nghìn xe container nối đuôi nhau hàng chục cây số, tài xế ăn chực nằm chờ cả tuần lễ, nông sản thối hỏng bốc mùi khiến chủ hàng thiệt hại hàng trăm triệu đồng mỗi chuyến xe! Nhưng một bước ngoặt lịch sử đang diễn ra: Cửa khẩu thông minh với xe tự hành AGV vận hành xuyên đêm 24/7 không người lái.",
            "core": "1. Mô hình vận hành thực tế của Cửa khẩu thông minh Lạng Sơn: Xe tự hành AGV vận chuyển container trên đường chuyên dụng riêng biệt, sử dụng công nghệ định vị vệ tinh và camera nhận diện AI tự động quét mã QR biển số xe và container.\n2. Yêu cầu pháp lý bắt buộc theo Thông tư 128/2026/TT-BTC: Thủ tục khai báo hải quan trước (Pre-clearance), truyền dữ liệu số hóa danh sách hàng hóa và định vị xe hàng chậm nhất 4 tiếng trước khi phương tiện đến biên giới.\n3. 4 Bước quy trình thực chiến cho chủ hàng và tài xế xe container để khớp lệnh vào luồng xe AGV tự hành xuyên đêm không cần dừng chờ làm thủ tục giấy.\n4. Cảnh báo các lỗi kỹ thuật thường gặp khiến xe cont bị văng khỏi luồng tự động (sai biển số moóc kéo, dung lượng pin thiết bị định vị GPS bị gián đoạn, lệch trọng lượng thực tế so với phiếu cân điện tử).",
            "visual": "Bàn gỗ mộc, ấm trà nóng, iPad trình chiếu video tư liệu hoạt động của dàn xe tự hành AGV tại cửa khẩu thông minh Hữu Nghị - Tân Thanh. Bản đồ đồ họa 3D mô phỏng luồng di chuyển khép kín của xe cont qua trạm cân tự động, cổng quét RFID và đường ray chuyên dụng.",
            "cliffhanger": "Không áp dụng (Tập đơn lẻ)",
            "outro": "Với bản thân Thúy, sau bao nhiêu năm gắn bó với từng chuyến xe cont lên biên giới, mình luôn tự dặn lòng một điều: Công nghệ tự hành AGV hay camera AI dù có hiện đại đến đâu thì cũng chỉ là công cụ. Yếu tố quyết định sự hanh thông của một chuyến hàng nông sản xuất khẩu vẫn nằm ở sự chuẩn bị kỹ lưỡng về chất lượng mã số vùng trồng, quy cách đóng gói và sự chuẩn xác tuyệt đối trong hồ sơ khai báo trước của người làm logistics.",
            "filename": "CM1-HOT40_Cua_khau_thong_minh_AGV_Thong_tu_128.md",
            "week": "W45/2026"
        },
        # CM2
        {
            "sheet": "CM2_DICH_LUAT_TIENG_NGUOI",
            "code": "CM2-HOT41",
            "title": "Tờ Khai Hải Quan Bị Treo Trên ECUS5 Nhưng Hiện Trường Không Biết: 4 Bước Tra Cứu Log Truyền Nhận & Đối Soát Trực Tiếp Với Cổng Thông Tin Hải Quan Tránh Kẹt Hàng Ở Cảng",
            "format": "Video Đơn lẻ (8-10p)",
            "duration": "8 - 10 phút",
            "sources": "Luật Hải quan số 54/2014/QH13; Thông tư số 38/2015/TT-BTC và Thông tư số 39/2018/TT-BTC; Tài liệu hướng dẫn kỹ thuật hệ thống VNACCS/VCIS của Tổng cục Hải quan; Cổng thông tin điện tử Hải quan (customs.gov.vn).",
            "hook": "(0:00 - 0:45) Có một tình huống mà Thúy tin chắc rằng bất kỳ anh em nào làm hiện trường hay các bạn làm chứng từ xuất nhập khẩu cũng từng trải qua ít nhất một lần: Phần mềm khai báo ECUS5 trên máy tính báo đã gửi tờ khai thành công và hiện mã phản hồi màu vàng, nhưng khi anh em ra cảng trình tờ khai để lấy hàng thì công chức hải quan lại bảo: 'Trên hệ thống kiểm soát của chi cục chưa có số tờ khai này!' Tiền phạt phát sinh, xe cont chờ ở cổng, điện thoại reo liên tục. Bản chất của hiện tượng tờ khai bị 'treo' là gì?",
            "core": "1. Giải mã nguyên nhân kỹ thuật khiến tờ khai bị treo: Sự bất đối xứng dữ liệu giữa phần mềm đại lý trung gian (ECUS5/VNACCS) và cổng xử lý trung tâm Tổng cục Hải quan (lỗi nghẽn băng thông giờ cao điểm, trễ bản tin phản hồi thông điệp EDI, đứt gãy phiên truyền nhận chữ ký số).\n2. Phân biệt rõ giữa 2 trạng thái: 'Đã nhận phản hồi từ gateway' và 'Đã cập nhật vào cơ sở dữ liệu xử lý của chi cục hải quan nơi mở tờ khai'.\n3. 4 Bước tra cứu log truyền nhận chi tiết (EDI Log Viewer) ngay trong phần mềm ECUS5 để tìm mã lỗi chính xác (Error Code) và đối soát trực tiếp trên Cổng tra cứu tờ khai trực tuyến của Tổng cục Hải quan customs.gov.vn.\n4. Kỹ thuật liên hệ và lập phiếu yêu cầu hỗ trợ kỹ thuật gửi Đội CNTT Cục Hải quan hoặc hỗ trợ kỹ thuật Thái Sơn để đẩy dữ liệu cưỡng bức, giúp thông quan giải phóng hàng chỉ trong vòng 30 phút mà không cần hủy tờ khai làm lại.",
            "visual": "Bàn gỗ mộc, ấm trà nóng, màn hình máy tính mở trực tiếp phần mềm ECUS5 và cửa sổ Log truyền nhận thông điệp EDI. B-roll zoom cận cảnh vào các dòng mã phản hồi kỹ thuật, so sánh giao diện tờ khai trên phần mềm vs giao diện tra cứu trạng thái trên cổng customs.gov.vn.",
            "cliffhanger": "Không áp dụng (Tập đơn lẻ)",
            "outro": "Với bản thân Thúy, mười lăm năm làm nghề từ những ngày đầu còn cầm hồ sơ giấy chạy khắp các chi cục hải quan, mình luôn tự dặn lòng một câu: Làm nghề xuất nhập khẩu không sợ gặp sự cố kỹ thuật, chỉ sợ người làm hiện trường hoang mang rồi đổ lỗi lòng vòng cho phần mềm hay cơ quan chức năng. Khi mình hiểu sâu bản chất kỹ thuật của hệ thống và nắm chắc từng bước đối soát dữ liệu, thì mọi khúc mắc đều có thể bình tĩnh tháo gỡ một cách văn minh và nhanh chóng.",
            "filename": "CM2-HOT41_To_khai_treo_ECUS5_Doi_soat_thue.md",
            "week": "W45/2026"
        },
        # CM4
        {
            "sheet": "CM4_TAM_TINH_THUC_CHIEN",
            "code": "CM4-HOT39",
            "title": "Kỳ Án Hủy Chuyến Chủ Động (Blank Sailings) Mùa Cuối Năm: Khi Hãng Tàu Cố Tình Bỏ Trống Chuyến Tàu Ép Cước Lên 4.400 USD & 3 Vũ Khí Đàm Phán Của Chủ Hàng Việt Nam",
            "format": "Video Đơn lẻ (8-10p)",
            "duration": "8 - 10 phút",
            "sources": "Báo cáo Chỉ số Vận tải Container Thế giới (Drewry WCI); Dữ liệu nghiên cứu của Sea-Intelligence; Báo cáo thị trường logistics cuối năm của VLA; Hợp đồng vận chuyển mẫu BIMCO/FIATA.",
            "hook": "(0:00 - 0:45) Các bạn làm xuất khẩu hàng dệt may, da giày, đồ gỗ hay thủy sản sang thị trường Mỹ và Châu Âu, có bao giờ rơi vào tình cảnh cay đắng này chưa: Đã chốt booking từ cách đây hai tuần, giá cước 2.800 USD một container 40 feet, hàng đã đóng xong xuôi và đưa về bãi chờ bốc lên tàu. Đúng một ngày trước giờ cắt máng Closing time, hãng tàu gửi email thông báo hủy chuyến tàu (Blank Sailing) vì lý do 'kỹ thuật', rồi gạ chuyển sang chuyến tàu tuần sau nhưng giá cước bị đội lên 4.400 USD! Đằng sau chiêu bài này là gì?",
            "core": "1. Bản chất của chiến thuật 'Hủy chuyến chủ động' (Blank Sailings): Các liên minh hãng tàu biển lớn trên thế giới bắt tay nhau cắt giảm từ 15% đến 25% công suất tàu trên các tuyến hải trình trọng điểm Á - Mỹ, Á - Âu nhằm tạo ra sự khan hiếm công giả tạo để đẩy giá cước mùa cao điểm cuối năm.\n2. Thiệt hại dây chuyền đối với chủ hàng Việt Nam: Chậm hợp đồng giao hàng mùa Giáng sinh, bị đối tác nước ngoài phạt vi phạm hợp đồng, ứ đọng dòng tiền hàng triệu USD và tốn kém chi phí lưu kho bãi (Demurrage & Detention).\n3. 3 Vũ khí đàm phán sắc bén của chủ hàng xuất khẩu Việt Nam:\n   - Chiến lược đàm phán hợp đồng khối lượng lớn (Volume Contract) kèm điều khoản phạt hủy chỗ đối ứng (No-Roll Guarantee / Space Commitment Clause).\n   - Đa dạng hóa liên minh tàu (Multi-carrier Strategy): Không dồn 100% rổ hàng vào một liên minh để tránh bị động tuyệt đối khi một liên minh cắt chuyến.\n   - Sử dụng các công ty giao nhận vận tải uy tín (Tier-1 Forwarders) có cam kết phân bổ chỗ (Slot Allocation) cố định để được bảo vệ khi thị trường biến động.\n4. Bài học thực chiến quản trị chuỗi cung ứng mùa cao điểm: Nguyên tắc chốt booking trước tối thiểu 3 đến 4 tuần và kỹ thuật theo dõi dữ liệu lịch trình tàu thời gian thực trên các nền tảng phân tích hàng hải.",
            "visual": "Bàn gỗ mộc, ấm trà nóng, iPad hiển thị biểu đồ biến động giá cước của chỉ số Drewry WCI và thông báo Blank Sailings của các hãng tàu lớn. B-roll tư liệu cảnh cảng biển nhộn nhịp tàu bè, hàng nghìn thùng container xếp tầng, kèm infographic trực quan 3 chiến lược đàm phán.",
            "cliffhanger": "Không áp dụng (Tập đơn lẻ)",
            "outro": "Với bản thân Thúy, mười lăm năm làm nghề logistics, mình luôn tự dặn lòng một bài học rất sâu sắc: Trong kinh doanh vận tải quốc tế, con số rẻ nhất chưa bao giờ là con số tốt nhất. Một mức giá cước siêu rẻ mà không đi kèm với sự đảm bảo chắc chắn về chỗ trên tàu thì chỉ là một lời hứa ảo. Doanh nghiệp xuất khẩu khôn ngoan là doanh nghiệp biết xây dựng mối quan hệ đối tác chiến lược bền vững, chấp nhận một mức chi phí hợp lý để đổi lấy sự an toàn tuyệt đối cho dòng chảy hàng hóa và uy tín thương hiệu trên trường quốc tế.",
            "filename": "CM4-HOT39_Drewry_WCI_Blank_Sailings_Dam_phan_cuoc.md",
            "week": "W45/2026"
        }
    ]

    for item in DATA_HOT:
        sname = item["sheet"]
        if sname in wb.sheetnames:
            append_hot_topic(wb[sname], item)

    # 3. Cập nhật Sheet TONG_QUAN_LICH_PHAT_SONG
    if "TONG_QUAN_LICH_PHAT_SONG" in wb.sheetnames:
        ws_tq = wb["TONG_QUAN_LICH_PHAT_SONG"]
        cm1_count = wb["CM1_VAN_BAN_NONG"].max_row - 4
        cm2_count = wb["CM2_DICH_LUAT_TIENG_NGUOI"].max_row - 4
        cm3_count = wb["CM3_LICH_SU_THUONG_MAI"].max_row - 4
        cm4_count = wb["CM4_TAM_TINH_THUC_CHIEN"].max_row - 4
        
        ws_tq.cell(6, 4, f"{cm1_count} Chủ đề (Đã bổ sung HOT37, HOT38, HOT40)")
        ws_tq.cell(7, 4, f"{cm2_count} Chủ đề (Đã bổ sung HOT41)")
        ws_tq.cell(8, 4, f"{cm3_count} Chủ đề (Đã lọc đào thải)")
        ws_tq.cell(9, 4, f"{cm4_count} Chủ đề (Đã bổ sung HOT39)")
        print(f"📊 Đã cập nhật số lượng chủ đề trên Sheet TONG_QUAN: CM1={cm1_count}, CM2={cm2_count}, CM3={cm3_count}, CM4={cm4_count}")

    # 4. Ghi file an toàn
    atomic_save_workbook(wb, EXCEL_PATH)
    print("🎉 Hoàn tất 100% việc hoàn thiện và đồng bộ Ke_Hoach_Phat_Song_YouTube_Series.xlsx trên OneDrive!")

if __name__ == "__main__":
    main()
