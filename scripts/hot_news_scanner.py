"""
Hot News Scanner — Quét Tin Nóng Tuần 33/2026 (11-17/08)
Thêm kết quả vào Sheet 'Nghiên Cứu Từ Khóa E2E' của file SEO_Keyword_Trends_Eureka.xlsx
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime
import shutil

import os

SRC = r"D:\AI AGENT THUY\AI branding\reports\SEO_Keyword_Trends_Eureka.xlsx"

wb = openpyxl.load_workbook(SRC)
ws = wb["Nghiên Cứu Từ Khóa E2E"]

# Unmerge all merged cells in the sheet first
merged = list(ws.merged_cells.ranges)
for m in merged:
    ws.unmerge_cells(str(m))

# Find last data row (skip the total row)
last_data_row = ws.max_row
for r in range(ws.max_row, 5, -1):
    cell = ws.cell(row=r, column=1)
    val = cell.value
    if val is not None and str(val).startswith("Tổng"):
        last_data_row = r - 1
        break

# Determine next STT
max_stt = 0
for r in range(6, last_data_row + 1):
    val = ws.cell(row=r, column=1).value
    if val is not None and isinstance(val, (int, float)):
        max_stt = max(max_stt, int(val))

print(f"Last data row: {last_data_row}, Max STT: {max_stt}")

# Hot News data
today = datetime.now().strftime("%d/%m/%Y")
hot_news = [
    {
        "stt": max_stt + 1,
        "segment": "5. Customs & Inspection",
        "seed_kw": "quyet dinh 31 2026 12 nhom hang cua khau nhap hai quan",
        "cluster_kw": "QD 31/2026/QD-TTg, 12 nhom hang lam thu tuc tai cua khau nhap, xe oto nhap khau cua khau, ruou bia thuoc la nhap khau hai quan, may dieu hoa nhap khau cua khau, hang hoa tac dong quoc phong an ninh",
        "trends_idx": "N/A (Mới)",
        "vol_est": "Chưa ước tính",
        "top1": "Tiêu đề: QĐ 31/2026/QĐ-TTg: 12 nhóm hàng NK bắt buộc làm thủ tục tại cửa khẩu nhập\nLink: https://baochinhphu.vn/quyet-dinh-31-2026-qd-ttg\nNguồn: Báo Chính phủ, Công Thương, VTV",
        "pain": "Từ 14/08/2026 (2 ngày nữa!), 12 nhóm hàng NK phải làm thủ tục TẠI cửa khẩu nhập. DN nhập rượu, bia, ô tô, máy lạnh nếu chưa biết sẽ bị kẹt hàng tại cảng. Trường hợp chung vận đơn với hàng thuộc danh mục → phải làm tất cả tại CK nhập.",
        "source": "Web Search (customs.gov.vn, baochinhphu.vn, congthuong.vn)",
    },
    {
        "stt": max_stt + 2,
        "segment": "4. Freight & Transportation",
        "seed_kw": "cuoc van tai bien container thang 8 2026 drewry wci",
        "cluster_kw": "drewry world container index 8/2026, cuoc tau bien tuyen chau a my, phu phi nhien lieu BAF EFS thang 8, GRI cuoc container, gia cuoc FEU 40ft, blank sailing hang tau",
        "trends_idx": "Tăng 1%",
        "vol_est": "~4,200",
        "top1": "Tiêu đề: Drewry WCI tuần 32/2026: 4.297 USD/FEU (+1%)\nLink: https://drewry.co.uk/supply-chain-advisors/supply-chain-expertise/world-container-index-assessed-by-drewry\nNguồn: Drewry, Freightos",
        "pain": "WCI phục hồi lên 4.297 USD/FEU. Tuyến Shanghai-NY: 7.893 USD (+4%), Shanghai-LA: 5.894 USD (+3%). Hãng tàu áp GRI đầu T8. Blank sailings duy trì. Biển Đỏ vẫn bất ổn → tàu đi vòng Hảo Vọng (+10-14 ngày). DN cần kiểm tra phụ phí EFS, PSS trước khi booking.",
        "source": "Web Search (drewry.co.uk, globalcontainernetwork.com)",
    },
    {
        "stt": max_stt + 3,
        "segment": "2. Quality & Compliance",
        "seed_kw": "xuat khau sau rieng cao diem thang 8 2026 ma so vung trong GACC",
        "cluster_kw": "xuat khau sau rieng trung quoc 2026, ma so vung trong co so dong goi, nghi quyet 36 2026, kiem nghiem cadimi vang O, GACC Trung Quoc phe duyet ma so, logistics sau rieng Tay Nguyen",
        "trends_idx": "84.5",
        "vol_est": "~8,500",
        "top1": "Tiêu đề: Cao điểm sầu riêng T8-T10: 1.3 triệu tấn cần tiêu thụ\nLink: Bộ Nông nghiệp & Môi trường\nNguồn: Báo NNVN, Cục Bảo vệ Thực vật",
        "pain": "T8-T10 là cao điểm thu hoạch (riêng Đắk Lắk 500K tấn, Lâm Đồng 300K tấn). 7 tháng đầu năm XK 493K tấn, đạt 950.7 triệu USD (+31%). 1,201 mã vùng trồng + 120 mã CSĐG đang hoạt động. Hơn 1,100 mã đang chờ GACC duyệt. NQ 36/2026/NQ-CP đơn giản hóa thủ tục. Rủi ro: ùn ứ cửa khẩu nếu không đủ năng lực kiểm nghiệm.",
        "source": "Web Search (Bộ NN&MT, GACC, customs.gov.vn)",
    },
    {
        "stt": max_stt + 4,
        "segment": "5. Customs & Inspection",
        "seed_kw": "khai bao ma nhan dien hang hoa luong dung hai quan 2026",
        "cluster_kw": "cong van 19916 CHQ-GSQL hang luong dung, nghi dinh 259/2025/ND-CP danh muc luong dung, khai bao to khai hai quan hang luong dung, ma van ban phap luat khac, kiem tra sau thong quan",
        "trends_idx": "N/A (Mới)",
        "vol_est": "~500",
        "top1": "Tiêu đề: Hướng dẫn khai báo mã nhận diện hàng hóa lưỡng dụng\nLink: Cục Hải quan - Công văn 19916/CHQ-GSQL\nNguồn: customs.gov.vn",
        "pain": "CV 19916/CHQ-GSQL hướng dẫn DN khai báo mã nhận diện hàng lưỡng dụng theo NĐ 259/2025/NĐ-CP. Phải bổ sung chỉ tiêu 'Mã VBPL khác' trên tờ khai. Không khai đúng → sửa đổi bổ sung tờ khai hoặc kiểm tra sau thông quan. Nhiều DN chưa biết có quy định này.",
        "source": "Web Search (customs.gov.vn, thuvienphapluat.vn)",
    },
    {
        "stt": max_stt + 5,
        "segment": "5. Customs & Inspection",
        "seed_kw": "luat hai quan sua doi 2026 hai quan so AI quan ly rui ro",
        "cluster_kw": "du an luat sua doi luat hai quan 2014, hai quan so AI thong minh, thuong mai dien tu xuyen bien gioi hai quan, VNACCS 2.0, ky hop quoc hoi thang 10 2026",
        "trends_idx": "30.6",
        "vol_est": "~2,700",
        "top1": "Tiêu đề: Dự thảo Luật sửa đổi Luật Hải quan: Hải quan số & AI\nLink: https://customs.gov.vn\nNguồn: customs.gov.vn, chinhphu.vn",
        "pain": "Dự án Luật sửa đổi Luật HQ 2014 dự kiến trình QH tháng 10/2026. Trọng tâm: hải quan số, hải quan thông minh (AI), khung pháp lý TMĐT xuyên biên giới. DN cần theo dõi để chuẩn bị cho thay đổi lớn. Tiềm năng content rất cao vì ảnh hưởng toàn ngành.",
        "source": "Web Search (customs.gov.vn, chinhphu.vn)",
    },
]

# Styles
thin_border = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
data_align = Alignment(vertical="top", wrap_text=True)
new_fill = PatternFill(start_color="FFF8DC", end_color="FFF8DC", fill_type="solid")
hot_fill = PatternFill(start_color="FFE4E1", end_color="FFE4E1", fill_type="solid")

# Add separator row
sep_row = last_data_row + 1
sep_cell = ws.cell(row=sep_row, column=1, value=f"--- QUÉT TIN NÓNG TUẦN 33/2026 ({today}) ---")
sep_cell.font = Font(bold=True, color="CC0000", size=11)
sep_cell.fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
ws.merge_cells(start_row=sep_row, start_column=1, end_row=sep_row, end_column=10)

# Write hot news data
start_row = sep_row + 1
for i, news in enumerate(hot_news):
    r = start_row + i
    values = [
        news["stt"],
        news["segment"],
        news["seed_kw"],
        news["cluster_kw"],
        news["trends_idx"],
        news["vol_est"],
        news["top1"],
        news["pain"],
        news["source"],
        today,
    ]

    for col_idx, val in enumerate(values, 1):
        cell = ws.cell(row=r, column=col_idx, value=val)
        cell.border = thin_border
        cell.alignment = data_align
        # Highlight hot items
        if news["stt"] == max_stt + 1:  # QD 31 — most urgent (effective in 2 days!)
            cell.fill = hot_fill
        else:
            cell.fill = new_fill

# Re-add total row
total_row = start_row + len(hot_news)
ws.cell(row=total_row, column=1, value=f"Tổng: {max_stt + len(hot_news)} từ khóa (gồm {len(hot_news)} tin nóng mới)")
ws.cell(row=total_row, column=1).font = Font(bold=True, italic=True)

wb.save(SRC)
print(f"\nSaved: {SRC}")
print(f"Added {len(hot_news)} hot news items to 'Nghiên Cứu Từ Khóa E2E' sheet")
print(f"Rows: {start_row} - {start_row + len(hot_news) - 1}")
print("\n=== HOT NEWS SUMMARY ===")
for n in hot_news:
    urgency = "🔴 KHẨN CẤP" if n["stt"] == max_stt + 1 else "🟡"
    print(f"  [{n['stt']}] {urgency} {n['seed_kw'][:60]}")
