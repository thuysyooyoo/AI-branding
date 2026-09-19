"""
Tái cấu trúc Weekly_Content_Calendar.xlsx theo Content Strategy 2.0
- Thay mã P1-AI/P2-XNK/P3-MGT/P4-LIFE → CS/TN/HT/FF
- Chuyển nền tảng chính: Facebook + YouTube (LinkedIn/TikTok = repurpose)
- Bỏ cột Persona Giả định (chỉ dùng nội bộ, không hiển thị)
- Thay Góc tiếp cận cũ → Kỹ thuật Kể chuyện mới
- Cập nhật headers và chiến lược kênh
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from copy import copy
import shutil
import os
from datetime import datetime

SRC = r"D:\AI AGENT THUY\AI branding\reports\plan1\Weekly_Content_Calendar.xlsx"

wb = openpyxl.load_workbook(SRC)

# ============================================================
# MAPPING TABLES
# ============================================================

# Old Pillar → New Category mapping
PILLAR_MAP = {
    "P1-AI": "HT",      # AI tutorials → Học Thực chiến (YouTube)
    "P2-XNK": "TN",     # XNK policy → Tin Ngành (Facebook debate)
    "P3-MGT": "CS",     # Management → Case Study (analysis)
    "P4-LIFE": "CS",    # Life stories → Case Study (personal stories)
}

CATEGORY_NAMES = {
    "CS": "CS — Case Study",
    "TN": "TN — Tin Ngành",
    "HT": "HT — Học Thực chiến",
    "FF": "FF — Funfact & Insight",
}

# Determine if a topic is better as FF (Funfact) based on keywords
FF_KEYWORDS = [
    "funfact", "bí mật", "bí ẩn", "bạn có biết", "sự thật", "lịch sử",
    "phát minh", "malcolm", "container lạnh", "reefer", "plimsoll",
    "twistlock", "nghịch lý", "chở không khí", "xe buýt xuyên biên giới",
    "lego", "sự thật bất ngờ", "1956"
]

# Old Angle → New Storytelling Technique mapping
ANGLE_MAP = {
    "Góc 1: Cảnh báo": "Problem-First",
    "Góc 2: Giải pháp": "Trao giá trị thực chiến",
    "Góc 3: AI Vận hành": "Trao giá trị thực chiến",
    "Góc 4: Đơn giản hóa": "Kích thích tư duy",
}

# Platform strategy: which categories go where
# Facebook: TN (debate), CS (discussion) — don't give full value
# YouTube: HT (tutorials), FF (funfact) — give full value
CATEGORY_PLATFORM = {
    "CS": "Facebook",    # Case studies → Facebook debate
    "TN": "Facebook",    # Industry news → Facebook debate
    "HT": "YouTube",     # Tutorials → YouTube education
    "FF": "YouTube",     # Funfact → YouTube insight
}


def classify_category(old_pillar, title, source_info, old_channel):
    """Classify content into new category based on old pillar, title keywords, and context."""
    title_lower = (title or "").lower()
    source_lower = (source_info or "").lower()

    # Check if it's a Funfact first (regardless of old pillar)
    if any(kw in title_lower for kw in FF_KEYWORDS):
        return "FF"
    if "funfact" in source_lower:
        return "FF"

    # P2-XNK: Check if it's more of a tutorial (HT) or news (TN)
    if old_pillar == "P2-XNK":
        tutorial_keywords = ["hướng dẫn", "mẹo", "5 phút", "5p", "checklist",
                           "quy trình", "từng bước", "cách", "học xnk", "bài tập",
                           "form e", "tra cứu", "kiểm tra"]
        if any(kw in title_lower for kw in tutorial_keywords):
            return "HT"
        return "TN"

    # P1-AI: Mostly tutorials about AI application → HT
    if old_pillar == "P1-AI":
        return "HT"

    # P3-MGT: Management & strategy → CS (case study / analysis)
    if old_pillar == "P3-MGT":
        return "CS"

    # P4-LIFE: Personal stories → CS (real stories from Thúy)
    if old_pillar == "P4-LIFE":
        return "CS"

    return PILLAR_MAP.get(old_pillar, "CS")


def determine_platform(category, old_channel):
    """Determine primary platform based on new category."""
    return CATEGORY_PLATFORM.get(category, "Facebook")


def determine_technique(old_angle, title, category):
    """Map old angle to new storytelling technique."""
    title_lower = (title or "").lower()

    # Category-based defaults
    if category == "FF":
        if any(kw in title_lower for kw in ["lịch sử", "1956", "malcolm", "phát minh", "bí ẩn"]):
            return "Liên hệ lịch sử"
        return "Dẫn chứng quốc tế"

    if category == "TN":
        return "Problem-First"

    if category == "CS":
        if any(kw in title_lower for kw in ["nếu", "giả sử", "tình huống"]):
            return "Tình huống giả định mở"
        if any(kw in title_lower for kw in ["bài học", "kinh nghiệm", "câu chuyện"]):
            return "Góc nhìn cá nhân Thúy"
        return "Tình huống giả định mở"

    if category == "HT":
        return "Trao giá trị thực chiến"

    # Fallback to angle mapping
    return ANGLE_MAP.get(old_angle, "Trao giá trị thực chiến")


def determine_format(category, platform):
    """Determine content format based on category and platform."""
    if platform == "YouTube":
        if category == "HT":
            return "Video Học XNK 5 phút"
        if category == "FF":
            return "Video Funfact & Insight (8-12p)"
        if category == "CS":
            return "Video Phân tích Chuyên đề (8-15p)"
        return "Video dài"

    if platform == "Facebook":
        if category == "TN":
            return "Bài tranh luận + Ảnh"
        if category == "CS":
            return "Bài Case Study + Ảnh"
        return "Bài viết + Ảnh"

    return "Bài viết"


# ============================================================
# PROCESS WEEKLY CALENDAR SHEET
# ============================================================
ws = wb["Weekly Calendar"]

# New headers
NEW_HEADERS = [
    "STT", "Ngày đăng", "Nền tảng", "Chuyên mục",
    "Kỹ thuật Kể chuyện", "Định dạng",
    "Option A (Chủ đề chính)", "Option B (Dự phòng 1)",
    "Option C (Dự phòng 2)", "Nguồn SSoT / Cốt truyện",
    "Trạng thái Duyệt"
]

# Update title row
ws.cell(row=1, column=1, value="LỊCH BIÊN TẬP TUẦN - HỆ THỐNG PBAN 2.0 (Content Strategy 2.0)")
ws.cell(row=2, column=1, value="Chiến lược: Facebook (Tranh luận, đặt vấn đề) + YouTube (Giáo dục, cung cấp giá trị). LinkedIn/TikTok = Repurpose.")

# Write new headers in row 4
header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
header_font = Font(color="FFFFFF", bold=True, size=11)
header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

for col_idx, header in enumerate(NEW_HEADERS, 1):
    cell = ws.cell(row=4, column=col_idx, value=header)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = header_align

# Clear old column 12 header if exists
if ws.max_column >= 12:
    ws.cell(row=4, column=12, value=None)

# Process data rows
new_rows = []
row_idx = 5
while row_idx <= ws.max_row:
    row_data = [ws.cell(row=row_idx, column=c).value for c in range(1, 13)]
    stt = row_data[0]
    if stt is None:
        row_idx += 1
        continue

    old_day = row_data[1]         # Ngày đăng
    old_channel = row_data[2]     # Nền tảng
    old_pillar = row_data[3]      # Pillar
    old_persona = row_data[4]     # Persona (will be removed)
    old_format = row_data[5]      # Định dạng
    old_angle = row_data[6]       # Góc tiếp cận
    opt_a = row_data[7]           # Option A
    opt_b = row_data[8]           # Option B
    opt_c = row_data[9]           # Option C
    source = row_data[10]         # Nguồn SSoT
    status = row_data[11]         # Trạng thái

    # Classify into new category
    new_category = classify_category(old_pillar, opt_a, source, old_channel)

    # Determine platform
    new_platform = determine_platform(new_category, old_channel)

    # If old channel was LinkedIn/TikTok, mark as repurpose
    if old_channel in ("LinkedIn", "TikTok"):
        repurpose_platform = "LinkedIn (Repurpose)" if old_channel == "LinkedIn" else "YT Shorts (Repurpose)"
        new_platform = repurpose_platform

    # Determine storytelling technique
    new_technique = determine_technique(old_angle, opt_a, new_category)

    # Determine format
    if "Repurpose" in new_platform:
        new_format = f"Repurpose từ {CATEGORY_PLATFORM.get(new_category, 'Facebook')}"
    else:
        new_format = determine_format(new_category, new_platform)

    new_rows.append({
        "stt": stt,
        "day": old_day,
        "platform": new_platform,
        "category": new_category,
        "technique": new_technique,
        "format": new_format,
        "opt_a": opt_a,
        "opt_b": opt_b,
        "opt_c": opt_c,
        "source": source,
        "status": status,
    })
    row_idx += 1

# Write new data
# Category colors
CAT_FILLS = {
    "CS": PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"),  # Light yellow
    "TN": PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid"),  # Light blue
    "HT": PatternFill(start_color="D5F5E3", end_color="D5F5E3", fill_type="solid"),  # Light green
    "FF": PatternFill(start_color="F5D6E0", end_color="F5D6E0", fill_type="solid"),  # Light pink
}

PLATFORM_FILLS = {
    "Facebook": PatternFill(start_color="E8F0FE", end_color="E8F0FE", fill_type="solid"),
    "YouTube": PatternFill(start_color="FCE4EC", end_color="FCE4EC", fill_type="solid"),
    "LinkedIn (Repurpose)": PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid"),
    "YT Shorts (Repurpose)": PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid"),
}

thin_border = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

data_align = Alignment(vertical="top", wrap_text=True)

for i, row in enumerate(new_rows):
    r = i + 5  # Start from row 5
    values = [
        row["stt"], row["day"], row["platform"],
        CATEGORY_NAMES.get(row["category"], row["category"]),
        row["technique"], row["format"],
        row["opt_a"], row["opt_b"], row["opt_c"],
        row["source"], row["status"]
    ]

    for col_idx, val in enumerate(values, 1):
        cell = ws.cell(row=r, column=col_idx, value=val)
        cell.border = thin_border
        cell.alignment = data_align

        # Apply category color to category column
        if col_idx == 4:
            fill = CAT_FILLS.get(row["category"])
            if fill:
                cell.fill = fill
            cell.font = Font(bold=True)

        # Apply platform color
        if col_idx == 3:
            fill = PLATFORM_FILLS.get(row["platform"])
            if fill:
                cell.fill = fill

    # Clear old column 12 data
    ws.cell(row=r, column=12, value=None)

# Clear any remaining old data rows
for r in range(len(new_rows) + 5, ws.max_row + 1):
    for c in range(1, 13):
        ws.cell(row=r, column=c, value=None)

# Set column widths
col_widths = {
    1: 6, 2: 22, 3: 22, 4: 26,
    5: 25, 6: 28,
    7: 55, 8: 55, 9: 55,
    10: 45, 11: 18
}
for col, width in col_widths.items():
    ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width


# Update Facebook sheet header
if "Facebook" in wb.sheetnames:
    ws_fb = wb["Facebook"]
    ws_fb.cell(row=1, column=1, value="BÀI VIẾT VÀ KIỂM DUYỆT CHI TIẾT - KÊNH FACEBOOK (Tranh luận & Đặt vấn đề)")
    ws_fb.cell(row=2, column=1, value="Chiến lược: Đặt vấn đề tranh luận, đưa góc nhìn cá nhân, KHÔNG cho hết giá trị (kéo traffic sang YouTube).")

# Update YouTube sheet header
if "YouTube" in wb.sheetnames:
    ws_yt = wb["YouTube"]
    ws_yt.cell(row=1, column=1, value="KỊCH BẢN VÀ KIỂM DUYỆT CHI TIẾT - KÊNH YOUTUBE (Giáo dục & Chuỗi Series)")
    ws_yt.cell(row=2, column=1, value="Chiến lược: Cung cấp kiến thức chuyên sâu, hướng dẫn thực chiến, Series nhiều tập + Cliffhanger Outro.")

# Remove legacy LinkedIn and TikTok sheets if present
for legacy_sheet in ["LinkedIn", "TikTok"]:
    if legacy_sheet in wb.sheetnames:
        del wb[legacy_sheet]


# ============================================================
# SAVE
# ============================================================
wb.save(SRC)
print(f"\nSaved: {SRC}")

# Print summary
from collections import Counter
cat_counts = Counter(r["category"] for r in new_rows)
plat_counts = Counter(r["platform"] for r in new_rows)
tech_counts = Counter(r["technique"] for r in new_rows)

print(f"\nTotal rows: {len(new_rows)}")
print(f"\nCategory distribution:")
for cat, count in sorted(cat_counts.items()):
    print(f"  {CATEGORY_NAMES.get(cat, cat)}: {count}")

print(f"\nPlatform distribution:")
for plat, count in sorted(plat_counts.items()):
    print(f"  {plat}: {count}")

print(f"\nStorytelling technique distribution:")
for tech, count in sorted(tech_counts.items()):
    print(f"  {tech}: {count}")
