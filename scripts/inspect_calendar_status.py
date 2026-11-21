import openpyxl
import os
import glob
import re

excel_path = r"d:\AI AGENT THUY\AI branding\reports\plan1\Weekly_Content_Calendar.xlsx"
wb = openpyxl.load_workbook(excel_path, data_only=True)

ws_wc = wb['Weekly Calendar']
ws_fb = wb['Facebook']
ws_yt = wb['YouTube']

# 1. Analyze Sheet 2: Facebook
print("=== SHEET 2: FACEBOOK DETAILED SCAN ===")
fb_pos = {}
fb_needs_work = []
for r in range(5, ws_fb.max_row + 1):
    vals = [ws_fb.cell(r, c).value for c in range(1, 12)]
    stt, po, date_str, title, seo, link, content, visual, fc, report, pub_status = vals
    if not po and not title:
        continue
    has_content = bool(content and len(str(content).strip()) > 30)
    fb_pos[po] = {
        'row': r, 'stt': stt, 'po': po, 'date': date_str, 'title': title,
        'has_content': has_content, 'content_len': len(str(content or '')),
        'fc': fc, 'report': report, 'pub_status': pub_status, 'link': link
    }
    # Check if empty, or needs revision, or notes
    is_empty = not has_content
    is_rev = fc and 'REVISION' in str(fc).upper()
    is_escalated = fc and 'ESCALATED' in str(fc).upper()
    
    # check for notes from Thuy
    notes_from_thuy = []
    for field_name, field_val in [('PubStatus', pub_status), ('Report', report), ('Title', title)]:
        if field_val:
            s = str(field_val)
            for kw in ['sửa', 'yêu cầu', 'chỉnh', 'chị thúy', 'c thúy', 'thay', 'lại']:
                if kw in s.lower() and not ('sẵn sàng đăng' in s.lower() and len(s) < 20):
                    notes_from_thuy.append(f"{field_name}: {s}")
                    
    if is_empty or is_rev or is_escalated or notes_from_thuy:
        fb_needs_work.append((r, po, title, is_empty, fc, pub_status, notes_from_thuy))

print(f"Total FB entries: {len(fb_pos)}")
print(f"FB entries needing work ({len(fb_needs_work)}):")
for item in fb_needs_work:
    print("  ", item)

# 2. Analyze Sheet 3: YouTube
print("\n=== SHEET 3: YOUTUBE DETAILED SCAN ===")
yt_pos = {}
yt_needs_work = []
for r in range(5, ws_yt.max_row + 1):
    vals = [ws_yt.cell(r, c).value for c in range(1, 12)]
    stt, po, title, cat, seo, link, content, visual, fc, report, pub_status = vals
    if not po and not title:
        continue
    has_content = bool(content and len(str(content).strip()) > 30)
    yt_pos[po] = {
        'row': r, 'stt': stt, 'po': po, 'title': title,
        'has_content': has_content, 'content_len': len(str(content or '')),
        'fc': fc, 'report': report, 'pub_status': pub_status, 'link': link
    }
    is_empty = not has_content
    is_rev = fc and 'REVISION' in str(fc).upper()
    is_escalated = fc and 'ESCALATED' in str(fc).upper()
    
    notes_from_thuy = []
    for field_name, field_val in [('PubStatus', pub_status), ('Report', report), ('Title', title)]:
        if field_val:
            s = str(field_val)
            for kw in ['sửa', 'yêu cầu', 'chỉnh', 'chị thúy', 'c thúy', 'thay', 'lại']:
                if kw in s.lower() and not ('sẵn sàng dựng' in s.lower() or 'đã duyệt quay' in s.lower()):
                    notes_from_thuy.append(f"{field_name}: {s}")
                    
    if is_empty or is_rev or is_escalated or notes_from_thuy:
        yt_needs_work.append((r, po, title, is_empty, fc, pub_status, notes_from_thuy))

print(f"Total YT entries: {len(yt_pos)}")
print(f"YT entries needing work ({len(yt_needs_work)}):")
for item in yt_needs_work:
    print("  ", item)

# 3. Analyze Sheet 1 (Weekly Calendar) vs POs
print("\n=== SHEET 1 (Weekly Calendar) VS PRODUCED ARTICLES ===")
wc_items = []
unproduced_wc = []
for r in range(5, ws_wc.max_row + 1):
    vals = [ws_wc.cell(r, c).value for c in range(1, 12)]
    stt, date_str, platform, cat, tech, fmt, opt_a, opt_b, opt_c, ssot, status = vals
    if stt is None and date_str is None:
        continue
    # determine PO pattern for this STT
    # Look at POs in scratch/production_orders or PO format
    wc_items.append({
        'row': r, 'stt': stt, 'date': date_str, 'platform': platform,
        'cat': cat, 'opt_a': opt_a, 'status': status
    })

print(f"Total Weekly Calendar items: {len(wc_items)}")

