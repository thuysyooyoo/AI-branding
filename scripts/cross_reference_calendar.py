import openpyxl
import os
import glob
import re

excel_path = r"d:\AI AGENT THUY\AI branding\reports\plan1\Weekly_Content_Calendar.xlsx"
wb = openpyxl.load_workbook(excel_path, data_only=True)

ws_wc = wb['Weekly Calendar']
ws_fb = wb['Facebook']
ws_yt = wb['YouTube']

# Build map of FB items: by PO, by Title, by Date
fb_by_po = {}
fb_by_title = {}
for r in range(5, ws_fb.max_row + 1):
    po = ws_fb.cell(r, 2).value
    title = ws_fb.cell(r, 4).value
    content = ws_fb.cell(r, 7).value
    fc = ws_fb.cell(r, 9).value
    pub = ws_fb.cell(r, 11).value
    has_content = bool(content and len(str(content).strip()) > 30)
    if po:
        fb_by_po[str(po).strip()] = {'row': r, 'title': title, 'has_content': has_content, 'fc': fc, 'pub': pub}
    if title:
        fb_by_title[str(title).strip().lower()] = {'row': r, 'po': po, 'has_content': has_content, 'fc': fc, 'pub': pub}

# Build map of YT items: by PO, by Title
yt_by_po = {}
yt_by_title = {}
for r in range(5, ws_yt.max_row + 1):
    po = ws_yt.cell(r, 2).value
    title = ws_yt.cell(r, 3).value
    content = ws_yt.cell(r, 7).value
    fc = ws_yt.cell(r, 9).value
    pub = ws_yt.cell(r, 11).value
    has_content = bool(content and len(str(content).strip()) > 30)
    if po:
        yt_by_po[str(po).strip()] = {'row': r, 'title': title, 'has_content': has_content, 'fc': fc, 'pub': pub}
    if title:
        yt_by_title[str(title).strip().lower()] = {'row': r, 'po': po, 'has_content': has_content, 'fc': fc, 'pub': pub}

print("=== CROSS-REFERENCING SHEET 1 (STT 1 to 70) ===")
unproduced_in_wc = []

for r in range(5, ws_wc.max_row + 1):
    stt = ws_wc.cell(r, 1).value
    date_str = ws_wc.cell(r, 2).value
    platform = ws_wc.cell(r, 3).value
    opt_a = ws_wc.cell(r, 7).value
    status = ws_wc.cell(r, 11).value
    if stt is None:
        continue
    
    # Expected PO format: PO-2026-WXX-YY
    # Let's see if we can find matching PO or Title in FB/YT
    matched = False
    details = None
    
    # Check if STT corresponds to PO:
    # e.g., STT 1 -> W33-01? Let's check files in scratch/production_orders/
    # Find matching PO files
    po_candidates = glob.glob(f"scratch/production_orders/PO-*-{stt:02d}.md")
    po_candidates += glob.glob(f"scratch/production_orders/PO-*-{stt}.md")
    
    po_code = None
    if po_candidates:
        po_code = os.path.basename(po_candidates[0]).replace('.md', '')
    
    # Check in FB or YT
    if platform == 'Facebook':
        if po_code and po_code in fb_by_po:
            matched = True
            details = fb_by_po[po_code]
        else:
            # check by title match
            opt_lower = str(opt_a).strip().lower() if opt_a else ''
            for t, data in fb_by_title.items():
                if opt_lower and (opt_lower[:30] in t or t[:30] in opt_lower):
                    matched = True
                    details = data
                    break
    elif platform == 'YouTube':
        if po_code and po_code in yt_by_po:
            matched = True
            details = yt_by_po[po_code]
        else:
            opt_lower = str(opt_a).strip().lower() if opt_a else ''
            for t, data in yt_by_title.items():
                if opt_lower and (opt_lower[:30] in t or t[:30] in opt_lower):
                    matched = True
                    details = data
                    break

    if not matched or not (details and details['has_content']):
        unproduced_in_wc.append((stt, date_str, platform, po_code, opt_a, status, matched, details))
        print(f"MISSING/INCOMPLETE: STT {stt:2d} | {platform:8s} | PO: {po_code} | {date_str} | OptA: {str(opt_a)[:40]} | Status: {status} | Matched: {matched}")

print(f"\nTotal missing/incomplete from Sheet 1: {len(unproduced_in_wc)}")
