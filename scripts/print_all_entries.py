import openpyxl

excel_path = r"d:\AI AGENT THUY\AI branding\reports\plan1\Weekly_Content_Calendar.xlsx"
wb = openpyxl.load_workbook(excel_path, data_only=True)

ws_fb = wb['Facebook']
ws_yt = wb['YouTube']

print("=== ALL FACEBOOK ROWS IN SHEET 2 ===")
for r in range(5, ws_fb.max_row + 1):
    vals = [ws_fb.cell(r, c).value for c in range(1, 12)]
    stt, po, date_str, title, seo, link, content, visual, fc, report, pub = vals
    if po or title:
        has_c = bool(content and len(str(content).strip()) > 30)
        print(f"R{r:2d} | STT {stt} | PO: {po:18s} | Date: {str(date_str):25s} | HasC: {str(has_c):5s} | FC: {fc} | Pub: {pub}")

print("\n=== ALL YOUTUBE ROWS IN SHEET 3 ===")
for r in range(5, ws_yt.max_row + 1):
    vals = [ws_yt.cell(r, c).value for c in range(1, 12)]
    stt, po, title, cat, seo, link, content, visual, fc, report, pub = vals
    if po or title:
        has_c = bool(content and len(str(content).strip()) > 30)
        print(f"R{r:2d} | STT {stt} | PO: {po:18s} | HasC: {str(has_c):5s} | FC: {fc} | Pub: {pub} | Title: {str(title)[:45]}")
