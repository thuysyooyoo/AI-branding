import openpyxl

wb_path = r"d:\AI AGENT THUY\AI branding\reports\plan1\Weekly_Content_Calendar.xlsx"
wb = openpyxl.load_workbook(wb_path, data_only=True)

print("=== CHECKING FOR ANY REVISION REQUESTS ACROSS ALL SHEETS ===")
revision_found = []

for sname in wb.sheetnames:
    ws = wb[sname]
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            val = ws.cell(r, c).value
            if val is not None:
                val_lower = str(val).lower()
                # Check for revision indicators
                for rev_word in ['yêu cầu sửa', 'cần sửa', 'chỉnh sửa lại', 'viết lại', 'revision required', 'escalated', 'không đạt', 'sửa lại', 'từ chối']:
                    if rev_word in val_lower:
                        revision_found.append((sname, r, c, ws.cell(4, c).value, val))

print(f"Total revision requests found: {len(revision_found)}")
for item in revision_found:
    print(f"[{item[0]}] Row {item[1]}, Col {item[2]} ({item[3]}): {str(item[4])[:120]}")
