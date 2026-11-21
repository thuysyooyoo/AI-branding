import os
import glob
import re

print("=== SEARCHING FOR THUY'S NOTES OR FEEDBACK ACROSS PROJECT ===")
search_dirs = ['scratch', 'reports', 'knowledge', 'workflows']
keywords = ['thúy', 'chị thúy', 'c thúy', 'feedback', 'nhận xét', 'sửa lại', 'yêu cầu sửa', 'revision', 'ghi chú']

for sdir in search_dirs:
    for root, dirs, files in os.walk(sdir):
        for f in files:
            if f.endswith('.md') or f.endswith('.txt') or f.endswith('.json'):
                path = os.path.join(root, f)
                try:
                    with open(path, 'r', encoding='utf-8') as file:
                        content = file.read()
                        # look for specific patterns
                        for line_no, line in enumerate(content.splitlines(), 1):
                            lower = line.lower()
                            if any(k in lower for k in ['chị thúy', 'c thúy', 'yêu cầu sửa', 'cần sửa', 'sửa đổi', 'phản hồi', 'chưa đạt', 'feedback']):
                                # exclude standard templates if possible
                                if not ('chị nguyễn thị ngọc thúy đã phê duyệt' in lower or 'nguyễn thị ngọc thúy' in lower and len(line) < 50):
                                    print(f"[{path}:{line_no}] {line.strip()[:140]}")
                except Exception as e:
                    pass

