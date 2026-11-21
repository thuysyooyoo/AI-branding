import glob
import os
import re

po_files = sorted(glob.glob('scratch/production_orders/PO-*.md'))
draft_fb_files = set([os.path.basename(f) for f in glob.glob('scratch/production_orders/DRAFT-FB-*.md')])
draft_yt_files = set([os.path.basename(f) for f in glob.glob('scratch/production_orders/DRAFT-YT-*.md')])

print(f"Total PO files found: {len(po_files)}")

po_status = []
for p in po_files:
    fname = os.path.basename(p)
    po_code = fname.replace('.md', '')
    
    # Read PO platform and title
    with open(p, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Find platform
    platform = 'Unknown'
    m_plat = re.search(r'\*\*(?:Nền tảng|Platform|Kênh).*?:\*\*\s*([^\n\r]+)', content, re.IGNORECASE)
    if m_plat:
        platform = m_plat.group(1).strip()
    
    # Find title
    title = ''
    m_title = re.search(r'\*\*(?:Chủ đề|Tiêu đề|Title).*?:\*\*\s*([^\n\r]+)', content, re.IGNORECASE)
    if m_title:
        title = m_title.group(1).strip()
    elif '# PHIẾU GIAO VIỆC' in content:
        lines = content.splitlines()
        for l in lines[:10]:
            if l.startswith('# '):
                title = l.replace('# ', '').strip()
    
    has_fb_draft = f"DRAFT-FB-{po_code}.md" in draft_fb_files
    has_yt_draft = any(po_code in f for f in draft_yt_files)
    
    po_status.append({
        'po': po_code,
        'platform': platform,
        'title': title,
        'has_fb': has_fb_draft,
        'has_yt': has_yt_draft,
        'has_draft': has_fb_draft or has_yt_draft
    })

no_draft = [p for p in po_status if not p['has_draft']]
print(f"POs without drafts: {len(no_draft)}")
for p in no_draft[:30]:
    print(f"PO: {p['po']:20s} | Plat: {p['platform']:25s} | Title: {p['title'][:60]}")
if len(no_draft) > 30:
    print(f"... and {len(no_draft) - 30} more")
