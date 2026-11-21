# -*- coding: utf-8 -*-
"""
XUẤT KỊCH BẢN YOUTUBE THUYTALK (.MD) SANG WORD (.DOCX) CHUYÊN NGHIỆP
Hỗ trợ 5 kịch bản: HOT47, HOT48, HOT49, HOT50, HOT51
Lưu đồng bộ vào:
1. Local: d:\AI AGENT THUY\AI branding\reports\plan1\video_scripts\<Ten_File>.docx
2. OneDrive: ...\BRANDING\youtube\<Kịch bản CM1 / CM2 / CM4>\<Ten_File>.docx
"""

import os
import sys
import re
import docx
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

try:
    sys.stdout.reconfigure(encoding="utf-8")
except AttributeError:
    pass

LOCAL_DIR = r"d:\AI AGENT THUY\AI branding\reports\plan1\video_scripts"
ONEDRIVE_YT = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\youtube"

SCRIPTS_TO_EXPORT = [
    {
        "md": "YS-PO-2026-W46-HOT47.md",
        "docx": "ThuyTalk_CM2-HOT47_Xuat_Tra_Hang_Nhap_Khau_Loai_Hinh_B13.docx",
        "onedrive_sub": "Kịch bản CM2",
    },
    {
        "md": "YS-PO-2026-W46-HOT48.md",
        "docx": "ThuyTalk_CM2-HOT48_Nhap_Hang_Mau_Mien_KTCN_4_Cai_Bay.docx",
        "onedrive_sub": "Kịch bản CM2",
    },
    {
        "md": "YS-PO-2026-W46-HOT49.md",
        "docx": "ThuyTalk_CM2-HOT49_May_Moc_Cu_Duoi_10_Nam_5_Rao_Can_QD18.docx",
        "onedrive_sub": "Kịch bản CM2",
    },
    {
        "md": "YS-PO-2026-W46-HOT50.md",
        "docx": "ThuyTalk_CM4-HOT50_Cu_Lua_Mua_Hang_Term_CIF_Bay_Local_Charges.docx",
        "onedrive_sub": "Kịch bản CM4",
    },
    {
        "md": "YS-PO-2026-W46-HOT51.md",
        "docx": "ThuyTalk_CM1-HOT51_Cong_Van_24083_Tham_Van_Gia_Va_KTSTQ.docx",
        "onedrive_sub": "Kịch bản CM1",
    },
]

NAVY = RGBColor(0x1B, 0x36, 0x5D)
TEAL = RGBColor(0x0F, 0x52, 0x8A)
DARK_TEXT = RGBColor(0x22, 0x22, 0x22)
MUTED_TEXT = RGBColor(0x55, 0x5E, 0x68)
AMBER_TEXT = RGBColor(0x92, 0x40, 0x0E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)


def sanitize_markdown_text(raw_text: str) -> str:
    """Làm sạch ký tự lỗi LaTeX/CR trong các file .md nguồn trước khi parse sang Word."""
    # Fix bị gãy dòng do "\r" trong chuỗi python cũ sinh ra "$\r\nightarrow$" hoặc "$<CR>ightarrow$"
    text = raw_text.replace("\r\n", "\n")
    text = re.sub(r"\$\s*[\r\n]*\s*ightarrow\$", " → ", text)
    text = re.sub(r"\$\s*\\rightarrow\s*\$", " → ", text)
    text = text.replace("\r", " ")
    # Fix typo nhẹ tiêu đề Hồi 4 của HOT47 ("CHốt CHẶN" -> "CHỐT CHẶN")
    text = text.replace("CHốt CHẶN", "CHỐT CHẶN")
    return text


def set_cell_shading(cell, hex_color: str):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def set_cell_margins(cell, top=90, bottom=90, left=130, right=130):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def set_table_borders(table, border_color="CBD5E1"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="6" w:space="0" w:color="{border_color}"/>'
        f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="{border_color}"/>'
        f'<w:left w:val="single" w:sz="6" w:space="0" w:color="{border_color}"/>'
        f'<w:right w:val="single" w:sz="6" w:space="0" w:color="{border_color}"/>'
        f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="{border_color}"/>'
        f'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="{border_color}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)


def set_paragraph_shading_and_border(paragraph, fill_hex: str, left_border_hex: str = None, bottom_border_hex: str = None):
    pPr = paragraph._p.get_or_add_pPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill_hex}"/>')
    pPr.append(shd)
    if left_border_hex or bottom_border_hex:
        bdr_xml = f'<w:pBdr {nsdecls("w")}>'
        if left_border_hex:
            bdr_xml += f'<w:left w:val="single" w:sz="18" w:space="8" w:color="{left_border_hex}"/>'
        if bottom_border_hex:
            bdr_xml += f'<w:bottom w:val="single" w:sz="8" w:space="4" w:color="{bottom_border_hex}"/>'
        bdr_xml += "</w:pBdr>"
        pPr.append(parse_xml(bdr_xml))


def add_formatted_runs(paragraph, text: str, base_size=11.5, base_color=DARK_TEXT, base_bold=False, base_italic=False):
    """Parse **bold**, *italic*, `code`, and <br> into clean docx runs."""
    # Split <br> lines first
    sublines = re.split(r"<br\s*/?>", text)
    for line_idx, sub in enumerate(sublines):
        if line_idx > 0:
            paragraph.add_run("\n")
        # Tokenize **bold**, ***bolditalic***, *italic*, `code`
        pattern = re.compile(r"(\*\*\*.*?\*\*\*|\*\*.*?\*\*|\*[^*\n]+?\*|`[^`\n]+?`)")
        pos = 0
        for m in pattern.finditer(sub):
            if m.start() > pos:
                r = paragraph.add_run(sub[pos:m.start()])
                r.font.name = "Calibri"
                r.font.size = Pt(base_size)
                r.font.color.rgb = base_color
                r.bold = base_bold
                r.italic = base_italic
            tok = m.group(0)
            r = paragraph.add_run()
            r.font.name = "Calibri"
            r.font.size = Pt(base_size)
            r.font.color.rgb = base_color
            if tok.startswith("***") and tok.endswith("***"):
                r.text = tok[3:-3]
                r.bold = True
                r.italic = True
            elif tok.startswith("**") and tok.endswith("**"):
                r.text = tok[2:-2]
                r.bold = True
                r.italic = base_italic
            elif tok.startswith("*") and tok.endswith("*"):
                r.text = tok[1:-1]
                r.bold = base_bold
                r.italic = True
            elif tok.startswith("`") and tok.endswith("`"):
                r.text = tok[1:-1]
                r.bold = True if base_bold else False
                r.font.color.rgb = TEAL
            pos = m.end()
        if pos < len(sub):
            r = paragraph.add_run(sub[pos:])
            r.font.name = "Calibri"
            r.font.size = Pt(base_size)
            r.font.color.rgb = base_color
            r.bold = base_bold
            r.italic = base_italic


def add_page_number_to_run(run):
    fldChar1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls("w"))
    instrText = parse_xml(r'<w:instrText %s xml:space="preserve"> PAGE </w:instrText>' % nsdecls("w"))
    fldChar2 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls("w"))
    fldChar3 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls("w"))
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)


def build_docx_from_markdown(md_path: str, out_paths: list):
    with open(md_path, "r", encoding="utf-8") as f:
        raw_text = f.read()

    clean_text = sanitize_markdown_text(raw_text)
    lines = clean_text.splitlines()

    doc = docx.Document()

    # Setup page margins & font
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.0)

    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Calibri"
    normal_style.font.size = Pt(11.5)
    normal_style.font.color.rgb = DARK_TEXT

    # Header & Footer
    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hr = hp.add_run("ThuyTalk | Kịch bản YouTube Đọc Thu Âm & Sản Xuất — Eureka Logistics")
    hr.font.name = "Calibri"
    hr.font.size = Pt(9)
    hr.font.color.rgb = MUTED_TEXT
    hr.italic = True

    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fr1 = fp.add_run("Trang ")
    fr1.font.name = "Calibri"
    fr1.font.size = Pt(9)
    fr1.font.color.rgb = MUTED_TEXT
    fr_pg = fp.add_run()
    fr_pg.font.name = "Calibri"
    fr_pg.font.size = Pt(9)
    fr_pg.font.color.rgb = MUTED_TEXT
    add_page_number_to_run(fr_pg)

    in_spoken_section = False
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        stripped = line.strip()

        # Skip horizontal rules cleanly (or add light spacing)
        if stripped == "---":
            i += 1
            continue

        if not stripped:
            i += 1
            continue

        # Markdown Table
        if stripped.startswith("|") and i + 1 < len(lines) and lines[i + 1].strip().startswith("|"):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].strip())
                i += 1
            # Parse header + separator + rows
            rows_data = []
            for tl in table_lines:
                cells = [c.strip() for c in tl.strip("|").split("|")]
                # Skip separator row (---, :---:)
                if all(re.match(r"^:?-{2,}:?$", c.strip()) or not c.strip() for c in cells):
                    continue
                rows_data.append(cells)

            if rows_data:
                ncols = max(len(r) for r in rows_data)
                tbl = doc.add_table(rows=len(rows_data), cols=ncols)
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                set_table_borders(tbl)

                # Column widths for A4 usable width ~16.8cm
                if ncols == 5:
                    col_widths = [Cm(1.1), Cm(3.0), Cm(2.9), Cm(5.0), Cm(4.8)]
                elif ncols == 4:
                    col_widths = [Cm(2.8), Cm(3.6), Cm(5.2), Cm(5.2)]
                else:
                    col_widths = [Cm(16.8 / ncols)] * ncols

                for r_idx, r_cells in enumerate(rows_data):
                    row = tbl.rows[r_idx]
                    trPr = row._tr.get_or_add_trPr()
                    trPr.append(OxmlElement("w:cantSplit"))
                    if r_idx == 0:
                        trPr.append(OxmlElement("w:tblHeader"))
                    for c_idx in range(ncols):
                        val = r_cells[c_idx] if c_idx < len(r_cells) else ""
                        cell = row.cells[c_idx]
                        cell.width = col_widths[c_idx]
                        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
                        set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
                        p = cell.paragraphs[0]
                        p.paragraph_format.space_before = Pt(0)
                        p.paragraph_format.space_after = Pt(2)
                        p.paragraph_format.line_spacing = 1.15
                        if r_idx == 0:
                            set_cell_shading(cell, "1B365D")
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            add_formatted_runs(p, val, base_size=10, base_color=WHITE, base_bold=True)
                        else:
                            if r_idx % 2 == 0:
                                set_cell_shading(cell, "F8FAFC")
                            if ncols == 5 and c_idx == 0:
                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            add_formatted_runs(p, val, base_size=9.5, base_color=DARK_TEXT)
                spacer = doc.add_paragraph()
                spacer.paragraph_format.space_before = Pt(0)
                spacer.paragraph_format.space_after = Pt(4)
            continue

        # Title H1 (# ...)
        if stripped.startswith("# "):
            title_txt = stripped[2:].strip()
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.keep_with_next = True
            set_paragraph_shading_and_border(p, "EEF3F9", bottom_border_hex="1B365D")
            add_formatted_runs(p, title_txt, base_size=14.5, base_color=NAVY, base_bold=True)
            i += 1
            continue

        # Section H2 (## ...)
        if stripped.startswith("## "):
            h2_txt = stripped[3:].strip()
            in_spoken_section = "PHẦN 2" in h2_txt.upper()
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.keep_with_next = True
            set_paragraph_shading_and_border(p, "EBF2FA", left_border_hex="1B365D")
            add_formatted_runs(p, "  " + h2_txt, base_size=12.5, base_color=NAVY, base_bold=True)
            i += 1
            continue

        # Act Heading H3 (### ...)
        if stripped.startswith("### "):
            h3_txt = stripped[4:].strip()
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.keep_with_next = True
            set_paragraph_shading_and_border(p, "FFFFFF", bottom_border_hex="CBD5E1")
            add_formatted_runs(p, h3_txt, base_size=12, base_color=TEAL, base_bold=True)
            i += 1
            continue

        # Subheading H4 (#### ...)
        if stripped.startswith("#### "):
            h4_txt = stripped[5:].strip()
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.keep_with_next = True
            add_formatted_runs(p, h4_txt, base_size=11.5, base_color=NAVY, base_bold=True)
            i += 1
            continue

        # Blockquote (> ...) — e.g. Cầu dẫn chuyển tiếp
        if stripped.startswith(">"):
            quote_lines = []
            while i < len(lines) and lines[i].strip().startswith(">"):
                q_line = lines[i].strip()[1:].strip()
                if q_line:
                    quote_lines.append(q_line)
                i += 1
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.4)
            p.paragraph_format.right_indent = Cm(0.3)
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(8)
            p.paragraph_format.line_spacing = 1.25
            set_paragraph_shading_and_border(p, "FFF8E7", left_border_hex="D97706")
            add_formatted_runs(p, "\n".join(quote_lines), base_size=11, base_color=AMBER_TEXT)
            continue

        # Inline Bridge marker (*[CÂU DẪN CHUYỂN TIẾP...]*) + following paragraph
        if stripped.startswith("*[CÂU DẪN CHUYỂN TIẾP"):
            bridge_text = stripped
            if i + 1 < len(lines) and lines[i + 1].strip() and not lines[i + 1].strip().startswith(("#", "---", "*[")):
                bridge_text = stripped.strip("*") + "\n" + lines[i + 1].strip()
                i += 2
            else:
                i += 1
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.4)
            p.paragraph_format.right_indent = Cm(0.3)
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(8)
            p.paragraph_format.line_spacing = 1.25
            set_paragraph_shading_and_border(p, "FFF8E7", left_border_hex="D97706")
            add_formatted_runs(p, f"**{bridge_text.splitlines()[0]}**" + ("\n" + "\n".join(bridge_text.splitlines()[1:]) if "\n" in bridge_text else ""), base_size=11, base_color=AMBER_TEXT)
            continue

        # Bullet list (- ... or * ...) or Numbered list (1. ...)
        bullet_match = re.match(r"^(\s*)([-*+])\s+(.*)$", line)
        num_match = re.match(r"^(\s*)(\d+\.)\s+(.*)$", line)
        if bullet_match:
            indent_spaces = len(bullet_match.group(1))
            content = bullet_match.group(3).strip()
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.left_indent = Cm(0.8 + (0.5 if indent_spaces >= 2 else 0))
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.2
            add_formatted_runs(p, content, base_size=11 if not in_spoken_section else 11.5, base_color=DARK_TEXT)
            i += 1
            continue

        if num_match:
            indent_spaces = len(num_match.group(1))
            num_prefix = num_match.group(2)
            content = num_match.group(3).strip()
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.6 + (0.4 if indent_spaces >= 2 else 0))
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.2
            add_formatted_runs(p, f"**{num_prefix}** {content}", base_size=11, base_color=DARK_TEXT)
            i += 1
            continue

        # Italic instruction note line *(...)*
        if stripped.startswith("*(") and stripped.endswith(")*"):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(6)
            add_formatted_runs(p, stripped, base_size=10, base_color=MUTED_TEXT, base_italic=True)
            i += 1
            continue

        # Normal paragraph (Metadata or Spoken Dialogue)
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(6 if in_spoken_section else 3)
        p.paragraph_format.line_spacing = 1.28 if in_spoken_section else 1.18
        add_formatted_runs(p, stripped, base_size=11.5 if in_spoken_section else 10.8, base_color=DARK_TEXT)
        i += 1

    # Save atomically to each destination
    for out_path in out_paths:
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        tmp_path = out_path + ".tmp.docx"
        doc.save(tmp_path)
        os.replace(tmp_path, out_path)
        print(f"   ✅ Đã xuất Word (.docx): {out_path}")


if __name__ == "__main__":
    print("=" * 80)
    print("📄 BẮT ĐẦU XUẤT 5 KỊCH BẢN YOUTUBE THUYTALK SANG WORD (.DOCX)")
    print("=" * 80)
    for item in SCRIPTS_TO_EXPORT:
        md_file = os.path.join(LOCAL_DIR, item["md"])
        if not os.path.exists(md_file):
            print(f"   ❌ Không tìm thấy file nguồn: {md_file}")
            sys.exit(1)
        dest_local = os.path.join(LOCAL_DIR, item["docx"])
        dest_onedrive = os.path.join(ONEDRIVE_YT, item["onedrive_sub"], item["docx"])
        print(f"\n📌 Đang chuyển đổi: {item['md']}")
        build_docx_from_markdown(md_file, [dest_local, dest_onedrive])
    print("\n🎉 HOÀN TẤT XUẤT TOÀN BỘ FILE WORD (.DOCX)!")
