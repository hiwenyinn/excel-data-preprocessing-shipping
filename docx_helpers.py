import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def add_header_footer(doc):
    for s in doc.sections:
        s.top_margin = Inches(1)
        s.bottom_margin = Inches(1)
        s.left_margin = Inches(1)
        s.right_margin = Inches(1)
        
        # 頁首
        header = s.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hrun = hp.add_run("航運行政人員 Excel 資料前處理實戰指南 | 航運數據分析基礎篇")
        hrun.font.name = "微軟正黑體"
        hrun.font.size = Pt(8.5)
        hrun.font.color.rgb = RGBColor(128, 128, 128)
        
        # 頁尾
        footer = s.footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        frun1 = fp.add_run("Page ")
        frun1.font.name = "Calibri"
        frun1.font.size = Pt(9)
        frun1.font.color.rgb = RGBColor(128, 128, 128)
        
        # 插入 Word 頁碼欄位 PAGE
        fldSimple = parse_xml(r'<w:fldSimple %s w:instr="PAGE"/>' % nsdecls('w'))
        fp._p.append(fldSimple)
        
        frun2 = fp.add_run(" of ")
        frun2.font.name = "Calibri"
        frun2.font.size = Pt(9)
        frun2.font.color.rgb = RGBColor(128, 128, 128)
        
        fldSimple2 = parse_xml(r'<w:fldSimple %s w:instr="NUMPAGES"/>' % nsdecls('w'))
        fp._p.append(fldSimple2)

def add_toc_field(paragraph):
    p = paragraph._p
    fldSimple = parse_xml(r'<w:fldSimple %s w:instr="TOC \o &quot;1-3&quot; \h \z \u"/>' % nsdecls('w'))
    p.append(fldSimple)

def add_figure_toc_field(paragraph):
    p = paragraph._p
    fldSimple = parse_xml(r'<w:fldSimple %s w:instr="TOC \c &quot;圖&quot; \h \z"/>' % nsdecls('w'))
    p.append(fldSimple)

def add_table_toc_field(paragraph):
    p = paragraph._p
    fldSimple = parse_xml(r'<w:fldSimple %s w:instr="TOC \c &quot;表&quot; \h \z"/>' % nsdecls('w'))
    p.append(fldSimple)

def style_paragraph(p, space_before=4, space_after=6, line_spacing=1.15):
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing

def add_callout(doc, text, title="重點提示", callout_type="info"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    
    bg = "F0F4F8"
    border_color = "0078D4"
    if callout_type == "warning":
        bg = "FFF4CE"
        border_color = "D83B01"
    elif callout_type == "success":
        bg = "DFF6DD"
        border_color = "107C41"
        
    set_cell_background(cell, bg)
    set_cell_margins(cell, top=140, bottom=140, left=200, right=200)
    
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'''<w:tcBorders {nsdecls("w")}>
        <w:top w:val="none"/>
        <w:left w:val="single" w:sz="24" w:space="0" w:color="{border_color}"/>
        <w:bottom w:val="none"/>
        <w:right w:val="none"/>
    </w:tcBorders>''')
    tcPr.append(tcBorders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    
    r_title = p.add_run(f"【{title}】\n")
    r_title.bold = True
    r_title.font.name = "微軟正黑體"
    r_title.font.size = Pt(10.5)
    r_title.font.color.rgb = RGBColor(16, 124, 65) if callout_type=="success" else (RGBColor(216, 59, 1) if callout_type=="warning" else RGBColor(0, 120, 212))
    
    r_body = p.add_run(text)
    r_body.font.name = "微軟正黑體"
    r_body.font.size = Pt(9.5)
    r_body.font.color.rgb = RGBColor(50, 49, 48)
    
    doc.add_paragraph() # 間距

print("Docx helper module ready!")
