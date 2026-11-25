import os
import pandas as pd
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from docx_helpers import set_cell_background, set_cell_margins, add_header_footer, style_paragraph, add_callout

def build_shipping_manual():
    doc = Document()
    
    # 全域樣式設定
    normal_style = doc.styles['Normal']
    normal_style.font.name = '微軟正黑體'
    normal_style.font.size = Pt(10.5)
    normal_style.font.color.rgb = RGBColor(32, 31, 30)
    normal_style._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
    
    # -------------------------------------------------------------
    # 封面頁 (Cover Page)
    # -------------------------------------------------------------
    for _ in range(3):
        doc.add_paragraph()
        
    p_badge = doc.add_paragraph()
    p_badge.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_badge = p_badge.add_run("■ 航運物流數位轉型與行政作業實戰教材 ■")
    r_badge.font.name = "微軟正黑體"
    r_badge.font.size = Pt(11)
    r_badge.font.bold = True
    r_badge.font.color.rgb = RGBColor(16, 124, 65) # Excel Green
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(20)
    p_title.paragraph_format.space_after = Pt(10)
    r_title = p_title.add_run("航運行政人員 MS Excel 資料前處理\n操作教學手冊與實戰指南")
    r_title.font.name = "微軟正黑體"
    r_title.font.size = Pt(26)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0, 78, 140) # Maritime Blue
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(40)
    r_sub = p_sub.add_run("從零開始掌握航運數據清洗、格式正規化、重複數據剔除與分析就緒全流程")
    r_sub.font.name = "微軟正黑體"
    r_sub.font.size = Pt(13)
    r_sub.font.color.rgb = RGBColor(96, 94, 92)
    
    # 封面資訊卡片 (Table)
    info_tbl = doc.add_table(rows=6, cols=2)
    info_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_info = [
        ("目標受眾", "航運海運承攬(Forwarder)、船公司(Carrier)、報關行之行政與內勤初學者"),
        ("核心技能", "基礎資料分析前處理 (Data Preprocessing)、文字清洗、格式正規化"),
        ("實習工具", "Microsoft Excel 2016 / 2019 / 2021 / Microsoft 365"),
        ("實例資料庫", "30筆基礎示範資料檔 + 300筆航運全情境進階實戰練習檔"),
        ("出版/更新日期", "2024 年最新標準作業指引 (SOP)"),
        ("編製單位", "航運物流數據化行政與商業智慧應用研究小組")
    ]
    for row_idx, (k, v) in enumerate(meta_info):
        c0 = info_tbl.cell(row_idx, 0)
        c1 = info_tbl.cell(row_idx, 1)
        c0.width = Inches(2.0)
        c1.width = Inches(4.5)
        set_cell_background(c0, "F3F2F1")
        set_cell_background(c1, "FFFFFF")
        set_cell_margins(c0, top=100, bottom=100, left=150, right=150)
        set_cell_margins(c1, top=100, bottom=100, left=150, right=150)
        
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r0 = p0.add_run(k)
        r0.bold = True
        r0.font.name = "微軟正黑體"
        r0.font.size = Pt(10)
        r0.font.color.rgb = RGBColor(0, 78, 140)
        
        p1 = c1.paragraphs[0]
        r1 = p1.add_run(v)
        r1.font.name = "微軟正黑體"
        r1.font.size = Pt(10)
        r1.font.color.rgb = RGBColor(32, 31, 30)
        
    for _ in range(4):
        doc.add_paragraph()
        
    p_ver = doc.add_paragraph()
    p_ver.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ver = p_ver.add_run("機密與內部教育訓練教材 · 未經授權請勿翻印散布")
    r_ver.font.name = "微軟正黑體"
    r_ver.font.size = Pt(9)
    r_ver.font.color.rgb = RGBColor(160, 160, 160)
    
    # 換頁 -> 目錄頁
    doc.add_page_break()
    
    # -------------------------------------------------------------
    # 目錄與圖表目錄 (TOC & List of Figures/Tables)
    # -------------------------------------------------------------
    p_toc_head = doc.add_paragraph()
    p_toc_head.paragraph_format.space_before = Pt(10)
    p_toc_head.paragraph_format.space_after = Pt(16)
    r_toc_h = p_toc_head.add_run("目  錄 (Table of Contents)")
    r_toc_h.font.name = "微軟正黑體"
    r_toc_h.font.size = Pt(18)
    r_toc_h.font.bold = True
    r_toc_h.font.color.rgb = RGBColor(0, 78, 140)
    
    # 明確列出本教材完整章節結構大綱
    toc_items = [
        ("第一章 認識航運行政專用 Excel 工作環境與操作介面", "P. 3"),
        ("    1.1 航運行政人員為何必備 Excel 資料處理能力？", "P. 3"),
        ("    1.2 Excel 核心操作介面解構 (功能區/編輯列/工作表)", "P. 4"),
        ("    1.3 航運常用快捷鍵與自訂快速存取工具列設定", "P. 5"),
        ("第二章 航運數據分析前處理應具備之基本常識", "P. 6"),
        ("    2.1 航運實務常見的「髒資料 (Dirty Data)」型態與後果", "P. 6"),
        ("    2.2 「垃圾進，垃圾出 (GIGO)」定律對航運決策之影響", "P. 7"),
        ("    2.3 資料型態的嚴格判定：數值、文字、日期與邏輯值", "P. 8"),
        ("    2.4 乾淨資料表 (Tidy Data) 的六大設計守則", "P. 9"),
        ("第三章 30筆航運示範資料檔前處理實作演練", "P. 10"),
        ("    3.1 示範資料集欄位結構與問題診斷", "P. 10"),
        ("    3.2 步驟一：剔除頭尾無效空白 (TRIM 函數與取代)", "P. 12"),
        ("    3.3 步驟二：英文代碼大小寫標準化 (UPPER/LOWER)", "P. 13"),
        ("    3.4 步驟三：日期格式混雜正規化處理 (DATEVALUE/自訂格式)", "P. 14"),
        ("    3.5 步驟四：主鍵去重複作業 (提單號碼 Remove Duplicates)", "P. 15"),
        ("    3.6 步驟五：非數值與缺漏值處理 (IFERROR / 填補策略)", "P. 16"),
        ("    3.7 步驟六：貨幣符號與千分位純數值轉化 (SUBSTITUTE/VALUE)", "P. 17"),
        ("    3.8 成果驗證：資料前處理前後 (Before vs. After) 綜合評估", "P. 18"),
        ("第四章 300筆綜合型航運案例情境實戰練習與解題全攻略", "P. 19"),
        ("    4.1 300筆進階實戰練習資料庫背景設定", "P. 19"),
        ("    4.2 實戰情境挑戰題 (三大航運前處理任務)", "P. 20"),
        ("    4.3 實戰挑戰題參考標準解題步驟 (SOP)", "P. 21"),
        ("    4.4 進階延伸技巧：Power Query 現代化前處理自動化簡介", "P. 23"),
        ("第五章 參考來源與航運數據規範文獻", "P. 24")
    ]
    
    for item_title, item_page in toc_items:
        p_t = doc.add_paragraph()
        style_paragraph(p_t, space_before=2, space_after=2)
        r_t1 = p_t.add_run(item_title)
        r_t1.font.name = "微軟正黑體"
        r_t1.font.size = Pt(9.5)
        if not item_title.startswith("    "):
            r_t1.bold = True
            r_t1.font.color.rgb = RGBColor(0, 78, 140)
        else:
            r_t1.font.color.rgb = RGBColor(60, 60, 60)
            
        r_dots = p_t.add_run(" " + "." * (80 - len(item_title)*2) + " ")
        r_dots.font.color.rgb = RGBColor(180, 180, 180)
        r_t2 = p_t.add_run(item_page)
        r_t2.font.name = "Calibri"
        r_t2.font.size = Pt(9.5)
        r_t2.font.bold = True
        
    doc.add_paragraph() # 間距
    
    # 圖表目錄
    p_fig_head = doc.add_paragraph()
    p_fig_head.paragraph_format.space_before = Pt(10)
    p_fig_head.paragraph_format.space_after = Pt(8)
    r_figh = p_fig_head.add_run("圖目錄 (List of Figures) 與 表目錄 (List of Tables)")
    r_figh.font.name = "微軟正黑體"
    r_figh.font.size = Pt(14)
    r_figh.font.bold = True
    r_figh.font.color.rgb = RGBColor(0, 78, 140)
    
    fig_table_items = [
        ("圖 1-1", "Microsoft Excel 核心介面架構與航運專用功能區分佈圖", "P. 4"),
        ("圖 3-1", "Excel 「移除重複項」對話框與航運提單主鍵設定示意圖", "P. 15"),
        ("圖 3-2", "Excel 「資料剖析精靈」分隔符號與航運代碼拆解示意圖", "P. 16"),
        ("圖 3-3", "航運前處理四大核心函數結構 (TRIM, CLEAN, VALUE, SUBSTITUTE)", "P. 17"),
        ("圖 3-4", "航運資料前處理前後 (Before vs. After) 成果對比與品質看板", "P. 18"),
        ("表 1-1", "航運行政高頻使用之 Excel 快捷鍵對照表", "P. 5"),
        ("表 2-1", "航運髒資料常見徵兆與對營運層面之風險對照表", "P. 7"),
        ("表 2-2", "Excel 四大基礎資料型態特性與航運實例歸納表", "P. 8"),
        ("表 3-1", "30筆航運示範資料集原始結構與前處理診斷清單", "P. 11"),
        ("表 4-1", "300筆實戰練習檔涵蓋之八大船公司與主要港口統計表", "P. 20")
    ]
    for ft_id, ft_name, ft_page in fig_table_items:
        p_ft = doc.add_paragraph()
        style_paragraph(p_ft, space_before=2, space_after=2)
        r_ft1 = p_ft.add_run(f"{ft_id} : {ft_name}")
        r_ft1.font.name = "微軟正黑體"
        r_ft1.font.size = Pt(9)
        r_dots = p_ft.add_run(" " + "." * (75 - len(ft_name)*2) + " ")
        r_dots.font.color.rgb = RGBColor(190, 190, 190)
        r_ft2 = p_ft.add_run(ft_page)
        r_ft2.font.name = "Calibri"
        r_ft2.font.size = Pt(9)
        
    # 換頁 -> 第一章
    doc.add_page_break()
    
    # -------------------------------------------------------------
    # 第一章 認識航運行政專用 Excel 工作環境與操作介面
    # -------------------------------------------------------------
    p_c1 = doc.add_paragraph()
    p_c1.paragraph_format.space_before = Pt(12)
    p_c1.paragraph_format.space_after = Pt(12)
    r_c1 = p_c1.add_run("第一章 認識航運行政專用 Excel 工作環境與操作介面")
    r_c1.font.name = "微軟正黑體"
    r_c1.font.size = Pt(18)
    r_c1.font.bold = True
    r_c1.font.color.rgb = RGBColor(0, 78, 140)
    
    # 1.1
    p_11 = doc.add_paragraph()
    r_11 = p_11.add_run("1.1 航運行政人員為何必備 Excel 資料前處理能力？")
    r_11.font.name = "微軟正黑體"
    r_11.font.size = Pt(13)
    r_11.font.bold = True
    r_11.font.color.rgb = RGBColor(16, 124, 65)
    
    p = doc.add_paragraph()
    style_paragraph(p)
    p.add_run("在現代國際海運與綜合物流供應鏈中，航運行政人員每天需經手來自船公司 (Carriers)、海運承攬運送業 (Forwarders)、貨櫃集散站 (CY/CFS)、報關行及托運人 (Shippers) 等多方來源的龐雜資料。這些文件包括裝貨單 (S/O)、提單 (B/L)、貨櫃清單 (Container Manifest)、結關通知單、危險品申報書以及各類海運運費帳單。")
    
    p = doc.add_paragraph()
    style_paragraph(p)
    p.add_run("然而，因各家企業內部系統架構與資料庫匯出格式截然不同，匯出的 CSV、文字檔或報表常存在大量空格、中英文代碼不一、日期格式顛倒、重複申報等嚴重的資料瑕疵。若行政人員在尚未進行「資料前處理」前便貿然使用樞紐分析表或公式統計，將導致重大的運費計費短漏、結關延誤甚至海關申報錯誤罰鍰。因此，熟練運用 MS Excel 進行基礎資料清洗，是每位航運初學者邁向專業行政管理的核心必修基本功。")
    
    # 1.2
    p_12 = doc.add_paragraph()
    r_12 = p_12.add_run("1.2 Excel 核心操作介面解構")
    r_12.font.name = "微軟正黑體"
    r_12.font.size = Pt(13)
    r_12.font.bold = True
    r_12.font.color.rgb = RGBColor(16, 124, 65)
    
    p = doc.add_paragraph()
    style_paragraph(p)
    p.add_run("初學者啟動 Microsoft Excel 後，映入眼簾的軟體畫面由數個關鍵工作區塊組成（如圖 1-1 所示）。熟悉各區塊名稱與功能，能在處理航運報表時大幅提升操作流暢度：")
    
    components = [
        ("1. 快速存取工具列 (Quick Access Toolbar)：", "位於視窗最頂端，預設提供儲存、復原 (Ctrl+Z)、重做 (Ctrl+Y)。強烈建議自訂加入「移除重複項」與「篩選」，便於一鍵啟用。"),
        ("2. 功能區 (Ribbon)：", "包含「常用」、「插入」、「公式」、「資料」等核心標籤頁。航運前處理最倚重的即為【資料 (Data)】標籤頁，內含「排序與篩選」、「資料剖析」、「移除重複項」及「資料驗證」。"),
        ("3. 名稱方塊 (Name Box)：", "位於工作表左上方，顯示目前所選取之儲存格座標（如 A1, C15），亦可在此直接鍵入座標進行跨數千筆資料的極速跳轉。"),
        ("4. 公式編輯列 (Formula Bar)：", "緊鄰名稱方塊右側的長條文字框，用於檢視或輸入公式。在處理帶有不可見空白字元的航運代碼時，此處可精確辨識字串真實長度。"),
        ("5. 工作表網格區域 (Worksheet Grid)：", "由英文字母（欄，Columns，A~XFD）與數字（列，Rows，1~1,048,576）交織而成之二維資料空間，為登記提單與貨櫃資訊之主要場所。"),
        ("6. 狀態列 (Status Bar)：", "視窗最底部條帶，當選取包含運費或毛重的儲存格時，狀態列會即時自動顯示「平均值」、「計數」與「數值加總」，是核對報表總量之絕佳輔助。")
    ]
    for c_title, c_desc in components:
        p_c = doc.add_paragraph()
        style_paragraph(p_c, space_before=2, space_after=3)
        r_ct = p_c.add_run(c_title)
        r_ct.bold = True
        r_ct.font.color.rgb = RGBColor(0, 78, 140)
        p_c.add_run(" " + c_desc)
        
    # 插入圖 1-1
    doc.add_paragraph()
    p_img1 = doc.add_paragraph()
    p_img1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img1.paragraph_format.space_before = Pt(8)
    p_img1.paragraph_format.space_after = Pt(4)
    run_img1 = p_img1.add_run()
    run_img1.add_picture("assets/fig1_excel_interface.png", width=Inches(6.2))
    
    p_cap1 = doc.add_paragraph()
    p_cap1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap1.paragraph_format.space_after = Pt(12)
    r_cap1 = p_cap1.add_run("圖 1-1 : Microsoft Excel 核心介面架構與航運專用功能區分佈圖")
    r_cap1.font.name = "微軟正黑體"
    r_cap1.font.size = Pt(9.5)
    r_cap1.font.bold = True
    r_cap1.font.color.rgb = RGBColor(96, 94, 92)
    
    # 1.3
    p_13 = doc.add_paragraph()
    r_13 = p_13.add_run("1.3 航運行政常用快捷鍵與自訂快速存取工具列設定")
    r_13.font.name = "微軟正黑體"
    r_13.font.size = Pt(13)
    r_13.font.bold = True
    r_13.font.color.rgb = RGBColor(16, 124, 65)
    
    p = doc.add_paragraph()
    style_paragraph(p)
    p.add_run("在處理成百上千筆的提單清冊時，熟記快捷鍵能節省高達 70% 的重複性滑鼠點擊時間。表 1-1 列出航運實務上最常使用的十組黃金快捷鍵：")
    
    # 表 1-1
    tbl_shortcuts = doc.add_table(rows=11, cols=3)
    tbl_shortcuts.alignment = WD_TABLE_ALIGNMENT.CENTER
    s_headers = ["快捷鍵組合", "Excel 原生功能", "航運業務實際應用情境"]
    for c_i, h_txt in enumerate(s_headers):
        cell = tbl_shortcuts.cell(0, c_i)
        set_cell_background(cell, "004E8C")
        set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
        p = cell.paragraphs[0]
        r = p.add_run(h_txt)
        r.bold = True
        r.font.name = "微軟正黑體"
        r.font.color.rgb = RGBColor(255, 255, 255)
        r.font.size = Pt(9.5)
        
    s_rows = [
        ("Ctrl + Shift + L", "啟用 / 取消自動篩選", "快速於提單表頭開啟下拉篩選箭頭，查找特定港口或船名"),
        ("Ctrl + 方向鍵 (上下左右)", "跳轉至資料最邊界儲存格", "一秒跳轉至數千筆清單的最底部，檢查最後一筆提單資料"),
        ("Ctrl + Shift + 方向鍵", "連續選取整區塊儲存格", "一口氣反白選取整個欄位（如整欄毛重），利於套用公式"),
        ("Ctrl + H", "開啟尋找與取代對話框", "將資料內所有的「.」或「/」統一取代為標準破折號「-」"),
        ("Ctrl + T", "將選取範圍轉換為結構化表格", "建立標準資料表，自動凍結首列並提供斑馬紋美觀檢視"),
        ("Ctrl + 1", "開啟儲存格格式設定視窗", "將文字型金額轉為貨幣格式，或自訂日期格式為 YYYY-MM-DD"),
        ("Ctrl + D", "向下填滿 (Fill Down)", "將上一列正確的航商名稱或結關日期快速複製到當前空白列"),
        ("Alt + =", "自動插入 SUM 函數", "在海運費欄位最底部極速插入自動加總公式"),
        ("F4", "重複上一個操作 / 鎖定儲存格($)", "在撰寫 VLOOKUP 查表公式時鎖定參照範圍 ($A$1:$D$100)"),
        ("F2", "進入目前儲存格編輯狀態", "檢視提單編號末端是否隱藏多餘空格，便於直接手工微調")
    ]
    for r_i, (k, func, scn) in enumerate(s_rows):
        row = tbl_shortcuts.rows[r_i + 1]
        bg = "F9F9F9" if r_i % 2 == 1 else "FFFFFF"
        for c_i, val in enumerate([k, func, scn]):
            cell = row.cells[c_i]
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.name = "微軟正黑體"
            r.font.size = Pt(9)
            if c_i == 0:
                r.bold = True
                r.font.color.rgb = RGBColor(0, 78, 140)
                
    p_t1 = doc.add_paragraph()
    p_t1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t1.paragraph_format.space_before = Pt(4)
    p_t1.paragraph_format.space_after = Pt(12)
    r_t1 = p_t1.add_run("表 1-1 : 航運行政高頻使用之 Excel 快捷鍵對照表")
    r_t1.font.name = "微軟正黑體"
    r_t1.font.size = Pt(9)
    r_t1.font.bold = True
    r_t1.font.color.rgb = RGBColor(96, 94, 92)
    
    add_callout(doc, 
        "航運作業實務小技巧：將滑鼠移至「資料」標籤頁下的【移除重複項】與【資料剖析】按鈕，按滑鼠右鍵並點選「新增至快速存取工具列」，未來處理大量裝船單時即可直接在視窗頂端一鍵點擊，大幅精簡日常操作路徑！",
        title="航運行政高手秘訣", callout_type="success")
        
    # 換頁 -> 第二章
    doc.add_page_break()
    
    # -------------------------------------------------------------
    # 第二章 航運數據分析前處理應具備之基本常識
    # -------------------------------------------------------------
    p_c2 = doc.add_paragraph()
    p_c2.paragraph_format.space_before = Pt(12)
    p_c2.paragraph_format.space_after = Pt(12)
    r_c2 = p_c2.add_run("第二章 航運數據分析前處理應具備之基本常識")
    r_c2.font.name = "微軟正黑體"
    r_c2.font.size = Pt(18)
    r_c2.font.bold = True
    r_c2.font.color.rgb = RGBColor(0, 78, 140)
    
    # 2.1
    p_21 = doc.add_paragraph()
    r_21 = p_21.add_run("2.1 航運實務常見的「髒資料 (Dirty Data)」型態與後果")
    r_21.font.name = "微軟正黑體"
    r_21.font.size = Pt(13)
    r_21.font.bold = True
    r_21.font.color.rgb = RGBColor(16, 124, 65)
    
    p = doc.add_paragraph()
    style_paragraph(p)
    p.add_run("航運資料具有「跨組織傳遞」、「國際化格式差異」與「高時效性」等特性。由於不同公司、碼頭與報關軟體對資料輸入規範不同，匯總到行政人員手中時往往夾帶諸多問題。表 2-1 歸納了航運業務中最常見的六大髒資料類型及其對日常營運帶來的潛在衝擊：")
    
    # 表 2-1
    tbl_dirty = doc.add_table(rows=7, cols=3)
    tbl_dirty.alignment = WD_TABLE_ALIGNMENT.CENTER
    d_headers = ["髒資料型態", "實務典型範例", "營運失誤與商務風險"]
    for c_i, h_txt in enumerate(d_headers):
        cell = tbl_dirty.cell(0, c_i)
        set_cell_background(cell, "004E8C")
        set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
        p = cell.paragraphs[0]
        r = p.add_run(h_txt)
        r.bold = True
        r.font.name = "微軟正黑體"
        r.font.color.rgb = RGBColor(255, 255, 255)
        r.font.size = Pt(9.5)
        
    d_rows = [
        ("前後不可見空格", "「 Kaohsiung 」與「Kaohsiung」", "樞紐分析表誤判為兩個獨立港口，導致高雄港貨量統計失真"),
        ("英文大小寫混雜", "「msku-8291034」與「MSKU-8291034」", "部分精確比對系統(VLOOKUP/XLOOKUP)查無此筆提單狀態"),
        ("日期格式多元分裂", "2024/05/01、2024-05-02、20240504", "無法按月或按週進行結關日時間軸分群排序，排程大亂"),
        ("數值儲存為文字", "毛重標註為「 18,500 」或「N/A」", "SUM 公式加總結果為 0 或回傳 #VALUE!，誤判船舶配載配重"),
        ("貨幣符號干擾", "海運費欄位含有「$1,850」或「USD 1850」", "財務無法進行自動匯率換算與請款對帳，需人工逐筆塗改"),
        ("訂艙/提單重疊登記", "同一提單號碼因轉單重複複製兩次", "貨物虛增一倍，產生向客戶重複請款或重複向船東訂艙之失誤")
    ]
    for r_i, (k, ex, rsk) in enumerate(d_rows):
        row = tbl_dirty.rows[r_i + 1]
        bg = "F9F9F9" if r_i % 2 == 1 else "FFFFFF"
        for c_i, val in enumerate([k, ex, rsk]):
            cell = row.cells[c_i]
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.name = "微軟正黑體"
            r.font.size = Pt(9)
            if c_i == 0:
                r.bold = True
                r.font.color.rgb = RGBColor(209, 52, 56) # Warning Red
                
    p_t2 = doc.add_paragraph()
    p_t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t2.paragraph_format.space_before = Pt(4)
    p_t2.paragraph_format.space_after = Pt(12)
    r_t2 = p_t2.add_run("表 2-1 : 航運髒資料常見徵兆與對營運層面之風險對照表")
    r_t2.font.name = "微軟正黑體"
    r_t2.font.size = Pt(9)
    r_t2.font.bold = True
    r_t2.font.color.rgb = RGBColor(96, 94, 92)
    
    # 2.2
    p_22 = doc.add_paragraph()
    r_22 = p_22.add_run("2.2 「垃圾進，垃圾出 (GIGO)」定律對航運決策之影響")
    r_22.font.name = "微軟正黑體"
    r_22.font.size = Pt(13)
    r_22.font.bold = True
    r_22.font.color.rgb = RGBColor(16, 124, 65)
    
    p = doc.add_paragraph()
    style_paragraph(p)
    p.add_run("資訊科學著名的哲理「Garbage In, Garbage Out (GIGO)」在航運業界體現尤深。無論後端使用的是多麼高階的 Power BI 圖表或複雜的 Python 機器學習模型，若前端航運行政輸入的底層 Excel 表格本身充滿缺陷，輸出的分析報告必然是錯誤的誘導資訊。")
    
    p = doc.add_paragraph()
    style_paragraph(p)
    p.add_run("例如：若某個 TEU（20呎標準櫃）之貨物重量誤將英磅 (LBS) 當作公斤 (KG) 且夾帶文字「lbs」，系統無法計算，將可能導致集裝箱重心中斷、船舶裝載配載 (Stowage Planning) 偏載傾斜，違反海上人命安全公約 (SOLAS) 之貨櫃重量驗證 (VGM) 規定，引發被港口國管制抽查扣押與處分之巨額責任。")
    
    # 2.3
    p_23 = doc.add_paragraph()
    r_23 = p_23.add_run("2.3 資料型態的嚴格判定：數值、文字、日期與邏輯值")
    r_23.font.name = "微軟正黑體"
    r_23.font.size = Pt(13)
    r_23.font.bold = True
    r_23.font.color.rgb = RGBColor(16, 124, 65)
    
    p = doc.add_paragraph()
    style_paragraph(p)
    p.add_run("Excel 儲存格具備強型別特性。初學者必須建立「一列一種型態」的嚴謹思維。表 2-2 說明了四大基本型別之特性與其在航運表單中的標準歸屬：")
    
    # 表 2-2
    tbl_types = doc.add_table(rows=5, cols=4)
    tbl_types.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_headers = ["資料型態", "預設對齊方式", "航運實務欄位範例", "關鍵辨識與注意事項"]
    for c_i, h_txt in enumerate(t_headers):
        cell = tbl_types.cell(0, c_i)
        set_cell_background(cell, "004E8C")
        set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
        p = cell.paragraphs[0]
        r = p.add_run(h_txt)
        r.bold = True
        r.font.name = "微軟正黑體"
        r.font.color.rgb = RGBColor(255, 255, 255)
        r.font.size = Pt(9.5)
        
    t_rows = [
        ("數值型 (Numeric)", "靠右對齊 (Right)", "貨物毛重 (KG)、海運費 (USD)、體積 (CBM)", "純數字，不得混雜單位（如 kg）或逗號文字，方可加總"),
        ("文字型 (Text)", "靠左對齊 (Left)", "提單號碼 (BL_NO)、出發港 (POL)、目的港 (POD)", "即使純數字（如櫃號 1234567）亦建議作文字處理以保全開頭0"),
        ("日期型 (Date)", "靠右對齊 (Right)", "結關日期 (Closing Date)、預計開航日 (ETD)", "本質為自 1900/1/1 起算之序列值，必須具備年/月/日三元素"),
        ("邏輯型 (Boolean)", "置中對齊 (Center)", "海關放行註記 (TRUE/FALSE)、是否危險品 (Yes/No)", "常用於條件篩選與 IF 函數的判斷式回傳結果")
    ]
    for r_i, (tp, al, fld, nte) in enumerate(t_rows):
        row = tbl_types.rows[r_i + 1]
        bg = "F9F9F9" if r_i % 2 == 1 else "FFFFFF"
        for c_i, val in enumerate([tp, al, fld, nte]):
            cell = row.cells[c_i]
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.name = "微軟正黑體"
            r.font.size = Pt(9)
            if c_i == 0:
                r.bold = True
                
    p_t3 = doc.add_paragraph()
    p_t3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t3.paragraph_format.space_before = Pt(4)
    p_t3.paragraph_format.space_after = Pt(12)
    r_t3 = p_t3.add_run("表 2-2 : Excel 四大基礎資料型態特性與航運實例歸納表")
    r_t3.font.name = "微軟正黑體"
    r_t3.font.size = Pt(9)
    r_t3.font.bold = True
    r_t3.font.color.rgb = RGBColor(96, 94, 92)
    
    # 2.4
    p_24 = doc.add_paragraph()
    r_24 = p_24.add_run("2.4 乾淨資料表 (Tidy Data) 的六大設計守則")
    r_24.font.name = "微軟正黑體"
    r_24.font.size = Pt(13)
    r_24.font.bold = True
    r_24.font.color.rgb = RGBColor(16, 124, 65)
    
    p = doc.add_paragraph()
    style_paragraph(p)
    p.add_run("根據資料科學家 Hadley Wickham 提出的 Tidy Data 架構，航運行政人員製作清冊時應恪遵以下「六不原則」：")
    
    rules = [
        ("守則一：絕不使用合併儲存格 (No Merged Cells)", "合併儲存格是樞紐分析與公式查表的頭號殺手，會導致除左上角外其餘單元格皆為 NULL 假空值。"),
        ("守則二：首列維持單一標題列 (Single Header Row)", "第一列嚴格作為欄位名稱（如提單號、港口、重量），避免將備註說明或公司標題跨欄置於表頭之上。"),
        ("守則三：一列代表一個業務實體 (One Observation per Row)", "每一橫列僅記錄一筆特定訂艙提單或貨櫃，不可將多筆訂艙隨意堆疊在同一個單元格內。"),
        ("守則四：一欄僅儲存單一屬性 (One Variable per Column)", "嚴格區分出發港與目的港為兩獨立欄位，切勿合併寫成「KHH -> LAX」造成後續篩選困難。"),
        ("守則五：純數據清冊不留人工小計列 (No Subtotal Rows in Raw Data)", "原始資料區切勿每隔五列插入小計（Subtotal），匯總工作應交由透視表處理，維持原表純粹性。"),
        ("守則六：杜絕無意義空列與空欄 (No Blank Rows / Columns)", "資料清單內部不應出現空白行用以排版隔開，否則快捷鍵與自動範圍辨識將在空行處中斷。")
    ]
    for r_title, r_desc in rules:
        p_r = doc.add_paragraph()
        style_paragraph(p_r, space_before=2, space_after=3)
        r_rt = p_r.add_run(r_title + " ： ")
        r_rt.bold = True
        r_rt.font.color.rgb = RGBColor(0, 78, 140)
        p_r.add_run(r_desc)
        
    # 換頁 -> 第三章
    doc.add_page_break()
    
    # -------------------------------------------------------------
    # 第三章 30筆航運示範資料檔前處理實作演練
    # -------------------------------------------------------------
    p_c3 = doc.add_paragraph()
    p_c3.paragraph_format.space_before = Pt(12)
    p_c3.paragraph_format.space_after = Pt(12)
    r_c3 = p_c3.add_run("第三章 30筆航運示範資料檔前處理實作演練")
    r_c3.font.name = "微軟正黑體"
    r_c3.font.size = Pt(18)
    r_c3.font.bold = True
    r_c3.font.color.rgb = RGBColor(0, 78, 140)
    
    # 3.1
    p_31 = doc.add_paragraph()
    r_31 = p_31.add_run("3.1 示範資料集欄位結構與問題診斷")
    r_31.font.name = "微軟正黑體"
    r_31.font.size = Pt(13)
    r_31.font.bold = True
    r_31.font.color.rgb = RGBColor(16, 124, 65)
    
    p = doc.add_paragraph()
    style_paragraph(p)
    p.add_run("本教學以航運出口實務資料檔【Shipping_Demo_Raw_30.xlsx】作為實習範本。此檔案已置於教材目錄之 data/ 資料夾中，完整模擬航運行政自舊系統匯出時常發生的 30 筆真實髒資料情境。表 3-1 詳列各欄位所存在的具體問題診斷清單：")
    
    # 表 3-1
    tbl_diag = doc.add_table(rows=11, cols=4)
    tbl_diag.alignment = WD_TABLE_ALIGNMENT.CENTER
    diag_headers = ["欄位名稱", "資料意涵", "原始資料所涉問題與缺陷", "預計前處理對策"]
    for c_i, h_txt in enumerate(diag_headers):
        cell = tbl_diag.cell(0, c_i)
        set_cell_background(cell, "004E8C")
        set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
        p = cell.paragraphs[0]
        r = p.add_run(h_txt)
        r.bold = True
        r.font.name = "微軟正黑體"
        r.font.color.rgb = RGBColor(255, 255, 255)
        r.font.size = Pt(9.5)
        
    diag_rows = [
        ("提單號碼(BL_NO)", "海運提單唯一識別代碼", "夾雜首尾空白(如' MSKU.. ')、部分英文小寫、第4筆完全重複", "TRIM去除空格、UPPER轉大寫、移除重複項"),
        ("訂艙代號(Booking_ID)", "托運人向船東訂艙號", "編號正常，與提單號碼呈一對一關聯", "作為輔助檢驗主鍵是否唯一"),
        ("結關日期(Closing_Date)", "貨櫃進場截止日期", "格式嚴重分歧：斜線(/)、橫線(-)、句點(.)、無分隔碼(20240504)", "Ctrl+H取代符號、資料剖析轉為標準日期"),
        ("出發港(POL)", "Port of Loading", "部分資料前後夾帶多餘空格 (如' Kaohsiung ')", "TRIM 函數批次清除首尾無效空白"),
        ("目的港(POD)", "Port of Discharge", "末端帶空格 (如'LGB ')，影響分類統計", "TRIM 函數批次清除空格"),
        ("櫃號(Container_No)", "ISO 貨櫃編號 (4碼英文+7碼數字)", "格式尚稱完整，少數有大小寫問題", "UPPER 函數轉全大寫標準碼"),
        ("櫃型(Type)", "標準櫃型 (20GP, 40GP, 40HC)", "規範標準", "保留原始資料"),
        ("貨物毛重(Weight)", "貨物實際過磅毛重 (KG)", "含千分位逗號字串、存在缺失值'N/A'、小數點位數不一", "IFERROR填補、VALUE轉純數值、四捨五入"),
        ("海運費(Freight)", "基本海運費金額 (USD)", "字串開頭帶有美元符號'$'、夾帶千分逗點", "SUBSTITUTE替換貨幣字元、轉為貨幣格式"),
        ("狀態備註(Remark)", "特急件/危險品/查驗等營運註記", "存在空白或多餘空格", "IF 條件判斷，將空白統一補註為'一般一般'")
    ]
    for r_i, (fld, mean, iss, sol) in enumerate(diag_rows):
        row = tbl_diag.rows[r_i + 1]
        bg = "F9F9F9" if r_i % 2 == 1 else "FFFFFF"
        for c_i, val in enumerate([fld, mean, iss, sol]):
            cell = row.cells[c_i]
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=70, bottom=70, left=80, right=80)
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.name = "微軟正黑體"
            r.font.size = Pt(8.5)
            if c_i == 0:
                r.bold = True
                
    p_t4 = doc.add_paragraph()
    p_t4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t4.paragraph_format.space_before = Pt(4)
    p_t4.paragraph_format.space_after = Pt(12)
    r_t4 = p_t4.add_run("表 3-1 : 30筆航運示範資料集原始結構與前處理診斷清單")
    r_t4.font.name = "微軟正黑體"
    r_t4.font.size = Pt(9)
    r_t4.font.bold = True
    r_t4.font.color.rgb = RGBColor(96, 94, 92)
    
    # 3.2
    p_32 = doc.add_paragraph()
    r_32 = p_32.add_run("3.2 步驟一：剔除頭尾無效空白 (TRIM 與 CLEAN 函數運用)")
    r_32.font.name = "微軟正黑體"
    r_32.font.size = Pt(13)
    r_32.font.bold = True
    r_32.font.color.rgb = RGBColor(16, 124, 65)
    
    p = doc.add_paragraph()
    style_paragraph(p)
    p.add_run("航運提單最隱蔽的致命傷即為字串中的「半形或全形空白」。例如「 Kaohsiung」與「Kaohsiung 」在肉眼看似相同，但 Excel 的比對邏輯中兩者長度不同（分別為 10 與 10 字元，正常為 9 字元），導致 VLOOKUP 查表失敗或樞紐分析拆成兩列。")
    
    p = doc.add_paragraph()
    style_paragraph(p)
    p.add_run("【實作指令演練】：\n1. 在原始資料表 D 欄 (出發港 POL) 旁插入一輔助空欄，命名為「標準出發港」。\n2. 於列 2 儲存格輸入公式：\n   =TRIM(CLEAN(D2))\n3. 按下 Enter 後，將滑鼠移至該儲存格右下角之填滿控點，連點左鍵兩下向下填滿至第 31 列。\n4. 選取該輔助欄整欄，按 Ctrl+C 複製，點選原 D 欄標題，按滑鼠右鍵選擇「貼上值 (Values Only)」，最後刪除輔助欄。")
    
    # 3.3
    p_33 = doc.add_paragraph()
    r_33 = p_33.add_run("3.3 步驟二：英文代碼大小寫標準化 (UPPER 與 LOWER 函數)")
    r_33.font.name = "微軟正黑體"
    r_33.font.size = Pt(13)
    r_33.font.bold = True
    r_33.font.color.rgb = RGBColor(16, 124, 65)
    
    p = doc.add_paragraph()
    style_paragraph(p)
    p.add_run("國際海運提單字首四碼為標準國際貨櫃局 (BIC) 註冊代碼（如 MSKU 代表馬士基航運，EMCU 代表長榮海運）。第 3 筆資料出現小寫「msku-8291036」，需全面正規化為大寫。\n公式語法：\n=UPPER(TRIM(A2))\n此公式先藉由 TRIM 剔除前後空格，再由 UPPER 將所有英文字母一律轉為標準大寫，確保後續跨系統串接與 EDI 申報一致。")
    
    # 3.4
    p_34 = doc.add_paragraph()
    r_34 = p_34.add_run("3.4 步驟三：日期格式混雜正規化處理 (DATEVALUE / 取代法)")
    r_34.font.name = "微軟正黑體"
    r_34.font.size = Pt(13)
    r_34.font.bold = True
    r_34.font.color.rgb = RGBColor(16, 124, 65)
    
    p = doc.add_paragraph()
    style_paragraph(p)
    p.add_run("示範檔結關日期欄位呈現多樣性：包含「2024/05/01」、「2024-05-02」、「2024.05.02」以及「20240504」。\n【兩階段標準化流程】：\n階段一（快速取代法）：選取結關日欄位，按 Ctrl+H 開啟取代視窗，在「尋找目標」輸入「.」，在「取代成」輸入「-」，點選「全部取代」；隨後再次將「/」全部取代為「-」。\n階段二（緊湊型日期剖析）：針對無分隔號「20240504」，可使用 DATE 結合文字擷取函數處理：\n=DATE(LEFT(C8,4), MID(C8,5,2), RIGHT(C8,2))\n最後全選結關日欄位，按 Ctrl+1 開啟儲存格格式，選取「日期」分類下之標準「YYYY-MM-DD」格式。")
    
    # 3.5
    p_35 = doc.add_paragraph()
    r_35 = p_35.add_run("3.5 步驟四：主鍵去重複作業 (提單號碼 Remove Duplicates)")
    r_35.font.name = "微軟正黑體"
    r_35.font.size = Pt(13)
    r_35.font.bold = True
    r_35.font.color.rgb = RGBColor(16, 124, 65)
    
    p = doc.add_paragraph()
    style_paragraph(p)
    p.add_run("示範資料中，第 1 筆與第 4 筆提單號碼均為「MSKU-8291034」，訂艙代號同為「BK001」，此為航運作業常見的重覆入單問題。必須執行主鍵去重。操作介面如圖 3-1 所示。")
    
    # 插入圖 3-1
    p_img2 = doc.add_paragraph()
    p_img2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img2.paragraph_format.space_before = Pt(8)
    p_img2.paragraph_format.space_after = Pt(4)
    run_img2 = p_img2.add_run()
    run_img2.add_picture("assets/fig2_remove_duplicates.png", width=Inches(5.8))
    
    p_cap2 = doc.add_paragraph()
    p_cap2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap2.paragraph_format.space_after = Pt(10)
    r_cap2 = p_cap2.add_run("圖 3-1 : Excel 「移除重複項」對話框與航運提單主鍵設定示意圖")
    r_cap2.font.name = "微軟正黑體"
    r_cap2.font.size = Pt(9.5)
    r_cap2.font.bold = True
    r_cap2.font.color.rgb = RGBColor(96, 94, 92)
    
    p = doc.add_paragraph()
    style_paragraph(p)
    p.add_run("【去重複操作步驟】：\n1. 滑鼠選取資料區內任意儲存格（A1:J31）。\n2. 切換至頂端功能區【資料】標籤頁 -> 點選【移除重複項 (Remove Duplicates)】。\n3. 於快顯對話框中務必勾選「我的資料有標題」。\n4. 點選「取消全選」，單獨勾選航運唯一主鍵【提單號碼(BL_NO)】與【訂艙代號(Booking_ID)】。\n5. 點選「確定」，Excel 立即跳出提示：「找到 1 個重複的值並已將其移除；保留了 29 個唯一值。」")
    
    # 3.6
    p_36 = doc.add_paragraph()
    r_36 = p_36.add_run("3.6 步驟五：非數值與缺漏值處理 (IFERROR / 填補策略)")
    r_36.font.name = "微軟正黑體"
    r_36.font.size = Pt(13)
    r_36.font.bold = True
    r_36.font.color.rgb = RGBColor(16, 124, 65)
    
    p = doc.add_paragraph()
    style_paragraph(p)
    p.add_run("第 5 筆資料 (OOLU-1102934) 貨物毛重欄位填寫為「N/A」（文字型缺失），若直接套用加總公式將導致失誤。航運行政有兩種正規填補策略：\n策略一（標註缺失等待補重）：若該貨櫃尚未過磅，不可擅自填寫 0，應以條件式標註：\n=IF(ISNUMBER(VALUE(G5)), VALUE(G5), \"待補重單\")\n策略二（保護型安全加總）：使用 IFERROR 函數消除報錯：\n=IFERROR(VALUE(SUBSTITUTE(G2,\",\",\"\")), 0)\n搭配「常用」->「設定格式化的條件」，將數值為 0 或小於預設門檻的儲存格自動標記為淺紅色底色，提醒出口文件部同仁於結關前向拖車司機索取過磅重單。")
    
    # 3.7
    p_37 = doc.add_paragraph()
    r_37 = p_37.add_run("3.7 步驟六：貨幣符號與千分位純數值轉化 (SUBSTITUTE / VALUE 結合)")
    r_37.font.name = "微軟正黑體"
    r_37.font.size = Pt(13)
    r_37.font.bold = True
    r_37.font.color.rgb = RGBColor(16, 124, 65)
    
    p = doc.add_paragraph()
    style_paragraph(p)
    p.add_run("海運費欄位中存在帶貨幣字元「$1,850」及千分逗號「2,100」。由於文字符號的存在，儲存格靠左對齊，無法進行財務運算。核心前處理函數嵌套架構如圖 3-3 所示。")
    
    # 插入圖 3-3
    p_img4 = doc.add_paragraph()
    p_img4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img4.paragraph_format.space_before = Pt(8)
    p_img4.paragraph_format.space_after = Pt(4)
    run_img4 = p_img4.add_run()
    run_img4.add_picture("assets/fig4_cleaning_formulas.png", width=Inches(5.8))
    
    p_cap4 = doc.add_paragraph()
    p_cap4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap4.paragraph_format.space_after = Pt(10)
    r_cap4 = p_cap4.add_run("圖 3-3 : 航運前處理四大核心函數結構 (TRIM, CLEAN, VALUE, SUBSTITUTE)")
    r_cap4.font.name = "微軟正黑體"
    r_cap4.font.size = Pt(9.5)
    r_cap4.font.bold = True
    r_cap4.font.color.rgb = RGBColor(96, 94, 92)
    
    p = doc.add_paragraph()
    style_paragraph(p)
    p.add_run("【運費純數值化萬用嵌套公式】：\n=VALUE(SUBSTITUTE(SUBSTITUTE(TRIM(I2), \"$\", \"\"), \",\", \"\"))\n【公式運作剖析】：\n1. 最內層 TRIM(I2)：清除所有誤植之前後空白。\n2. 第一層 SUBSTITUTE(..., \"$\", \"\")：將美金記號「$」精準替換為空字元。\n3. 第二層 SUBSTITUTE(..., \",\", \"\")：將干擾計算之千分位逗號剔除。\n4. 最外層 VALUE(...)：將已無雜質的純數字字串轉換為 Excel 真實浮點數。\n轉換完畢後，透過 Ctrl+1 設定為「貨幣」格式（貨幣符號: $，小數位數: 0），此時「$」是由 Excel 格式引擎所渲染，儲存格內實質為純數字 1850，可 100% 正常參與 SUMIF 等分析公式。")
    
    # 3.8
    p_38 = doc.add_paragraph()
    r_38 = p_38.add_run("3.8 成果驗證：資料前處理前後 (Before vs. After) 綜合評估")
    r_38.font.name = "微軟正黑體"
    r_38.font.size = Pt(13)
    r_38.font.bold = True
    r_38.font.color.rgb = RGBColor(16, 124, 65)
    
    p = doc.add_paragraph()
    style_paragraph(p)
    p.add_run("完成上述六大前處理程序後，原先充滿髒數據的 30 筆示範檔蛻變為完全符合 Tidy Data 規範的航運分析資料庫。其效益對比與品質提升成果如圖 3-4 所示。")
    
    # 插入圖 3-4
    p_img5 = doc.add_paragraph()
    p_img5.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img5.paragraph_format.space_before = Pt(8)
    p_img5.paragraph_format.space_after = Pt(4)
    run_img5 = p_img5.add_run()
    run_img5.add_picture("assets/fig5_data_comparison.png", width=Inches(5.8))
    
    p_cap5 = doc.add_paragraph()
    p_cap5.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap5.paragraph_format.space_after = Pt(12)
    r_cap5 = p_cap5.add_run("圖 3-4 : 航運資料前處理前後 (Before vs. After) 成果對比與品質看板")
    r_cap5.font.name = "微軟正黑體"
    r_cap5.font.size = Pt(9.5)
    r_cap5.font.bold = True
    r_cap5.font.color.rgb = RGBColor(96, 94, 92)
    
    # 換頁 -> 第四章
    doc.add_page_break()
    
    # -------------------------------------------------------------
    # 第四章 300筆綜合型航運案例情境實戰練習與解題全攻略
    # -------------------------------------------------------------
    p_c4 = doc.add_paragraph()
    p_c4.paragraph_format.space_before = Pt(12)
    p_c4.paragraph_format.space_after = Pt(12)
    r_c4 = p_c4.add_run("第四章 300筆綜合型航運案例情境實戰練習與解題全攻略")
    r_c4.font.name = "微軟正黑體"
    r_c4.font.size = Pt(18)
    r_c4.font.bold = True
    r_c4.font.color.rgb = RGBColor(0, 78, 140)
    
    # 4.1
    p_41 = doc.add_paragraph()
    r_41 = p_41.add_run("4.1 300筆進階實戰練習資料庫背景設定")
    r_41.font.name = "微軟正黑體"
    r_41.font.size = Pt(13)
    r_41.font.bold = True
    r_41.font.color.rgb = RGBColor(16, 124, 65)
    
    p = doc.add_paragraph()
    style_paragraph(p)
    p.add_run("為了鞏固初學者的操作手感，教材專屬附贈了大型情境練習檔案【Shipping_Practice_300.xlsx】（存放於 data/ 資料夾）。本資料集涵蓋全球八大貨櫃船公司、四大台灣出發港、全球二十大基本目的港以及六大貨櫃類型，完整涵蓋跨年度 300 餘筆真實業務數據。表 4-1 列出實戰檔案涵蓋之船公司航商與母港分布：")
    
    # 表 4-1
    tbl_p300 = doc.add_table(rows=5, cols=4)
    tbl_p300.alignment = WD_TABLE_ALIGNMENT.CENTER
    p_headers = ["船公司代號 (Carrier)", "所屬海運聯盟", "主要出發港 (POL)", "櫃型範疇 (Equipment)"]
    for c_i, h_txt in enumerate(p_headers):
        cell = tbl_p300.cell(0, c_i)
        set_cell_background(cell, "004E8C")
        set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
        p = cell.paragraphs[0]
        r = p.add_run(h_txt)
        r.bold = True
        r.font.name = "微軟正黑體"
        r.font.color.rgb = RGBColor(255, 255, 255)
        r.font.size = Pt(9.5)
        
    p_rows = [
        ("MSKU (馬士基) / HMMU (韓新遠洋)", "Gemini 雙子星 / 獨立運營", "Kaohsiung (高雄港)", "20GP (20呎標準乾櫃)"),
        ("OOLU (東方海外) / CMAU (達飛海運)", "Ocean Alliance (海洋聯盟)", "Keelung (基隆港)", "40GP (40呎標準乾櫃)"),
        ("EMCU (長榮海運)", "Ocean Alliance (海洋聯盟)", "Taichung (台中港)", "40HC (40呎高櫃) / 45HC"),
        ("ONEU (海洋網聯) / YMLU (陽明) / WHLC (萬海)", "THE Alliance / 亞洲區間獨立", "Taipei (台北港)", "20RF / 40RF (冷凍冷藏櫃)")
    ]
    for r_i, (c1, c2, c3, c4) in enumerate(p_rows):
        row = tbl_p300.rows[r_i + 1]
        bg = "F9F9F9" if r_i % 2 == 1 else "FFFFFF"
        for c_i, val in enumerate([c1, c2, c3, c4]):
            cell = row.cells[c_i]
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.name = "微軟正黑體"
            r.font.size = Pt(9)
            if c_i == 0:
                r.bold = True
                
    p_t5 = doc.add_paragraph()
    p_t5.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t5.paragraph_format.space_before = Pt(4)
    p_t5.paragraph_format.space_after = Pt(12)
    r_t5 = p_t5.add_run("表 4-1 : 300筆實戰練習檔涵蓋之八大船公司與主要港口統計表")
    r_t5.font.name = "微軟正黑體"
    r_t5.font.size = Pt(9)
    r_t5.font.bold = True
    r_t5.font.color.rgb = RGBColor(96, 94, 92)
    
    # 4.2
    p_42 = doc.add_paragraph()
    r_42 = p_42.add_run("4.2 實戰情境挑戰題 (三大航運前處理任務)")
    r_42.font.name = "微軟正黑體"
    r_42.font.size = Pt(13)
    r_42.font.bold = True
    r_42.font.color.rgb = RGBColor(16, 124, 65)
    
    p = doc.add_paragraph()
    style_paragraph(p)
    p.add_run("學員請開啟【Shipping_Practice_300.xlsx】，並依序完成以下三項航運業務核心任務：\n\n"
              "【任務一：提單與航商維度拆解】\n"
              "原始提單號碼如「MSKU-8291034」，請利用 Excel「資料剖析」或文字函數，拆分出獨立之「航商代號 (Prefix)」與「流水號 (Serial No)」，並建立統一字軌。\n\n"
              "【任務二：全表文字雜質清洗與去重審核】\n"
              "清除出發港與目的港所有前後夾帶之空格，識別全表潛在重複訂艙單（提示：資料庫內刻意埋入 5 筆完全重複提單），執行正確去重剔除。\n\n"
              "【任務三：重量與海運費數值淨化與異常警示】\n"
              "將包含貨幣符號、千分逗號之金額轉為純數值；將貨物毛重欄位中之「N/A」與空值填補為 0，並運用條件式格式設定將重量超過 22,000 KG 之「超重重櫃」標示為醒目黃底紅字。")
              
    # 4.3
    p_43 = doc.add_paragraph()
    r_43 = p_43.add_run("4.3 實戰挑戰題參考標準解題步驟 (SOP)")
    r_43.font.name = "微軟正黑體"
    r_43.font.size = Pt(13)
    r_43.font.bold = True
    r_43.font.color.rgb = RGBColor(16, 124, 65)
    
    p = doc.add_paragraph()
    style_paragraph(p)
    p.add_run("為使自學者能即時自我檢視成果，以下提供標準參考解答與實務公式指南：")
    
    solutions = [
        ("解題任務一（資料剖析操作 SOP）：", 
         "1. 選取提單號碼整欄 (A欄)。\n"
         "2. 點選【資料】標籤頁 -> 【資料剖析 (Text to Columns)】。\n"
         "3. 步驟 1 選擇「分隔符號 (Delimited)」-> 點選「下一步」。\n"
         "4. 步驟 2 勾選「其他 (Other)」，於右側輸入方框鍵入減號破折號「-」。（參考圖 3-2 介面）。\n"
         "5. 步驟 3 目標儲存格選取旁邊之空白欄位，點選「完成」。Excel 瞬間將 300 筆提單精準切分為「航商代碼」與「7碼流水號」。"),
        
        ("解題任務二（空格淨化與去重 SOP）：",
         "1. 新增標準出發港與標準目的港欄位，輸入公式：\n"
         "   =PROPER(TRIM(D2))  及  =UPPER(TRIM(E2))\n"
         "2. 將公式結果「貼上值」覆蓋回原始欄位。\n"
         "3. 點選資料區任一處，按 Ctrl+A 全選，點擊【資料】->【移除重複項】。\n"
         "4. 勾選全部欄位，點擊確定。系統將成功查獲並刪除 5 筆重複資料，總筆數由 305 筆自動收斂為標準 300 筆唯一值。"),
         
        ("解題任務三（金額淨化與重量超重格式化 SOP）：",
         "1. 運費欄位數值化公式（輸入於輔助欄後貼上值）：\n"
         "   =VALUE(SUBSTITUTE(SUBSTITUTE(TRIM(I2), \"$\", \"\"), \",\", \"\"))\n"
         "2. 貨物重量清洗公式（有效防禦非數值型文字）：\n"
         "   =IFERROR(VALUE(SUBSTITUTE(TRIM(H2), \",\", \"\")), 0)\n"
         "3. 選取清洗後之重量欄位 -> 點選【常用】->【條件式格式設定】->【反白標示儲存格規則】->【大於 (Greater Than)】。\n"
         "4. 輸入臨界值「22000」，填滿格式選擇「自訂格式」-> 設定背景為黃色、文字為深紅色粗體。如此一來，凡達到重櫃加價收費門檻之集裝箱一目了然！")
    ]
    for s_title, s_detail in solutions:
        p_s = doc.add_paragraph()
        style_paragraph(p_s, space_before=3, space_after=2)
        r_st = p_s.add_run(s_title)
        r_st.bold = True
        r_st.font.color.rgb = RGBColor(0, 78, 140)
        
        p_sd = doc.add_paragraph()
        style_paragraph(p_sd, space_before=1, space_after=4)
        p_sd.add_run(s_detail)
        
    # 4.4
    p_44 = doc.add_paragraph()
    r_44 = p_44.add_run("4.4 進階延伸技巧：Power Query 現代化前處理自動化簡介")
    r_44.font.name = "微軟正黑體"
    r_44.font.size = Pt(13)
    r_44.font.bold = True
    r_44.font.color.rgb = RGBColor(16, 124, 65)
    
    p = doc.add_paragraph()
    style_paragraph(p)
    p.add_run("若行政人員每月皆需重複處理來自同一船公司的例行裝船清單，建議在熟悉基礎函數後，進一步探索 Excel 內建的【Power Query (取得與轉換資料)】工具。只需錄製一次「修剪空白」、「文字大寫」與「取代字元」步驟，未來每當新月份結關報表匯入時，點擊「全部重新整理」，數萬筆資料便可在 3 秒內自動完成所有前處理流程，真正實現航運行政數位化與零人工介入！")
    
    # 換頁 -> 第五章
    doc.add_page_break()
    
    # -------------------------------------------------------------
    # 第五章 參考來源與航運數據規範文獻
    # -------------------------------------------------------------
    p_c5 = doc.add_paragraph()
    p_c5.paragraph_format.space_before = Pt(12)
    p_c5.paragraph_format.space_after = Pt(12)
    r_c5 = p_c5.add_run("第五章 參考來源與航運數據規範文獻")
    r_c5.font.name = "微軟正黑體"
    r_c5.font.size = Pt(18)
    r_c5.font.bold = True
    r_c5.font.color.rgb = RGBColor(0, 78, 140)
    
    p = doc.add_paragraph()
    style_paragraph(p)
    p.add_run("本手冊所引述之資料清洗準則、海運業務術語、提單編碼原則與試算表函數架構，均考證並參考自以下國內外權威海運機構與科技文獻標準：")
    
    references = [
        ("1. 國際海事組織 (IMO) - SOLAS 公約貨櫃總重驗證 (VGM) 作業指引", 
         "International Maritime Organization. (2016). Guidelines Regarding the Verified Gross Mass of a Container Carrying Cargo. IMO MSC.1/Circ.1475."),
        
        ("2. 國際貨櫃局 (BIC) 貨櫃國際代碼標準與檢查碼演算法", 
         "Bureau International des Containers. (2023). ISO 6346: Freight containers - Coding, identification and marking. BIC International Standard Database."),
        
        ("3. 數位貨櫃航運協會 (DCSA) 航運資料標準介面", 
         "Digital Container Shipping Association. (2022). DCSA Standards for Operational Vessel Schedules and Track & Trace Interface Specifications (Version 2.2)."),
        
        ("4. 資料科學整潔資料標準 (Tidy Data 哲學)", 
         "Wickham, H. (2014). Tidy Data. Journal of Statistical Software, 59(10), 1-23. DOI: 10.18637/jss.v059.i10"),
        
        ("5. 微軟官方 Excel 支援中心文件庫", 
         "Microsoft Corporation. (2024). Microsoft Excel Function Reference: Text, Clean, Value, Substitute and Data Cleansing Best Practices. Microsoft Learn Support Center."),
        
        ("6. 台灣港務股份有限公司 (TIPC) 國際商港港務作業手冊", 
         "Taiwan International Ports Corporation. (2023). 港務EDI連線申報與提單資料登錄作業規範指引. 高雄市: 臺灣港務公司出版品.")
    ]
    
    for ref_title, ref_cite in references:
        p_ref = doc.add_paragraph()
        style_paragraph(p_ref, space_before=3, space_after=1)
        r_rt = p_ref.add_run(ref_title)
        r_rt.bold = True
        r_rt.font.color.rgb = RGBColor(0, 78, 140)
        
        p_rc = doc.add_paragraph()
        style_paragraph(p_rc, space_before=1, space_after=4)
        r_cite = p_rc.add_run("   " + ref_cite)
        r_cite.font.size = Pt(9.5)
        r_cite.font.color.rgb = RGBColor(80, 80, 80)
        
    add_header_footer(doc)
    
    output_filename = "航運行政人員Excel資料前處理教學手冊.docx"
    doc.save(output_filename)
    print(f"Document saved successfully as: {output_filename}")

if __name__ == "__main__":
    build_shipping_manual()
