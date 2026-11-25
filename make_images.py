import os
from PIL import Image, ImageDraw, ImageFont

os.makedirs("assets", exist_ok=True)

# 嘗試尋找合適字型 (微軟正黑體或 Arial)
def get_fonts():
    font_paths = [
        "C:\\Windows\\Fonts\\msjh.ttc",
        "C:\\Windows\\Fonts\\msjhbd.ttc",
        "C:\\Windows\\Fonts\\arial.ttf",
        "C:\\Windows\\Fonts\\seguiemj.ttf"
    ]
    font_main = None
    font_bold = None
    font_sm = None
    font_lg = None
    try:
        font_main = ImageFont.truetype("C:\\Windows\\Fonts\\msjh.ttc", 16)
        font_bold = ImageFont.truetype("C:\\Windows\\Fonts\\msjhbd.ttc", 18)
        font_sm = ImageFont.truetype("C:\\Windows\\Fonts\\msjh.ttc", 13)
        font_lg = ImageFont.truetype("C:\\Windows\\Fonts\\msjhbd.ttc", 22)
    except Exception as e:
        print("Fallback font:", e)
        font_main = ImageFont.load_default()
        font_bold = font_main
        font_sm = font_main
        font_lg = font_main
    return font_main, font_bold, font_sm, font_lg

f_main, f_bold, f_sm, f_lg = get_fonts()

# 截圖 1: Excel 介面解析 (快速存取列, 功能區 Ribbon, 名稱方塊, 資料編輯列, 工作表網格, 狀態列)
def make_ui_interface_img():
    w, h = 1000, 560
    im = Image.new("RGB", (w, h), "#F3F2F1")
    d = ImageDraw.Draw(im)
    
    # 頂部綠色標題列 (Excel 品牌色 #107C41)
    d.rectangle([0, 0, w, 40], fill="#107C41")
    d.text((20, 10), "Autosave   |   Shipping_Demo_Raw_30.xlsx - Excel", fill="white", font=f_bold)
    
    # 功能區標籤頁 (Ribbon Tabs)
    d.rectangle([0, 40, w, 75], fill="#FFFFFF")
    tabs = ["檔案 (File)", "常用 (Home)", "插入 (Insert)", "頁面配置 (Page Layout)", "公式 (Formulas)", "資料 (Data)", "校閱 (Review)", "檢視 (View)"]
    tx = 20
    for idx, tab in enumerate(tabs):
        if idx == 5: # 資料標籤突出
            d.rectangle([tx-8, 45, tx+100, 75], fill="#E1DFDD")
            d.text((tx, 50), tab, fill="#107C41", font=f_bold)
            tx += 115
        else:
            d.text((tx, 50), tab, fill="#323130", font=f_main)
            tx += 105
            
    # 功能區工具欄 (Ribbon Panel)
    d.rectangle([0, 75, w, 150], fill="#F3F2F1")
    d.line([0, 150, w, 150], fill="#D2D0CE", width=1)
    
    # 劃分功能區塊群組
    d.rectangle([20, 82, 220, 142], outline="#C8C6C4", fill="#FFFFFF")
    d.text((30, 88), "取得與轉換資料 (Get Data)", fill="#605E5C", font=f_sm)
    d.text((30, 110), "[從文字/CSV]  [從Web]", fill="#201F1E", font=f_main)
    
    d.rectangle([230, 82, 450, 142], outline="#C8C6C4", fill="#FFFFFF")
    d.text((240, 88), "排序與篩選 (Sort & Filter)", fill="#605E5C", font=f_sm)
    d.text((240, 110), "[ A->Z 遞增 ]  [ 篩選器 Filter ]", fill="#201F1E", font=f_main)
    
    d.rectangle([460, 82, 750, 142], outline="#C8C6C4", fill="#FFFFFF")
    d.text((470, 88), "資料工具 (Data Tools)", fill="#605E5C", font=f_sm)
    d.text((470, 110), "[資料剖析]  [移除重複項]  [資料驗證]", fill="#107C41", font=f_bold)
    
    # 名稱方塊 & 資料編輯列 (Formula Bar)
    d.rectangle([0, 152, w, 190], fill="#FFFFFF")
    d.line([0, 190, w, 190], fill="#D2D0CE", width=1)
    
    # 名稱方塊
    d.rectangle([10, 158, 100, 184], outline="#8A8886", fill="#FFFFFF")
    d.text((30, 162), "A1", fill="#201F1E", font=f_bold)
    
    d.text((115, 162), "fx", fill="#605E5C", font=f_bold)
    
    # 編輯列輸入框
    d.rectangle([140, 158, w-20, 184], outline="#8A8886", fill="#FFFFFF")
    d.text((150, 162), "=TRIM(CLEAN(SUBSTITUTE(A2, \"-\", \"\")))", fill="#004E8C", font=f_main)
    
    # 工作表儲存格區域 (Grid)
    d.rectangle([0, 192, w, 510], fill="#FFFFFF")
    # 欄標題列 A, B, C, D...
    d.rectangle([40, 192, w, 220], fill="#F3F2F1")
    col_w = [140, 110, 120, 120, 110, 120, 110, 120]
    col_names = ["A (提單號碼)", "B (訂艙代號)", "C (結關日)", "D (POL)", "E (POD)", "F (櫃號)", "G (毛重KG)", "H (運費USD)"]
    cx = 40
    for idx, (cw, cn) in enumerate(zip(col_w, col_names)):
        d.rectangle([cx, 192, cx+cw, 220], outline="#D2D0CE")
        d.text((cx+10, 198), cn, fill="#323130", font=f_sm)
        cx += cw
        
    # 列號與資料列
    demo_rows = [
        ["  MSKU-8291034 ", "BK001", "2024/05/01", " Kaohsiung ", "LAX", "TGHU1234567", " 18,500 ", "$1,850"],
        ["MSKU-8291035", "BK002", "2024-05-02", "Keelung", "LGB ", "MSKU9876543", "8200", "1200"],
        ["msku-8291036", "BK003", "2024.05.02", "Taichung", "RTM", "TEMU5544332", "22400.5", "2,100"],
        ["MSKU-8291034", "BK001", "2024/05/01", "Kaohsiung", "LAX", "TGHU1234567", "18500", "1850"],
        ["OOLU-1102934", "BK004", "2024/05/03", "Kaohsiung", "HAM", "OOLU3322110", "N/A", "2300"]
    ]
    
    ry = 220
    for r_idx, r_data in enumerate(demo_rows):
        d.rectangle([0, ry, 40, ry+32], fill="#F3F2F1", outline="#D2D0CE")
        d.text((15, ry+6), str(r_idx+1), fill="#605E5C", font=f_sm)
        
        cx = 40
        for c_idx, (cw, val) in enumerate(zip(col_w, r_data)):
            bg = "#FFF0F0" if r_idx == 3 else ("#F9F9F9" if r_idx%2==1 else "#FFFFFF")
            d.rectangle([cx, ry, cx+cw, ry+32], fill=bg, outline="#E1DFDD")
            d.text((cx+8, ry+6), val, fill="#D13438" if r_idx==3 else "#201F1E", font=f_sm)
            cx += cw
        ry += 32
        
    # 底部狀態列
    d.rectangle([0, 520, w, h], fill="#107C41")
    d.text((20, 530), "就緒 (Ready)   |   工作表 1/1   |   選取計數: 5   |   平均: 15,720   |   合計: 62,880", fill="white", font=f_sm)
    
    # 標籤引導箭頭與說明文字方塊 (Callouts)
    # 標籤1: 功能區 (Ribbon)
    d.rectangle([760, 85, 980, 140], fill="#FFF4CE", outline="#FFB900", width=2)
    d.text((770, 92), "【1. 功能區 Ribbon】", fill="#323130", font=f_bold)
    d.text((770, 115), "核心操作命令集中處", fill="#605E5C", font=f_sm)
    
    # 標籤2: 編輯列 (Formula Bar)
    d.rectangle([680, 155, 980, 186], fill="#FFF4CE", outline="#FFB900", width=2)
    d.text((690, 160), "【2. 名稱方塊與公式編輯列】", fill="#323130", font=f_bold)
    
    # 標籤3: 資料列
    d.rectangle([620, 310, 980, 365], fill="#FFF4CE", outline="#FFB900", width=2)
    d.text((630, 318), "【3. 工作表儲存格 (Grid)】", fill="#323130", font=f_bold)
    d.text((630, 340), "航運業務原始資料登記區", fill="#605E5C", font=f_sm)
    
    im.save("assets/fig1_excel_interface.png")
    print("Saved assets/fig1_excel_interface.png")

# 截圖 2: 資料去重複功能對話框 (Remove Duplicates Dialog)
def make_remove_duplicates_img():
    w, h = 800, 480
    im = Image.new("RGB", (w, h), "#FFFFFF")
    d = ImageDraw.Draw(im)
    
    # 模擬 Windows Excel 對話框視窗
    d.rectangle([50, 40, 750, 440], fill="#FFFFFF", outline="#0078D4", width=2)
    d.rectangle([52, 42, 748, 75], fill="#F3F2F1")
    d.text((70, 50), "移除重複項 (Remove Duplicates)", fill="#201F1E", font=f_bold)
    d.text((725, 48), "X", fill="#605E5C", font=f_bold)
    
    d.text((75, 90), "若要刪除重複的值，請選取包含重複資料的一個或多個欄。", fill="#323130", font=f_main)
    
    # 全選 / 取消全選按鈕
    d.rectangle([75, 125, 185, 155], outline="#8A8886", fill="#F3F2F1")
    d.text((95, 132), "全選 (Select All)", fill="#201F1E", font=f_sm)
    d.rectangle([195, 125, 335, 155], outline="#8A8886", fill="#F3F2F1")
    d.text((210, 132), "取消全選 (Unselect)", fill="#201F1E", font=f_sm)
    
    # 我的資料包含標題 Checkbox
    d.rectangle([480, 130, 495, 145], outline="#0078D4", fill="#0078D4")
    d.text((483, 129), "v", fill="white", font=f_bold)
    d.text((505, 130), "我的資料有標題 (Headers)", fill="#201F1E", font=f_main)
    
    # 清單列表區域
    d.rectangle([75, 170, 725, 360], outline="#C8C6C4", fill="#FAFAFA")
    
    check_cols = [
        ("提單號碼 (BL_NO) [主要鍵]", True),
        ("訂艙代號 (Booking_ID)", True),
        ("結關日期 (Closing_Date)", False),
        ("出發港 (POL)", False),
        ("目的港 (POD)", False),
        ("櫃號 (Container_No)", True),
        ("櫃型 (Type)", False)
    ]
    cy = 185
    for cname, is_chk in check_cols:
        d.rectangle([95, cy+3, 110, cy+18], outline="#0078D4" if is_chk else "#8A8886", fill="#0078D4" if is_chk else "#FFFFFF")
        if is_chk:
            d.text((98, cy+2), "v", fill="white", font=f_bold)
        d.text((125, cy), cname, fill="#201F1E" if is_chk else "#605E5C", font=f_main)
        if "[主要鍵]" in cname:
            d.rectangle([340, cy-2, 530, cy+20], fill="#DFF6DD", outline="#107C41")
            d.text((345, cy+1), "航運業務核心識別鍵 (Key)", fill="#107C41", font=f_sm)
        cy += 24
        
    # 確定 / 取消按鈕
    d.rectangle([540, 390, 620, 420], fill="#0078D4")
    d.text((560, 398), "確定", fill="white", font=f_bold)
    d.rectangle([635, 390, 715, 420], outline="#8A8886", fill="#F3F2F1")
    d.text((655, 398), "取消", fill="#201F1E", font=f_main)
    
    im.save("assets/fig2_remove_duplicates.png")
    print("Saved assets/fig2_remove_duplicates.png")

# 截圖 3: 資料剖析精靈 (Text to Columns Wizard)
def make_text_to_columns_img():
    w, h = 820, 500
    im = Image.new("RGB", (w, h), "#FFFFFF")
    d = ImageDraw.Draw(im)
    
    # 視窗外框
    d.rectangle([30, 30, 790, 470], fill="#FFFFFF", outline="#0078D4", width=2)
    d.rectangle([32, 32, 788, 65], fill="#F3F2F1")
    d.text((50, 40), "文字探勘 / 轉換成欄精靈 - 步驟 2/3 (分隔符號設定)", fill="#201F1E", font=f_bold)
    d.text((765, 40), "X", fill="#605E5C", font=f_bold)
    
    d.text((50, 80), "這個畫面讓您設定資料包含的分隔符號 (Delimiters)。", fill="#323130", font=f_main)
    
    # 分隔符號核取清單
    d.rectangle([50, 110, 300, 260], outline="#C8C6C4", fill="#FAFAFA")
    d.text((60, 120), "分隔符號 (Delimiters)", fill="#605E5C", font=f_bold)
    
    delims = [("Tab 鍵", False), ("分號 (;)", False), ("逗號 (,)", True), ("空格 (Space)", False), ("其他 (-):", True)]
    dy = 150
    for name, is_on in delims:
        d.rectangle([70, dy+2, 85, dy+17], outline="#0078D4" if is_on else "#8A8886", fill="#0078D4" if is_on else "#FFFFFF")
        if is_on:
            d.text((73, dy+1), "v", fill="white", font=f_bold)
        d.text((95, dy), name, fill="#201F1E", font=f_main)
        dy += 22
        
    # 連續分隔符號視為單一處理 Checkbox
    d.rectangle([50, 275, 65, 290], outline="#0078D4", fill="#0078D4")
    d.text((53, 274), "v", fill="white", font=f_bold)
    d.text((75, 275), "將連續的分隔符號視為單一處理", fill="#201F1E", font=f_sm)
    
    # 資料預覽區域 (Data Preview)
    d.text((320, 110), "資料預覽 (Data Preview) - 提單拆解與金額格式化:", fill="#605E5C", font=f_bold)
    d.rectangle([320, 135, 760, 260], outline="#8A8886", fill="#FFFFFF")
    
    # 模擬資料表預覽
    d.line([320, 165, 760, 165], fill="#C8C6C4")
    d.line([460, 135, 460, 260], fill="#C8C6C4") # 分隔線 1
    d.line([600, 135, 600, 260], fill="#C8C6C4") # 分隔線 2
    
    d.text((330, 142), "航商代碼 (Prefix)", fill="#0078D4", font=f_sm)
    d.text((470, 142), "流水號碼 (Serial)", fill="#0078D4", font=f_sm)
    d.text((610, 142), "港口代碼 (Port)", fill="#0078D4", font=f_sm)
    
    sample_preview = [
        ("MSKU", "8291034", "LAX"),
        ("OOLU", "1102934", "HAM"),
        ("EMCU", "9923811", "NYC"),
        ("YMLU", "5510293", "DXB"),
    ]
    py = 175
    for s1, s2, s3 in sample_preview:
        d.text((330, py), s1, fill="#201F1E", font=f_main)
        d.text((470, py), s2, fill="#201F1E", font=f_main)
        d.text((610, py), s3, fill="#201F1E", font=f_main)
        py += 20
        
    # 操作提示醒目框
    d.rectangle([50, 320, 760, 390], fill="#EFF6FC", outline="#2B88D8")
    d.text((65, 330), "[航運實務小提示]", fill="#004E8C", font=f_bold)
    d.text((65, 355), "提單號碼如 MSKU-8291034 可透過「其他: -」迅速拆出船公司代號 (MSKU) 與編號，", fill="#323130", font=f_sm)
    d.text((65, 370), "便於後續分類匯總各航商之運載量與運費支出！", fill="#323130", font=f_sm)
    
    # 導覽按鈕
    d.rectangle([460, 420, 540, 450], outline="#8A8886", fill="#F3F2F1")
    d.text((475, 428), "< 上一步", fill="#201F1E", font=f_sm)
    d.rectangle([550, 420, 630, 450], fill="#0078D4")
    d.text((570, 428), "下一步 >", fill="white", font=f_bold)
    d.rectangle([640, 420, 720, 450], outline="#8A8886", fill="#F3F2F1")
    d.text((660, 428), "完成", fill="#605E5C", font=f_sm)
    
    im.save("assets/fig3_text_to_columns.png")
    print("Saved assets/fig3_text_to_columns.png")

# 截圖 4: 常用文字與清理函數示意圖 (TRIM, CLEAN, VALUE, TEXT)
def make_formula_diagram():
    w, h = 850, 460
    im = Image.new("RGB", (w, h), "#FFFFFF")
    d = ImageDraw.Draw(im)
    
    # 頂部橫幅
    d.rectangle([0, 0, w, 50], fill="#004E8C")
    d.text((25, 12), "航運資料前處理核心四大函數結構 (Text Cleaning Formulas)", fill="white", font=f_bold)
    
    cards = [
        ("1. TRIM 函數 (去除非必要空格)", 
         "=TRIM( text )", 
         "輸入範例:  \"  Kaohsiung  \"\n處理結果:  \"Kaohsiung\"\n航運場景: 消除碼頭、港口、船名複製貼上時多餘的首尾空白。",
         "#E8F4F8", "#0078D4"),
        
        ("2. CLEAN 函數 (清除不可見字元)", 
         "=CLEAN( text )", 
         "輸入範例:  \"EDI_TXT\\n\\r8291\"\n處理結果:  \"EDI_TXT8291\"\n航運場景: 清除自航運 EDI 或報關系統匯出之換行與控制碼。",
         "#EBF3EB", "#107C41"),
        
        ("3. VALUE 函數 (文字型數值轉換)", 
         "=VALUE( text )", 
         "輸入範例:  \" 18,500 \"\n處理結果:  18500 (純數值)\n航運場景: 將系統匯出的帶千分位或空格字串轉為可計算數值。",
         "#FFF4CE", "#D83B01"),
        
        ("4. SUBSTITUTE 函數 (精準符號替換)", 
         "=SUBSTITUTE( text, old_text, new_text )", 
         "輸入範例:  \"$1,850\" 拿掉 \"$\" -> \"1,850\"\n結合VALUE: =VALUE(SUBSTITUTE(A1, \"$\", \"\"))\n航運場景: 剔除幣別符號($/NTD)還原海運費純數值。",
         "#F3E5F5", "#5C2D91")
    ]
    
    positions = [(40, 70), (440, 70), (40, 260), (440, 260)]
    for (title, syntax, desc, bg_color, border_color), (cx, cy) in zip(cards, positions):
        d.rectangle([cx, cy, cx+370, cy+170], fill=bg_color, outline=border_color, width=2)
        d.text((cx+15, cy+12), title, fill=border_color, font=f_bold)
        
        d.rectangle([cx+15, cy+38, cx+355, cy+68], fill="#FFFFFF", outline="#C8C6C4")
        d.text((cx+25, cy+44), syntax, fill="#201F1E", font=f_bold)
        
        d.text((cx+15, cy+78), desc, fill="#323130", font=f_sm)
        
    im.save("assets/fig4_cleaning_formulas.png")
    print("Saved assets/fig4_cleaning_formulas.png")

# 截圖 5: 樞紐分析表與前處理前後比對
def make_data_comparison_img():
    w, h = 900, 480
    im = Image.new("RGB", (w, h), "#FAFAFA")
    d = ImageDraw.Draw(im)
    
    # 標題
    d.rectangle([0, 0, w, 45], fill="#323130")
    d.text((25, 10), "資料前處理前後 (Before vs. After) 成果對比展示", fill="white", font=f_bold)
    
    # 左側: 處理前 (Raw Data)
    d.rectangle([30, 65, 430, 440], fill="#FFFFFF", outline="#D13438", width=2)
    d.rectangle([30, 65, 430, 105], fill="#FDE7E9")
    d.text((45, 75), "[ 處理前 : 原始混亂髒資料 ]", fill="#A80000", font=f_bold)
    
    raw_flaws = [
        "- 提單編號夾帶首尾空格 (如 '  MSKU-8291034 ')",
        "- 港口英文大小寫不一、夾帶空白 (' Kaohsiung ')",
        "- 日期格式混雜 (2024/05/01、2024-05-02、2024.05.02)",
        "- 存在完全重複申報提單 (重複計費風險)",
        "- 貨物重量欄位混入 'N/A' 文字型，無法計算",
        "- 海運費欄位帶有 '$' 與千分號，無法進行加總公式",
        "=> 後果: 樞紐分析表統計失真、計算報錯 #VALUE!"
    ]
    ry = 120
    for flaw in raw_flaws:
        d.text((45, ry), flaw, fill="#323130" if not flaw.startswith("=>") else "#A80000", font=f_sm)
        ry += 42
        
    # 右側: 處理後 (Clean Data)
    d.rectangle([470, 65, 870, 440], fill="#FFFFFF", outline="#107C41", width=2)
    d.rectangle([470, 65, 870, 105], fill="#DFF6DD")
    d.text((485, 75), "[ 處理後 : 標準化結構化分析資料庫 ]", fill="#107C41", font=f_bold)
    
    clean_benefits = [
        "+ 提單欄位一律去除空格並轉為統一全大寫",
        "+ 港口統一正規名稱，消除因空格導致的重複分類",
        "+ 日期統一規範為 YYYY-MM-DD 標準日期格式",
        "+ 執行去重機制，精確剔除重複訂艙紀錄",
        "+ 異常文字/缺漏值予以填補或明確標註狀態",
        "+ 運費與重量淨化為純數值格式，套用正確貨幣顯示",
        "=> 成果: 支援 100% 正確的樞紐分析與運載統計報表！"
    ]
    cy = 120
    for ben in clean_benefits:
        d.text((485, cy), ben, fill="#323130" if not ben.startswith("=>") else "#107C41", font=f_sm)
        cy += 42
        
    im.save("assets/fig5_data_comparison.png")
    print("Saved assets/fig5_data_comparison.png")

make_ui_interface_img()
make_remove_duplicates_img()
make_text_to_columns_img()
make_formula_diagram()
make_data_comparison_img()
