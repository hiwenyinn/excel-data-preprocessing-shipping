import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# 建立 16:9 寬螢幕投影片簡報
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6] # 全空白版面，利於精準排版

# 主色調定義 (航運深海藍色系)
COLOR_BG = RGBColor(11, 19, 43)          # 深藍黑底色 #0B132B
COLOR_CARD = RGBColor(28, 37, 65)        # 卡片背景 #1C2541
COLOR_TEXT_MAIN = RGBColor(240, 244, 248)# 主文字白色
COLOR_TEXT_MUTED = RGBColor(160, 174, 192)# 次要文字淡灰
COLOR_EMERALD = RGBColor(16, 185, 129)   # 亮綠色 #10B981
COLOR_SKY = RGBColor(56, 189, 248)       # 淺藍色 #38BDF8
COLOR_AMBER = RGBColor(245, 158, 11)     # 琥珀橙 #F59E0B
COLOR_BORDER = RGBColor(45, 55, 72)      # 邊框灰 #2D3748

def add_base_slide(chapter_tag, title_text, slide_num):
    slide = prs.slides.add_slide(blank_layout)
    
    # 全域深藍背景
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = COLOR_BG
    bg.line.fill.background()
    
    # 頂部裝飾條
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.08))
    bar.fill.solid()
    bar.fill.fore_color.rgb = COLOR_EMERALD
    bar.line.fill.background()
    
    # 頂部標題區
    txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(10), Inches(1.1))
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    
    p0 = tf.paragraphs[0]
    p0.text = chapter_tag.upper()
    p0.font.name = "微軟正黑體"
    p0.font.size = Pt(11)
    p0.font.bold = True
    p0.font.color.rgb = COLOR_EMERALD
    p0.space_after = Pt(2)
    
    p1 = tf.add_paragraph()
    p1.text = title_text
    p1.font.name = "微軟正黑體"
    p1.font.size = Pt(22)
    p1.font.bold = True
    p1.font.color.rgb = COLOR_TEXT_MAIN
    
    # 右上角頁碼標籤
    p_box = slide.shapes.add_textbox(Inches(11.2), Inches(0.4), Inches(1.4), Inches(0.5))
    p_tf = p_box.text_frame
    p_tf.margin_left = p_tf.margin_top = p_tf.margin_right = p_tf.margin_bottom = 0
    p_p = p_tf.paragraphs[0]
    p_p.alignment = PP_ALIGN.RIGHT
    p_p.text = f"{slide_num:02d} / 10"
    p_p.font.name = "Calibri"
    p_p.font.size = Pt(13)
    p_p.font.bold = True
    p_p.font.color.rgb = COLOR_TEXT_MUTED
    
    # 底部註腳
    f_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.05), Inches(11.8), Inches(0.35))
    f_tf = f_box.text_frame
    f_tf.margin_left = f_tf.margin_top = f_tf.margin_right = f_tf.margin_bottom = 0
    fp = f_tf.paragraphs[0]
    fp.text = "航運行政人員 MS Excel 資料前處理實戰指南 | 國際物流數據化標準教材"
    fp.font.name = "微軟正黑體"
    fp.font.size = Pt(9.5)
    fp.font.color.rgb = RGBColor(100, 116, 139)
    
    return slide

# -------------------------------------------------------------
# Slide 1: 封面投影片 (Cover)
# -------------------------------------------------------------
slide1 = prs.slides.add_slide(blank_layout)
bg1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
bg1.fill.solid()
bg1.fill.fore_color.rgb = COLOR_BG
bg1.line.fill.background()

# 頂部翡翠綠裝飾條
t_bar = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.12))
t_bar.fill.solid()
t_bar.fill.fore_color.rgb = COLOR_EMERALD
t_bar.line.fill.background()

# 標題文字區
t_box = slide1.shapes.add_textbox(Inches(1.0), Inches(1.3), Inches(11.333), Inches(3.2))
ttf = t_box.text_frame
ttf.word_wrap = True

p_badge = ttf.paragraphs[0]
p_badge.text = "■ 航運物流數位轉型與行政作業實戰教材 ■"
p_badge.font.name = "微軟正黑體"
p_badge.font.size = Pt(14)
p_badge.font.bold = True
p_badge.font.color.rgb = COLOR_EMERALD
p_badge.space_after = Pt(14)

p_main = ttf.add_paragraph()
p_main.text = "航運行政人員 MS Excel 資料前處理\n操作教學手冊與實戰指南"
p_main.font.name = "微軟正黑體"
p_main.font.size = Pt(38)
p_main.font.bold = True
p_main.font.color.rgb = COLOR_TEXT_MAIN
p_main.space_after = Pt(14)

p_sub = ttf.add_paragraph()
p_sub.text = "從零開始掌握航運數據清洗、格式正規化、提單主鍵去重與分析就緒全流程"
p_sub.font.name = "微軟正黑體"
p_sub.font.size = Pt(16)
p_sub.font.color.rgb = COLOR_TEXT_MUTED

# 底部 4 個元數據卡片
meta_items = [
    ("目標受眾", "Forwarder / Carrier 行政初學者", COLOR_SKY),
    ("示範工具", "Microsoft Excel 365 / 2021", COLOR_EMERALD),
    ("配套練習庫", "30筆示範 + 300筆實戰練習", COLOR_AMBER),
    ("操作導向", "資料前處理 (Tidy Data 哲學)", COLOR_SKY)
]

for idx, (label, val, c_color) in enumerate(meta_items):
    x = Inches(1.0 + idx * 2.9)
    y = Inches(4.8)
    card = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(2.7), Inches(1.5))
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_CARD
    card.line.color.rgb = COLOR_BORDER
    
    ctf = card.text_frame
    ctf.word_wrap = True
    ctf.margin_left = ctf.margin_right = Inches(0.2)
    ctf.margin_top = Inches(0.2)
    
    cp1 = ctf.paragraphs[0]
    cp1.text = label
    cp1.font.name = "微軟正黑體"
    cp1.font.size = Pt(11)
    cp1.font.color.rgb = COLOR_TEXT_MUTED
    cp1.space_after = Pt(4)
    
    cp2 = ctf.add_paragraph()
    cp2.text = val
    cp2.font.name = "微軟正黑體"
    cp2.font.size = Pt(13)
    cp2.font.bold = True
    cp2.font.color.rgb = c_color

# -------------------------------------------------------------
# Slide 2: Excel 操作介面解構 (圖 1-1)
# -------------------------------------------------------------
slide2 = add_base_slide("第一章 · 介面環境", "認識航運行政專用 Excel 操作介面與核心功能區", 2)
# 插入圖 1-1 (左側)
slide2.shapes.add_picture("assets/fig1_excel_interface.png", Inches(0.8), Inches(1.7), Inches(6.8), Inches(4.8))

# 右側卡片群
right_cards = [
    ("1. 功能區 Ribbon【資料 (Data)】標籤頁", "前處理的核心司令部！內含「移除重複項」、「資料剖析」與「排序篩選」，為航運清洗最高頻使用區。", COLOR_EMERALD),
    ("2. 名稱方塊與公式編輯列", "精確檢驗提單字串真實長度，辨識藏在代碼末端之不可見空白與換行符號（避免 VLOOKUP 查無資料）。", COLOR_SKY),
    ("3. 狀態列 (Status Bar) 即時核帳", "反白選取海運費與毛重欄位時，右下角即時呈現「計數」與「數值合計」，不需寫公式即可快速驗證總金額。", COLOR_AMBER)
]

for idx, (title, desc, color) in enumerate(right_cards):
    y = Inches(1.7 + idx * 1.65)
    card = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.8), y, Inches(4.7), Inches(1.45))
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_CARD
    card.line.color.rgb = color
    card.line.width = Pt(1.5)
    
    ctf = card.text_frame
    ctf.word_wrap = True
    ctf.margin_left = ctf.margin_right = Inches(0.25)
    ctf.margin_top = Inches(0.2)
    
    cp1 = ctf.paragraphs[0]
    cp1.text = title
    cp1.font.name = "微軟正黑體"
    cp1.font.size = Pt(12)
    cp1.font.bold = True
    cp1.font.color.rgb = color
    cp1.space_after = Pt(4)
    
    cp2 = ctf.add_paragraph()
    cp2.text = desc
    cp2.font.name = "微軟正黑體"
    cp2.font.size = Pt(10)
    cp2.font.color.rgb = COLOR_TEXT_MUTED

# -------------------------------------------------------------
# Slide 3: 前處理基本常識與 GIGO 定律
# -------------------------------------------------------------
slide3 = add_base_slide("第二章 · 核心概念", "資料前處理應具備之基本常識與 GIGO 定律", 3)

c3_data = [
    ("【航運髒資料常見型態】", 
     "• 前後夾帶不可見空白（Kaohsiung 多空格）\n• 英文大小寫混亂（msku vs MSKU）\n• 日期多元混雜（2024/05/01 與 2024.05.02）\n• 重複入單導致運費虛增一倍\n\n營運風險：樞紐分析失真、海關退單、重複請款",
     RGBColor(239, 68, 68), "髒資料危害"),
    
    ("【GIGO 與 SOLAS VGM】", 
     "• 資訊科學核心法則：「Garbage In, Garbage Out」\n• 貨物毛重若混入文字 'N/A' 或單位 'kg'\n• SUM 統計為 0 或報錯 #VALUE!\n\n海事法規：違反國際公約貨櫃總重驗證(VGM)，將致船舶配載配重失衡傾斜遭扣船",
     COLOR_AMBER, "SOLAS 公約公信力"),
    
    ("【Tidy Data 乾淨六守則】", 
     "1. 絕不使用合併儲存格（分析大忌）\n2. 首列維持單一標題列（Header）\n3. 一列代表一筆訂艙提單實體\n4. 一欄僅存單一屬性（POL 與 POD 分立）\n5. 原始資料底層不留人工小計列\n6. 杜絕排版用空白空列與斷行",
     COLOR_EMERALD, "Tidy Data 規範")
]

for idx, (title, desc, color, tag) in enumerate(c3_data):
    x = Inches(0.8 + idx * 4.0)
    card = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.7), Inches(3.75), Inches(4.9))
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_CARD
    card.line.color.rgb = color
    card.line.width = Pt(1.5)
    
    ctf = card.text_frame
    ctf.word_wrap = True
    ctf.margin_left = ctf.margin_right = Inches(0.3)
    ctf.margin_top = Inches(0.3)
    
    cp1 = ctf.paragraphs[0]
    cp1.text = title
    cp1.font.name = "微軟正黑體"
    cp1.font.size = Pt(14)
    cp1.font.bold = True
    cp1.font.color.rgb = color
    cp1.space_after = Pt(10)
    
    cp2 = ctf.add_paragraph()
    cp2.text = desc
    cp2.font.name = "微軟正黑體"
    cp2.font.size = Pt(11)
    cp2.font.color.rgb = COLOR_TEXT_MAIN
    
# -------------------------------------------------------------
# Slide 4: 航運前處理四大核心函數
# -------------------------------------------------------------
slide4 = add_base_slide("第二/三章 · 函數精華", "航運文字與數據清洗四大核心函數 (Formulas)", 4)
slide4.shapes.add_picture("assets/fig4_cleaning_formulas.png", Inches(0.8), Inches(1.7), Inches(6.8), Inches(4.8))

func_cards = [
    ("=TRIM( text )", "剔除頭尾多餘空格", "將 \" Kaohsiung \" 還原為純淨港口名稱，解決 VLOOKUP 查無提單之通病。", COLOR_SKY),
    ("=CLEAN( text )", "清除 EDI 非列印控制碼", "移除自船公司 EDI 報關文字檔匯出之換行與特殊控制碼 (ASCII 0~31)。", COLOR_EMERALD),
    ("=VALUE( text )", "文字型數值轉換為計算數值", "將系統匯出之文字型重量 \" 18,500 \" 轉換為純數值 18500，方可正常加總。", COLOR_AMBER),
    ("=SUBSTITUTE( text, \"$\", \"\" )", "精準剔除幣別與千分位", "結合嵌套公式消除運費中之 \"$\" 與 \",\"，還原純數字供財務匯率計算。", RGBColor(168, 85, 247))
]

for idx, (syntax, label, desc, color) in enumerate(func_cards):
    y = Inches(1.7 + idx * 1.25)
    card = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.8), y, Inches(4.7), Inches(1.15))
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_CARD
    card.line.color.rgb = color
    card.line.width = Pt(1.2)
    
    ctf = card.text_frame
    ctf.word_wrap = True
    ctf.margin_left = ctf.margin_right = Inches(0.25)
    ctf.margin_top = Inches(0.12)
    
    cp1 = ctf.paragraphs[0]
    cp1.text = f"{syntax}  |  {label}"
    cp1.font.name = "Consolas"
    cp1.font.size = Pt(11)
    cp1.font.bold = True
    cp1.font.color.rgb = color
    cp1.space_after = Pt(2)
    
    cp2 = ctf.add_paragraph()
    cp2.text = desc
    cp2.font.name = "微軟正黑體"
    cp2.font.size = Pt(9.5)
    cp2.font.color.rgb = COLOR_TEXT_MUTED

# -------------------------------------------------------------
# Slide 5: 示範檔實作 - 提單主鍵去重
# -------------------------------------------------------------
slide5 = add_base_slide("第三章 · 示範檔實作", "步驟四：提單主鍵去重複作業 (Remove Duplicates)", 5)
slide5.shapes.add_picture("assets/fig2_remove_duplicates.png", Inches(0.8), Inches(1.7), Inches(6.8), Inches(4.8))

sop_box = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.8), Inches(1.7), Inches(4.7), Inches(3.2))
sop_box.fill.solid()
sop_box.fill.fore_color.rgb = COLOR_CARD
sop_box.line.color.rgb = COLOR_EMERALD
sop_box.line.width = Pt(1.5)

stf = sop_box.text_frame
stf.word_wrap = True
stf.margin_left = stf.margin_right = Inches(0.3)
stf.margin_top = Inches(0.25)

sp0 = stf.paragraphs[0]
sp0.text = "⚡ 去重複 4 步驟標準操作程序 (SOP)"
sp0.font.name = "微軟正黑體"
sp0.font.size = Pt(13)
sp0.font.bold = True
sp0.font.color.rgb = COLOR_EMERALD
sp0.space_after = Pt(8)

steps_text = (
    "1. 選取資料區任一儲存格（A1:J31）。\n"
    "2. 功能區切換至【資料 (Data)】-> 點選【移除重複項】。\n"
    "3. 務必勾選「我的資料有標題」，點選「取消全選」。\n"
    "4. 單獨勾選航運唯一主鍵【提單號碼(BL_NO)】與【訂艙代號】。\n"
    "5. 點選「確定」，系統提示成功移除 1 個重複值，保留 29 個唯一值！"
)
sp1 = stf.add_paragraph()
sp1.text = steps_text
sp1.font.name = "微軟正黑體"
sp1.font.size = Pt(10.5)
sp1.font.color.rgb = COLOR_TEXT_MAIN

# 航運提醒小框
warn_box = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.8), Inches(5.1), Inches(4.7), Inches(1.4))
warn_box.fill.solid()
warn_box.fill.fore_color.rgb = COLOR_CARD
warn_box.line.color.rgb = COLOR_AMBER
warn_box.line.width = Pt(1.5)

wtf = warn_box.text_frame
wtf.word_wrap = True
wtf.margin_left = wtf.margin_right = Inches(0.25)
wtf.margin_top = Inches(0.18)

wp0 = wtf.paragraphs[0]
wp0.text = "【航運專業小叮嚀】"
wp0.font.name = "微軟正黑體"
wp0.font.size = Pt(11)
wp0.font.bold = True
wp0.font.color.rgb = COLOR_AMBER
wp0.space_after = Pt(2)

wp1 = wtf.add_paragraph()
wp1.text = "不可盲目全選所有欄位去重！同一客戶可能同日出兩筆相同重之貨物，提單主鍵 (BL) 才是航運業判定唯一性的唯一基準！"
wp1.font.name = "微軟正黑體"
wp1.font.size = Pt(9.5)
wp1.font.color.rgb = COLOR_TEXT_MUTED

# -------------------------------------------------------------
# Slide 6: 資料剖析精靈與航商代碼拆解
# -------------------------------------------------------------
slide6 = add_base_slide("第三/四章 · 結構拆解", "利用「資料剖析」極速拆解提單航商代碼與流水號", 6)
slide6.shapes.add_picture("assets/fig3_text_to_columns.png", Inches(0.8), Inches(1.7), Inches(6.8), Inches(4.8))

p6_box = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.8), Inches(1.7), Inches(4.7), Inches(3.2))
p6_box.fill.solid()
p6_box.fill.fore_color.rgb = COLOR_CARD
p6_box.line.color.rgb = COLOR_SKY
p6_box.line.width = Pt(1.5)

p6tf = p6_box.text_frame
p6tf.word_wrap = True
p6tf.margin_left = p6tf.margin_right = Inches(0.3)
p6tf.margin_top = Inches(0.25)

pp0 = p6tf.paragraphs[0]
pp0.text = "提單代碼剖析操作步驟 (Delimiters)"
pp0.font.name = "微軟正黑體"
pp0.font.size = Pt(13)
pp0.font.bold = True
pp0.font.color.rgb = COLOR_SKY
pp0.space_after = Pt(8)

pp_txt = (
    "1. 選取提單號碼 A 欄（例如 MSKU-8291034）。\n"
    "2. 點選【資料】->【資料剖析 (Text to Columns)】。\n"
    "3. 步驟 1 選擇「分隔符號 (Delimited)」-> 下一步。\n"
    "4. 步驟 2 勾選「其他 (Other)」，於框中鍵入減號破折號「-」。\n"
    "5. 步驟 3 目的地選擇相鄰空白欄 -> 點擊完成！\n"
    "瞬間將數百筆提單拆解為「航商字軌」與「7碼流水號」。"
)
pp1 = p6tf.add_paragraph()
pp1.text = pp_txt
pp1.font.name = "微軟正黑體"
pp1.font.size = Pt(10.5)
pp1.font.color.rgb = COLOR_TEXT_MAIN

# 效益框
ben_box = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.8), Inches(5.1), Inches(4.7), Inches(1.4))
ben_box.fill.solid()
ben_box.fill.fore_color.rgb = COLOR_CARD
ben_box.line.color.rgb = COLOR_EMERALD
ben_box.line.width = Pt(1.5)

btf = ben_box.text_frame
btf.word_wrap = True
btf.margin_left = btf.margin_right = Inches(0.25)
btf.margin_top = Inches(0.18)

bp0 = btf.paragraphs[0]
bp0.text = "【航運營運管理效益】"
bp0.font.name = "微軟正黑體"
bp0.font.size = Pt(11)
bp0.font.bold = True
bp0.font.color.rgb = COLOR_EMERALD
bp0.space_after = Pt(2)

bp1 = btf.add_paragraph()
bp1.text = "獨立拆解出航商字軌 (MSKU, EMCU, WHLC, YMLU) 後，後續可極速以樞紐分析表匯總各船東年度運費支出與櫃量佔比！"
bp1.font.name = "微軟正黑體"
bp1.font.size = Pt(9.5)
bp1.font.color.rgb = COLOR_TEXT_MUTED

# -------------------------------------------------------------
# Slide 7: 前處理前後成果對照看板
# -------------------------------------------------------------
slide7 = add_base_slide("第三章 · 成果驗證", "示範檔前處理前後 (Before vs. After) 成果對比與品質看板", 7)
slide7.shapes.add_picture("assets/fig5_data_comparison.png", Inches(0.8), Inches(1.7), Inches(6.8), Inches(4.8))

c7_raw = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.8), Inches(1.7), Inches(4.7), Inches(2.25))
c7_raw.fill.solid()
c7_raw.fill.fore_color.rgb = COLOR_CARD
c7_raw.line.color.rgb = RGBColor(239, 68, 68)
c7_raw.line.width = Pt(1.5)

rtf = c7_raw.text_frame
rtf.word_wrap = True
rtf.margin_left = rtf.margin_right = Inches(0.25)
rtf.margin_top = Inches(0.18)
rp0 = rtf.paragraphs[0]
rp0.text = "[ 處理前 : 原始混亂髒資料 ]"
rp0.font.name = "微軟正黑體"
rp0.font.size = Pt(12)
rp0.font.bold = True
rp0.font.color.rgb = RGBColor(239, 68, 68)
rp0.space_after = Pt(4)
rp1 = rtf.add_paragraph()
rp1.text = "• 提單頭尾夾帶空格、大小寫不一 (msku vs MSKU)\n• 結關日期格式多元混雜 (斜線/橫線/句點/緊湊碼)\n• 存在完全重複申報提單、運費包含貨幣號 $ 無法加總"
rp1.font.name = "微軟正黑體"
rp1.font.size = Pt(10)
rp1.font.color.rgb = COLOR_TEXT_MUTED

c7_clean = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.8), Inches(4.15), Inches(4.7), Inches(2.35))
c7_clean.fill.solid()
c7_clean.fill.fore_color.rgb = COLOR_CARD
c7_clean.line.color.rgb = COLOR_EMERALD
c7_clean.line.width = Pt(1.5)

ctf = c7_clean.text_frame
ctf.word_wrap = True
ctf.margin_left = ctf.margin_right = Inches(0.25)
ctf.margin_top = Inches(0.18)
cp0 = ctf.paragraphs[0]
cp0.text = "[ 處理後 : 標準化結構化分析資料庫 ]"
cp0.font.name = "微軟正黑體"
cp0.font.size = Pt(12)
cp0.font.bold = True
cp0.font.color.rgb = COLOR_EMERALD
cp0.space_after = Pt(4)
cp1 = ctf.add_paragraph()
cp1.text = "• 提單全面去除空格並轉全大寫，完成唯一性去重\n• 日期標準化為 YYYY-MM-DD\n• 運費與重量淨化為純數值，支援 100% 正確樞紐分析！"
cp1.font.name = "微軟正黑體"
cp1.font.size = Pt(10)
cp1.font.color.rgb = COLOR_TEXT_MAIN

# -------------------------------------------------------------
# Slide 8: 300筆實戰練習檔情境設定
# -------------------------------------------------------------
slide8 = add_base_slide("第四章 · 實戰練習庫", "300筆航運案例進階實戰練習庫 (Shipping_Practice_300)", 8)

tasks = [
    ("【任務一：提單與航商維度拆解】",
     "• 檔案路徑：data/Shipping_Practice_300.xlsx\n• 涵蓋八大船公司 (MSKU, EMCU, ONEU...)\n• 任務要求：利用「資料剖析精靈」或文字函數，拆解提單為航商字軌與7碼流水號。\n\n推薦公式：\n=LEFT(A2, 4)  與  =RIGHT(A2, 7)",
     COLOR_SKY),
    
    ("【任務二：去空格與重複審核】",
     "• 清洗港口 (POL/POD) 多餘字串空格\n• 全表埋入 5 筆完全重複訂艙提單\n• 任務要求：執行主鍵去重，讓總筆數由 305 筆自動收斂為標準 300 筆唯一值。\n\n推薦操作：\n=PROPER(TRIM(D2)) 結合【移除重複項】",
     COLOR_EMERALD),
    
    ("【任務三：金額數值化與超重警示】",
     "• 運費含有 $ 記號與千分逗點，轉為純數值\n• 毛重欄位 N/A 缺失值填補為 0\n• 任務要求：設定條件式格式，將超過 22,000 KG 之「超重重櫃」標註為黃底紅字。\n\n推薦公式：\n=VALUE(SUBSTITUTE(TRIM(I2), \"$\", \"\"))",
     COLOR_AMBER)
]

for idx, (title, desc, color) in enumerate(tasks):
    x = Inches(0.8 + idx * 4.0)
    card = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.7), Inches(3.75), Inches(4.9))
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_CARD
    card.line.color.rgb = color
    card.line.width = Pt(1.5)
    
    ctf = card.text_frame
    ctf.word_wrap = True
    ctf.margin_left = ctf.margin_right = Inches(0.3)
    ctf.margin_top = Inches(0.3)
    
    cp1 = ctf.paragraphs[0]
    cp1.text = title
    cp1.font.name = "微軟正黑體"
    cp1.font.size = Pt(13)
    cp1.font.bold = True
    cp1.font.color.rgb = color
    cp1.space_after = Pt(10)
    
    cp2 = ctf.add_paragraph()
    cp2.text = desc
    cp2.font.name = "微軟正黑體"
    cp2.font.size = Pt(10.5)
    cp2.font.color.rgb = COLOR_TEXT_MAIN

# -------------------------------------------------------------
# Slide 9: 300筆實戰挑戰參考解題 SOP
# -------------------------------------------------------------
slide9 = add_base_slide("第四章 · 參考答案", "實戰挑戰題參考標準解題步驟 (SOP)", 9)

sol_items = [
    ("解題任務一（資料剖析極速切分）：",
     "1. 選取提單號碼 A 欄 -> 點選【資料】標籤頁 ->【資料剖析】。\n"
     "2. 步驟 1 選擇「分隔符號」-> 下一步。\n"
     "3. 步驟 2 勾選「其他」，輸入減號「-」-> 目的地指定為相鄰兩空欄 -> 完成。\n"
     "成果：瞬間將 300 筆提單精確切分為「航商代碼」與「7碼流水號」。",
     COLOR_SKY),
    
    ("解題任務二（文字淨化與去重 SOP）：",
     "1. 輔助欄輸入公式：=PROPER(TRIM(D2)) 及 =UPPER(TRIM(E2))。\n"
     "2. 選取輔助欄按 Ctrl+C，於原欄位「貼上值」覆蓋。\n"
     "3. 點選【資料】->【移除重複項】，單獨勾選提單號碼。\n"
     "成果：系統成功刪除 5 筆重複單，總筆數由 305 筆自動收斂為標準 300 筆唯一值！",
     COLOR_EMERALD),
    
    ("解題任務三（金額淨化與超重條件格式化 SOP）：",
     "1. 運費淨化公式：=VALUE(SUBSTITUTE(SUBSTITUTE(TRIM(I2), \"$\", \"\"), \",\", \"\"))。\n"
     "2. 重量防禦公式：=IFERROR(VALUE(SUBSTITUTE(TRIM(H2), \",\", \"\")), 0)。\n"
     "3. 選取重量欄 ->【常用】->【設定格式化的條件】->【反白標示儲存格規則】->【大於 22000】-> 設定黃底紅字。\n"
     "成果：超重集裝箱一目了然，便於向客戶徵收超重附加費 (Overweight Surcharge)！",
     COLOR_AMBER)
]

for idx, (title, desc, color) in enumerate(sol_items):
    y = Inches(1.7 + idx * 1.65)
    card = slide9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), y, Inches(11.733), Inches(1.48))
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_CARD
    card.line.color.rgb = color
    card.line.width = Pt(1.5)
    
    ctf = card.text_frame
    ctf.word_wrap = True
    ctf.margin_left = ctf.margin_right = Inches(0.3)
    ctf.margin_top = Inches(0.18)
    
    cp1 = ctf.paragraphs[0]
    cp1.text = title
    cp1.font.name = "微軟正黑體"
    cp1.font.size = Pt(12.5)
    cp1.font.bold = True
    cp1.font.color.rgb = color
    cp1.space_after = Pt(4)
    
    cp2 = ctf.add_paragraph()
    cp2.text = desc
    cp2.font.name = "微軟正黑體"
    cp2.font.size = Pt(10)
    cp2.font.color.rgb = COLOR_TEXT_MAIN

# -------------------------------------------------------------
# Slide 10: 參考文獻與課後學習資源
# -------------------------------------------------------------
slide10 = add_base_slide("第五章 · 權威來源", "參考來源與航運數據規範文獻 (References)", 10)

refs = [
    ("1. 國際海事組織 (IMO) - SOLAS 公約貨櫃總重驗證指引", "International Maritime Organization. Guidelines Regarding the Verified Gross Mass of a Container (MSC.1/Circ.1475)."),
    ("2. 國際貨櫃局 (BIC) 貨櫃國際代碼標準與檢查碼演算法", "Bureau International des Containers. ISO 6346: Freight containers - Coding, identification and marking."),
    ("3. 數位貨櫃航運協會 (DCSA) 航運資料標準介面", "Digital Container Shipping Association. DCSA Standards for Operational Vessel Schedules and Track & Trace (v2.2)."),
    ("4. 資料科學整潔資料標準 (Tidy Data 哲學)", "Hadley Wickham (2014). Tidy Data. Journal of Statistical Software, 59(10), 1-23."),
    ("5. 微軟官方 Excel 函數支援中心文件庫", "Microsoft Corporation. Microsoft Excel Function Reference: Data Cleansing Best Practices."),
    ("6. 臺灣港務股份有限公司 (TIPC) 國際商港作業手冊", "Taiwan International Ports Corporation. 港務 EDI 連線申報與提單資料登錄作業規範指引.")
]

for idx, (rtitle, rdesc) in enumerate(refs):
    col = idx % 2
    row = idx // 2
    x = Inches(0.8 + col * 5.95)
    y = Inches(1.7 + row * 1.5)
    
    card = slide10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(5.75), Inches(1.35))
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_CARD
    card.line.color.rgb = COLOR_BORDER
    
    ctf = card.text_frame
    ctf.word_wrap = True
    ctf.margin_left = ctf.margin_right = Inches(0.25)
    ctf.margin_top = Inches(0.18)
    
    cp1 = ctf.paragraphs[0]
    cp1.text = rtitle
    cp1.font.name = "微軟正黑體"
    cp1.font.size = Pt(11)
    cp1.font.bold = True
    cp1.font.color.rgb = COLOR_SKY
    cp1.space_after = Pt(2)
    
    cp2 = ctf.add_paragraph()
    cp2.text = rdesc
    cp2.font.name = "Calibri"
    cp2.font.size = Pt(9.5)
    cp2.font.color.rgb = COLOR_TEXT_MUTED

out_pptx = "航運行政Excel資料前處理教學簡報.pptx"
prs.save(out_pptx)
print("Saved successfully as:", out_pptx)
