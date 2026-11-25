# 航運行政人員 MS Excel 資料前處理實戰指南

專為**航運行政人員（海運承攬運送業 Forwarder、船公司 Carrier、報關行）**量身打造之 Excel 資料前處理教學教材與練習庫。

---

## 📌 專案內容與核心檔案清單

### 1. 教學手冊與簡報投影片
- 📘 **`航運行政人員Excel資料前處理教學手冊.docx`**：完整 Microsoft Word 格式教學手冊，含封面、章節目錄、圖目錄、表目錄、標準頁碼與 5 張高解析度操作截圖。
- 🖥️ **`航運行政Excel資料前處理教學簡報.pptx`**：Microsoft PowerPoint 16:9 寬螢幕投影片簡報，收錄 10 頁海運風格教學卡片與圖解。
- 🌐 **`航運行政Excel資料前處理教學簡報.html`**：單一檔案互動式網頁簡報，支援鍵盤方向鍵切換、全螢幕播放與進度條。

### 2. 配套實習資料庫 (`data/`)
- 📄 **`data/Shipping_Demo_Raw_30.xlsx`**：30 筆航運示範髒資料檔（含前後空白、大小寫混雜、日期格式分歧、重複提單、N/A 缺失值、貨幣符號）。
- 📊 **`data/Shipping_Practice_300.xlsx`**：300 筆綜合型航運案例進階實戰練習庫（涵蓋八大船公司、四大港口、全球二十大航線，附三大挑戰任務）。

### 3. 操作示意圖與品質看板 (`assets/`)
- `assets/fig1_excel_interface.png`：Excel 核心介面架構與功能區分佈。
- `assets/fig2_remove_duplicates.png`：提單主鍵去重複項對話框。
- `assets/fig3_text_to_columns.png`：資料剖析精靈拆解航商代碼。
- `assets/fig4_cleaning_formulas.png`：四大清洗核心函數（TRIM, CLEAN, VALUE, SUBSTITUTE）。
- `assets/fig5_data_comparison.png`：前處理前後 (Before vs. After) 成果對照品質看板。

### 4. 自動化生成 Python 腳本
- `make_datasets.py`：生成 30 筆示範檔與 300 筆實戰練習檔。
- `make_images.py`：生成所有 UI 截圖與流程圖。
- `make_manual_docx.py`：排版並編譯生成 Word (.docx) 文件。
- `make_pptx.py`：生成 PowerPoint (.pptx) 投影片簡報。
- `make_ppt_html.py`：生成單一 HTML 互動式簡報。

---

## 🚀 核心學習架構
1. **認識 Excel 介面**：功能區 Ribbon、公式編輯列與狀態列即時核帳。
2. **前處理基本常識**：GIGO 定律、SOLAS VGM 貨櫃配載風險、Tidy Data 六大守則。
3. **示範檔清洗步驟**：
   - 步驟一：`TRIM` 與 `CLEAN` 去除無效空白
   - 步驟二：`UPPER` 英文代碼大小寫標準化
   - 步驟三：結關日期格式正規化 (YYYY-MM-DD)
   - 步驟四：提單號碼 (BL_NO) 主鍵去重
   - 步驟五：`IFERROR` 填補非數值與缺失
   - 步驟六：`SUBSTITUTE` + `VALUE` 運費純數值化
4. **300 筆實戰挑戰與 SOP**：提單維度剖析、5 筆重複單剔除、22,000 KG 超重櫃條件式格式標註。
5. **參考文獻**：IMO SOLAS VGM、BIC ISO 6346、DCSA、Tidy Data (Hadley Wickham)。
