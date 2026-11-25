import json
import os

with open("assets_base64.json", "r") as f:
    raw_imgs = json.load(f)

imgs = {k.replace("\\", "/"): v for k, v in raw_imgs.items()}

img_ui = imgs["assets/fig1_excel_interface.png"]
img_dup = imgs["assets/fig2_remove_duplicates.png"]
img_split = imgs["assets/fig3_text_to_columns.png"]
img_formula = imgs["assets/fig4_cleaning_formulas.png"]
img_comp = imgs["assets/fig5_data_comparison.png"]

html_template = f'''<!DOCTYPE html>
<html lang="zh-Hant">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>航運行政人員 Excel 資料前處理 互動式簡報</title>
  <script src="https://www.gstatic.com/antigravity/web/dev/tailwindcss.min.js"></script>
  <style>
    @keyframes fadeIn {{
      from {{ opacity: 0; transform: translateY(8px); }}
      to {{ opacity: 1; transform: translateY(0); }}
    }}
    .slide-enter {{
      animation: fadeIn 0.35s cubic-bezier(0.16, 1, 0.3, 1) forwards;
    }}
    .progress-bar-smooth {{
      transition: width 0.3s ease-in-out;
    }}
  </style>
</head>
<body class="bg-slate-900 text-slate-100 min-h-screen flex flex-col font-sans select-none antialiased">

  <!-- 頂部控制列 -->
  <header class="h-14 border-b border-slate-800 bg-slate-950/80 backdrop-blur px-6 flex items-center justify-between shrink-0 z-20">
    <div class="flex items-center gap-3">
      <span class="w-3 h-3 rounded-full bg-emerald-500 animate-pulse"></span>
      <h1 class="text-sm font-semibold tracking-wide text-slate-200">航運行政 Excel 資料前處理實戰指南 | 互動式投影片</h1>
      <span class="text-xs px-2 py-0.5 rounded bg-emerald-950/80 text-emerald-400 border border-emerald-800/60 hidden sm:inline">航運數據分析專用</span>
    </div>
    <div class="flex items-center gap-4 text-xs text-slate-400">
      <span id="slideCounter" class="font-mono bg-slate-900 px-3 py-1 rounded-full border border-slate-800 text-slate-300">01 / 10</span>
      <button onclick="toggleFullscreen()" class="hover:text-emerald-400 transition p-1.5 rounded bg-slate-900 hover:bg-slate-800 border border-slate-800" title="全螢幕模式">
        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 8V4m0 0h4M4 4l5 5m11-5h-4m4 0v4m0-4l-5 5M4 16v4m0 0h4m-4 0l5-5m11 5l-5-5m5 5v-4m0 4h-4"/></svg>
      </button>
    </div>
  </header>

  <!-- 進度條 -->
  <div class="w-full bg-slate-800 h-1">
    <div id="progressBar" class="bg-emerald-500 h-1 progress-bar-smooth w-[10%]"></div>
  </div>

  <!-- 主舞台簡報區域 -->
  <main class="flex-1 flex items-center justify-center p-4 sm:p-8 overflow-hidden relative">

    <!-- 幻燈片容器 (16:9 投影片視窗) -->
    <div class="w-full max-w-6xl aspect-[16/9.2] max-h-[82vh] bg-slate-950 rounded-2xl border border-slate-800 shadow-2xl overflow-hidden flex flex-col relative" id="slideStage">

      <!-- 幻燈片 1: 封面 -->
      <section class="slide-content p-8 sm:p-14 flex flex-col justify-between h-full bg-gradient-to-br from-slate-950 via-slate-900 to-emerald-950/30" data-index="1">
        <div>
          <div class="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-emerald-500/10 border border-emerald-500/20 text-emerald-400 text-xs font-semibold mb-6">
            <span>■ 航運物流數位轉型與行政作業實戰教材 ■</span>
          </div>
          <h1 class="text-3xl sm:text-5xl font-black tracking-tight text-white mb-4 leading-tight">
            航運行政人員 MS Excel<br/>
            <span class="text-transparent bg-clip-text bg-gradient-to-r from-emerald-400 to-teal-200">資料前處理操作教學</span>
          </h1>
          <p class="text-slate-400 text-base sm:text-lg max-w-2xl font-light">
            從零開始掌握航運數據清洗、格式正規化、提單主鍵去重與分析就緒全流程！
          </p>
        </div>

        <div class="grid grid-cols-2 sm:grid-cols-4 gap-4 pt-6 border-t border-slate-800/80">
          <div class="bg-slate-900/60 p-3.5 rounded-xl border border-slate-800">
            <div class="text-slate-500 text-xs mb-1">目標受眾</div>
            <div class="text-slate-200 text-xs sm:text-sm font-medium">Forwarder / Carrier 行政初學者</div>
          </div>
          <div class="bg-slate-900/60 p-3.5 rounded-xl border border-slate-800">
            <div class="text-slate-500 text-xs mb-1">示範工具</div>
            <div class="text-slate-200 text-xs sm:text-sm font-medium">Microsoft Excel 365 / 2021</div>
          </div>
          <div class="bg-slate-900/60 p-3.5 rounded-xl border border-slate-800">
            <div class="text-slate-500 text-xs mb-1">配套練習庫</div>
            <div class="text-emerald-400 text-xs sm:text-sm font-medium">30筆示範 + 300筆實戰練習</div>
          </div>
          <div class="bg-slate-900/60 p-3.5 rounded-xl border border-slate-800">
            <div class="text-slate-500 text-xs mb-1">操作導向</div>
            <div class="text-teal-300 text-xs sm:text-sm font-medium">資料前處理 (Tidy Data)</div>
          </div>
        </div>
      </section>

      <!-- 幻燈片 2: Excel 操作介面解構 -->
      <section class="slide-content p-6 sm:p-10 flex flex-col h-full hidden" data-index="2">
        <div class="flex items-center justify-between mb-4">
          <div>
            <span class="text-emerald-400 font-mono text-xs uppercase tracking-wider font-semibold">第一章 · 介面環境</span>
            <h2 class="text-xl sm:text-2xl font-bold text-white mt-1">認識航運行政專用 Excel 操作介面</h2>
          </div>
          <span class="text-xs text-slate-400 bg-slate-900 border border-slate-800 px-2.5 py-1 rounded">圖 1-1 架構解析</span>
        </div>
        <div class="grid grid-cols-1 lg:grid-cols-12 gap-6 flex-1 items-center overflow-hidden">
          <div class="lg:col-span-7 h-full flex flex-col justify-center">
            <div class="rounded-xl overflow-hidden border border-slate-800 bg-black/40 shadow-xl group relative">
              <img src="{img_ui}" alt="Excel 介面圖" class="w-full object-contain max-h-[46vh] rounded-lg">
              <div class="absolute bottom-2 left-2 right-2 bg-slate-950/80 backdrop-blur px-3 py-1.5 rounded text-[11px] text-slate-300 flex justify-between">
                <span>圖 1-1：功能區、編輯列、工作表網格及狀態列</span>
                <span class="text-emerald-400">航運專用重點標註</span>
              </div>
            </div>
          </div>
          <div class="lg:col-span-5 space-y-3 text-xs sm:text-sm">
            <div class="p-3 rounded-lg bg-slate-900/80 border-l-4 border-emerald-500">
              <div class="font-bold text-white mb-0.5">1. 功能區 Ribbon【資料】標籤頁</div>
              <p class="text-slate-400 leading-relaxed text-xs">前處理的核心大本營！集中「移除重複項」、「資料剖析」與「排序篩選」。</p>
            </div>
            <div class="p-3 rounded-lg bg-slate-900/80 border-l-4 border-sky-500">
              <div class="font-bold text-white mb-0.5">2. 名稱方塊與公式編輯列</div>
              <p class="text-slate-400 leading-relaxed text-xs">精確核對提單字串長度，抓出藏在末端的不可見空格與換行控制字元。</p>
            </div>
            <div class="p-3 rounded-lg bg-slate-900/80 border-l-4 border-amber-500">
              <div class="font-bold text-white mb-0.5">3. 狀態列即時加總核算</div>
              <p class="text-slate-400 leading-relaxed text-xs">框選運費與毛重欄位時，右下角即時呈現「計數」與「數值合計」，快速核帳。</p>
            </div>
          </div>
        </div>
      </section>

      <!-- 幻燈片 3: 前處理基本常識與 GIGO 定律 -->
      <section class="slide-content p-6 sm:p-10 flex flex-col h-full hidden" data-index="3">
        <div class="mb-4">
          <span class="text-emerald-400 font-mono text-xs uppercase tracking-wider font-semibold">第二章 · 核心概念</span>
          <h2 class="text-xl sm:text-2xl font-bold text-white mt-1">資料前處理之基本常識與 GIGO 定律</h2>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-3 gap-4 flex-1 items-stretch overflow-hidden">
          <div class="bg-slate-900/80 p-4 rounded-xl border border-red-900/40 flex flex-col justify-between">
            <div>
              <div class="w-8 h-8 rounded-lg bg-red-950/80 text-red-400 flex items-center justify-center font-bold text-sm mb-3">!</div>
              <h3 class="text-white font-bold text-sm mb-2">航運髒資料的危害</h3>
              <p class="text-slate-400 text-xs leading-relaxed mb-2">提單號碼多一個空格、英文大小寫混亂、日期格式多元 (如 2024/05/01 與 2024.05.02)。</p>
              <div class="text-red-400 text-xs font-semibold">後果：樞紐分析失真、重複請款、報關退單</div>
            </div>
            <div class="mt-3 p-2 bg-black/30 rounded border border-red-900/20 text-[11px] text-slate-400">
              範例：<code class="text-red-300">" Kaohsiung "</code> ≠ <code class="text-emerald-300">"Kaohsiung"</code>
            </div>
          </div>

          <div class="bg-slate-900/80 p-4 rounded-xl border border-amber-900/40 flex flex-col justify-between">
            <div>
              <div class="w-8 h-8 rounded-lg bg-amber-950/80 text-amber-400 flex items-center justify-center font-bold text-sm mb-3">⚓</div>
              <h3 class="text-white font-bold text-sm mb-2">GIGO 與 SOLAS VGM</h3>
              <p class="text-slate-400 text-xs leading-relaxed mb-2">「垃圾進，垃圾出 (GIGO)」。毛重欄若混入文字 <code class="text-amber-300">N/A</code> 或單位 <code class="text-amber-300">kg</code>，加總會變成 0 或 <code class="text-red-400">#VALUE!</code>。</p>
              <div class="text-amber-400 text-xs font-semibold">衝擊：船舶配載重心配重失衡、違反國際法規遭扣船</div>
            </div>
            <div class="mt-3 p-2 bg-black/30 rounded border border-amber-900/20 text-[11px] text-slate-400">
              海事規範：國際海事組織 SOLAS 貨櫃總重驗證
            </div>
          </div>

          <div class="bg-slate-900/80 p-4 rounded-xl border border-emerald-900/40 flex flex-col justify-between">
            <div>
              <div class="w-8 h-8 rounded-lg bg-emerald-950/80 text-emerald-400 flex items-center justify-center font-bold text-sm mb-3">✓</div>
              <h3 class="text-white font-bold text-sm mb-2">Tidy Data 乾淨六守則</h3>
              <ul class="text-slate-400 text-xs space-y-1.5 leading-relaxed">
                <li>• <b>絕不使用合併儲存格</b> (分析大忌)</li>
                <li>• <b>首列維持單一標題列</b> (Header)</li>
                <li>• <b>一列一訂艙實體</b>，一欄單一屬性</li>
                <li>• <b>底層資料不留人工小計列</b></li>
                <li>• <b>杜絕排版用空白空列</b></li>
              </ul>
            </div>
            <div class="mt-3 p-2 bg-black/30 rounded border border-emerald-900/20 text-[11px] text-emerald-400 font-semibold">
              達成標準化結構，利於隨時樞紐分析
            </div>
          </div>
        </div>
      </section>

      <!-- 幻燈片 4: 四大前處理函數全解 -->
      <section class="slide-content p-6 sm:p-10 flex flex-col h-full hidden" data-index="4">
        <div class="flex items-center justify-between mb-4">
          <div>
            <span class="text-emerald-400 font-mono text-xs uppercase tracking-wider font-semibold">第二章 / 第三章 · 函數精華</span>
            <h2 class="text-xl sm:text-2xl font-bold text-white mt-1">航運文字與數據清洗四大核心函數</h2>
          </div>
          <span class="text-xs text-slate-400 bg-slate-900 border border-slate-800 px-2.5 py-1 rounded">圖 3-3 架構</span>
        </div>
        <div class="grid grid-cols-1 lg:grid-cols-12 gap-6 flex-1 items-center overflow-hidden">
          <div class="lg:col-span-7 h-full flex flex-col justify-center">
            <div class="rounded-xl overflow-hidden border border-slate-800 bg-black/40 shadow-xl">
              <img src="{img_formula}" alt="四大清洗函數" class="w-full object-contain max-h-[46vh] rounded-lg">
            </div>
          </div>
          <div class="lg:col-span-5 space-y-2.5 text-xs">
            <div class="p-2.5 rounded-lg bg-slate-900/80 border border-slate-800">
              <div class="font-mono text-emerald-400 font-bold">=TRIM( text )</div>
              <div class="text-slate-300 font-medium">去除頭尾空白字元</div>
              <div class="text-slate-500 text-[11px]">將 <code class="text-slate-300">"  Kaohsiung  "</code> 轉為純淨港口名稱。</div>
            </div>
            <div class="p-2.5 rounded-lg bg-slate-900/80 border border-slate-800">
              <div class="font-mono text-sky-400 font-bold">=CLEAN( text )</div>
              <div class="text-slate-300 font-medium">消除 EDI / 報關系統匯出不可見控制碼</div>
              <div class="text-slate-500 text-[11px]">剔除 ASCII 0~31 的非列印換行字符。</div>
            </div>
            <div class="p-2.5 rounded-lg bg-slate-900/80 border border-slate-800">
              <div class="font-mono text-amber-400 font-bold">=VALUE( text )</div>
              <div class="text-slate-300 font-medium">文字型數值轉換為計算數值</div>
              <div class="text-slate-500 text-[11px]">將字串 <code class="text-slate-300">" 18500 "</code> 轉為能計算的數字 18500。</div>
            </div>
            <div class="p-2.5 rounded-lg bg-slate-900/80 border border-slate-800">
              <div class="font-mono text-purple-400 font-bold">=SUBSTITUTE( text, "$", "" )</div>
              <div class="text-slate-300 font-medium">剔除海運費幣別符號與千分位逗點</div>
              <div class="text-slate-500 text-[11px]">運費欄位淨化必備嵌套函數。</div>
            </div>
          </div>
        </div>
      </section>

      <!-- 幻燈片 5: 30筆示範檔實作與去重複操作 -->
      <section class="slide-content p-6 sm:p-10 flex flex-col h-full hidden" data-index="5">
        <div class="flex items-center justify-between mb-4">
          <div>
            <span class="text-emerald-400 font-mono text-xs uppercase tracking-wider font-semibold">第三章 · 示範檔實作</span>
            <h2 class="text-xl sm:text-2xl font-bold text-white mt-1">提單主鍵去重 (Remove Duplicates)</h2>
          </div>
          <span class="text-xs text-slate-400 bg-slate-900 border border-slate-800 px-2.5 py-1 rounded">圖 3-1 實務對話框</span>
        </div>
        <div class="grid grid-cols-1 lg:grid-cols-12 gap-6 flex-1 items-center overflow-hidden">
          <div class="lg:col-span-7 h-full flex flex-col justify-center">
            <div class="rounded-xl overflow-hidden border border-slate-800 bg-black/40 shadow-xl">
              <img src="{img_dup}" alt="去重複介面" class="w-full object-contain max-h-[46vh] rounded-lg">
            </div>
          </div>
          <div class="lg:col-span-5 space-y-3 text-xs sm:text-sm">
            <div class="bg-slate-900/90 p-4 rounded-xl border border-slate-800 space-y-2">
              <div class="font-bold text-emerald-400 flex items-center gap-2">
                <span>⚡ 4 步驟標準操作程序 (SOP)</span>
              </div>
              <ol class="space-y-2 text-slate-300 text-xs">
                <li><b class="text-white">步驟 1：</b>選取資料範圍任一儲存格。</li>
                <li><b class="text-white">步驟 2：</b>切換至功能區<b>【資料】</b>標籤頁 -> 點選<b>【移除重複項】</b>。</li>
                <li><b class="text-white">步驟 3：</b>勾選「我的資料有標題」，點選「取消全選」後，單獨勾選<b>【提單號碼(BL_NO)】</b>與<b>【訂艙代號】</b>。</li>
                <li><b class="text-white">步驟 4：</b>點選「確定」，系統即刻自動剔除第 4 筆重複提單（保留 29 筆唯一值）。</li>
              </ol>
            </div>
            <div class="p-3 bg-amber-950/30 border border-amber-800/40 rounded-xl text-amber-300 text-xs">
              <b>航運提示：</b>切勿未經確認便全選所有欄位去重，因同一客戶可能分批寄送相同重量貨物，提單主鍵 (BL) 才是唯一判斷依據！
            </div>
          </div>
        </div>
      </section>

      <!-- 幻燈片 6: 資料剖析精靈與航商代碼拆解解密 -->
      <section class="slide-content p-6 sm:p-10 flex flex-col h-full hidden" data-index="6">
        <div class="flex items-center justify-between mb-4">
          <div>
            <span class="text-emerald-400 font-mono text-xs uppercase tracking-wider font-semibold">第三章 / 第四章 · 結構拆解</span>
            <h2 class="text-xl sm:text-2xl font-bold text-white mt-1">資料剖析 (Text to Columns) 拆解航商代碼</h2>
          </div>
          <span class="text-xs text-slate-400 bg-slate-900 border border-slate-800 px-2.5 py-1 rounded">圖 3-2 實務對話框</span>
        </div>
        <div class="grid grid-cols-1 lg:grid-cols-12 gap-6 flex-1 items-center overflow-hidden">
          <div class="lg:col-span-7 h-full flex flex-col justify-center">
            <div class="rounded-xl overflow-hidden border border-slate-800 bg-black/40 shadow-xl">
              <img src="{img_split}" alt="資料剖析精靈" class="w-full object-contain max-h-[46vh] rounded-lg">
            </div>
          </div>
          <div class="lg:col-span-5 space-y-3 text-xs sm:text-sm">
            <div class="bg-slate-900/90 p-4 rounded-xl border border-slate-800 space-y-2">
              <div class="font-bold text-sky-400">提單代碼拆解應用場景</div>
              <p class="text-slate-300 text-xs leading-relaxed">
                航運提單格式如 <code class="text-emerald-300">MSKU-8291034</code>，透過設定分隔符號為「其他: <code class="text-emerald-300">-</code>」，可一秒拆分出航商專屬代碼 (MSKU) 與流水號。
              </p>
              <div class="pt-2 border-t border-slate-800">
                <span class="text-slate-400 text-xs block mb-1">對應海運承攬實務好處：</span>
                <span class="text-teal-300 text-xs">迅速統計各航商（長榮、萬海、陽明、馬士基）年度訂艙櫃量與運費支出！</span>
              </div>
            </div>
            <div class="p-3 bg-slate-900/60 border border-slate-800 rounded-xl text-slate-400 text-xs">
              亦可搭配文字函數替代方案：<br/>
              <code class="text-sky-300">=LEFT(A2, 4)</code> 提取前綴 / <code class="text-sky-300">=MID(A2, 6, 7)</code> 提取號碼
            </div>
          </div>
        </div>
      </section>

      <!-- 幻燈片 7: 示範檔前處理前後成果看板 -->
      <section class="slide-content p-6 sm:p-10 flex flex-col h-full hidden" data-index="7">
        <div class="flex items-center justify-between mb-4">
          <div>
            <span class="text-emerald-400 font-mono text-xs uppercase tracking-wider font-semibold">第三章 · 成果驗證</span>
            <h2 class="text-xl sm:text-2xl font-bold text-white mt-1">資料前處理前後 (Before vs. After) 成果對照</h2>
          </div>
          <span class="text-xs text-slate-400 bg-slate-900 border border-slate-800 px-2.5 py-1 rounded">圖 3-4 品質看板</span>
        </div>
        <div class="grid grid-cols-1 lg:grid-cols-12 gap-6 flex-1 items-center overflow-hidden">
          <div class="lg:col-span-7 h-full flex flex-col justify-center">
            <div class="rounded-xl overflow-hidden border border-slate-800 bg-black/40 shadow-xl">
              <img src="{img_comp}" alt="Before vs After 對比" class="w-full object-contain max-h-[46vh] rounded-lg">
            </div>
          </div>
          <div class="lg:col-span-5 space-y-3 text-xs sm:text-sm">
            <div class="p-3 bg-red-950/20 border border-red-900/30 rounded-xl">
              <div class="font-bold text-red-400 mb-1">處理前 (Raw Data)</div>
              <ul class="text-slate-400 text-xs space-y-1">
                <li>• 提單頭尾空格、大小寫不一</li>
                <li>• 存在重複提單單號</li>
                <li>• 重量含有 N/A，海運費帶 $ 符號</li>
              </ul>
            </div>
            <div class="p-3 bg-emerald-950/20 border border-emerald-900/30 rounded-xl">
              <div class="font-bold text-emerald-400 mb-1">處理後 (Clean Data)</div>
              <ul class="text-slate-300 text-xs space-y-1">
                <li>• 100% 格式正規化、提單大寫標準碼</li>
                <li>• 完成唯一主鍵去重</li>
                <li>• 淨化為純數值，支援樞紐分析表匯總</li>
              </ul>
            </div>
          </div>
        </div>
      </section>

      <!-- 幻燈片 8: 300筆實戰練習題與情境設定 -->
      <section class="slide-content p-6 sm:p-10 flex flex-col h-full hidden" data-index="8">
        <div class="mb-4">
          <span class="text-emerald-400 font-mono text-xs uppercase tracking-wider font-semibold">第四章 · 實戰練習檔</span>
          <h2 class="text-xl sm:text-2xl font-bold text-white mt-1">300筆航運案例進階實戰練習庫 (Shipping_Practice_300)</h2>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-3 gap-4 flex-1 items-stretch overflow-hidden">
          <div class="bg-slate-900/80 p-4 rounded-xl border border-slate-800 flex flex-col justify-between">
            <div>
              <div class="text-emerald-400 font-mono text-xs font-bold mb-2">挑戰任務 01</div>
              <h3 class="text-white font-bold text-sm mb-2">提單與航商維度剖析</h3>
              <p class="text-slate-400 text-xs leading-relaxed">
                利用「資料剖析精靈」或文字函數，將提單字軌（如 MSKU-8291034）拆分出「航商代號」與「7碼流水號」。
              </p>
            </div>
            <div class="bg-black/30 p-2.5 rounded border border-slate-800 text-[11px] text-teal-300 font-mono">
              =LEFT(A2, 4) / =RIGHT(A2, 7)
            </div>
          </div>

          <div class="bg-slate-900/80 p-4 rounded-xl border border-slate-800 flex flex-col justify-between">
            <div>
              <div class="text-sky-400 font-mono text-xs font-bold mb-2">挑戰任務 02</div>
              <h3 class="text-white font-bold text-sm mb-2">全表去空格與去重審核</h3>
              <p class="text-slate-400 text-xs leading-relaxed">
                清洗出發港/目的港多餘空格，找出全表埋入的 <b>5 筆重複訂艙提單</b>並執行剔除，讓總數由 305 筆收斂為 300 筆。
              </p>
            </div>
            <div class="bg-black/30 p-2.5 rounded border border-slate-800 text-[11px] text-sky-300 font-mono">
              =PROPER(TRIM(D2)) + 去重複項
            </div>
          </div>

          <div class="bg-slate-900/80 p-4 rounded-xl border border-slate-800 flex flex-col justify-between">
            <div>
              <div class="text-purple-400 font-mono text-xs font-bold mb-2">挑戰任務 03</div>
              <h3 class="text-white font-bold text-sm mb-2">金額數值化與超重警示</h3>
              <p class="text-slate-400 text-xs leading-relaxed">
                將帶有 $ 符號之海運費純數值化；針對貨物重量 > 22,000 KG 之「超重櫃」設定條件式格式醒目提示。
              </p>
            </div>
            <div class="bg-black/30 p-2.5 rounded border border-slate-800 text-[11px] text-purple-300 font-mono">
              =VALUE(SUBSTITUTE(TRIM(I2),"$",""))
            </div>
          </div>
        </div>
      </section>

      <!-- 幻燈片 9: 300筆實戰題標準解題步驟 (SOP) -->
      <section class="slide-content p-6 sm:p-10 flex flex-col h-full hidden" data-index="9">
        <div class="mb-4">
          <span class="text-emerald-400 font-mono text-xs uppercase tracking-wider font-semibold">第四章 · 參考答案</span>
          <h2 class="text-xl sm:text-2xl font-bold text-white mt-1">實戰挑戰題參考標準解題步驟 (SOP)</h2>
        </div>
        <div class="space-y-3 flex-1 overflow-auto pr-2">
          <div class="p-3.5 bg-slate-900/80 rounded-xl border border-slate-800">
            <span class="text-xs font-bold text-emerald-400 mb-1 block">任務一解題：資料剖析極速切分</span>
            <p class="text-xs text-slate-300 leading-relaxed">
              選取提單欄 -> 點選【資料】->【資料剖析】-> 選取「分隔符號」-> 勾選「其他」並輸入破折號「<code class="text-emerald-300">-</code>」-> 目的地指定為相鄰兩空欄 -> 點擊完成，瞬間拆分 300 筆資料！
            </p>
          </div>
          <div class="p-3.5 bg-slate-900/80 rounded-xl border border-slate-800">
            <span class="text-xs font-bold text-sky-400 mb-1 block">任務二解題：空格淨化與主鍵剔除</span>
            <p class="text-xs text-slate-300 leading-relaxed">
              新增輔助欄輸入 <code class="text-sky-300">=PROPER(TRIM(D2))</code>，複製並「貼上值」覆蓋；隨後選取全表按【移除重複項】，勾選提單號碼為主鍵，精準刪除 5 筆重複單。
            </p>
          </div>
          <div class="p-3.5 bg-slate-900/80 rounded-xl border border-slate-800">
            <span class="text-xs font-bold text-purple-400 mb-1 block">任務三解題：運費數值化與條件式格式化超重標註</span>
            <p class="text-xs text-slate-300 leading-relaxed">
              套用嵌套公式 <code class="text-purple-300">=VALUE(SUBSTITUTE(SUBSTITUTE(TRIM(I2),"$",""),",",""))</code>；選取毛重欄 ->【設定格式化的條件】->【反白標示儲存格規則】->【大於 22000】，套用黃底紅字警示重櫃。
            </p>
          </div>
        </div>
      </section>

      <!-- 幻燈片 10: 參考文獻與課後資源 -->
      <section class="slide-content p-6 sm:p-10 flex flex-col justify-between h-full hidden" data-index="10">
        <div>
          <span class="text-emerald-400 font-mono text-xs uppercase tracking-wider font-semibold">第五章 · 權威來源</span>
          <h2 class="text-xl sm:text-2xl font-bold text-white mt-1 mb-4">參考文獻規範與課後學習資源</h2>
          <div class="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
            <div class="p-3 bg-slate-900/60 rounded-lg border border-slate-800">
              <div class="text-slate-200 font-bold mb-1">1. IMO SOLAS VGM 作業指引</div>
              <div class="text-slate-400 text-[11px]">國際海事組織貨櫃載重驗證標準規範 (MSC.1/Circ.1475)</div>
            </div>
            <div class="p-3 bg-slate-900/60 rounded-lg border border-slate-800">
              <div class="text-slate-200 font-bold mb-1">2. 國際貨櫃局 (BIC) ISO 6346</div>
              <div class="text-slate-400 text-[11px]">國際貨櫃標準編碼、船商字軌與校驗碼演算法</div>
            </div>
            <div class="p-3 bg-slate-900/60 rounded-lg border border-slate-800">
              <div class="text-slate-200 font-bold mb-1">3. 數位貨櫃航運協會 (DCSA)</div>
              <div class="text-slate-400 text-[11px]">DCSA Track & Trace 與船舶船期表數據介面標準</div>
            </div>
            <div class="p-3 bg-slate-900/60 rounded-lg border border-slate-800">
              <div class="text-slate-200 font-bold mb-1">4. Hadley Wickham - Tidy Data</div>
              <div class="text-slate-400 text-[11px]">資料科學整潔資料六大原則 (JSS, 2014)</div>
            </div>
          </div>
        </div>

        <div class="pt-4 border-t border-slate-800 flex flex-col sm:flex-row items-center justify-between gap-4">
          <div class="text-xs text-slate-400 text-center sm:text-left">
            恭喜您完成航運行政 Excel 資料前處理全部課程！<br/>
            配套檔案已完整編排至 <span class="text-emerald-400">航運行政人員Excel資料前處理教學手冊.docx</span>
          </div>
          <button onclick="goToSlide(1)" class="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-semibold transition">
            回到封面重新檢閱 ↺
          </button>
        </div>
      </section>

    </div>

  </main>

  <!-- 底部投影片巡覽導覽列 -->
  <footer class="h-16 border-t border-slate-800 bg-slate-950/90 px-6 flex items-center justify-between shrink-0 z-20">
    <div class="flex items-center gap-2">
      <button onclick="prevSlide()" class="px-3.5 py-1.5 rounded-lg bg-slate-900 hover:bg-slate-800 border border-slate-700 text-slate-200 text-xs font-medium transition flex items-center gap-1">
        <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 19l-7-7 7-7"/></svg>
        上一頁 (Prev)
      </button>
      <button onclick="nextSlide()" class="px-3.5 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold transition flex items-center gap-1 shadow-lg shadow-emerald-950">
        下一頁 (Next)
        <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5l7 7-7 7"/></svg>
      </button>
    </div>

    <!-- 投影片快速跳轉縮圖小圓點 -->
    <div class="hidden md:flex items-center gap-1.5" id="dotsContainer">
      <!-- 由 JS 動態產生 10 個小圓點 -->
    </div>

    <div class="text-[11px] text-slate-500">
      支援鍵盤 <kbd class="px-1.5 py-0.5 rounded bg-slate-900 border border-slate-800 text-slate-400 font-mono">←</kbd> <kbd class="px-1.5 py-0.5 rounded bg-slate-900 border border-slate-800 text-slate-400 font-mono">→</kbd> 或 <kbd class="px-1.5 py-0.5 rounded bg-slate-900 border border-slate-800 text-slate-400 font-mono">Space</kbd> 換頁
    </div>
  </footer>

  <script>
    let currentSlide = 1;
    const totalSlides = 10;
    const slides = document.querySelectorAll('.slide-content');
    const counter = document.getElementById('slideCounter');
    const progressBar = document.getElementById('progressBar');
    const dotsContainer = document.getElementById('dotsContainer');

    // 建立小圓點
    for (let i = 1; i <= totalSlides; i++) {{
      const dot = document.createElement('button');
      dot.className = `w-2.5 h-2.5 rounded-full transition-all duration-200 ${{i === 1 ? 'bg-emerald-400 w-6' : 'bg-slate-700 hover:bg-slate-500'}}`;
      dot.title = `跳至第 ${{i}} 頁`;
      dot.onclick = () => goToSlide(i);
      dotsContainer.appendChild(dot);
    }}

    function updateView() {{
      slides.forEach((slide) => {{
        const idx = parseInt(slide.getAttribute('data-index'));
        if (idx === currentSlide) {{
          slide.classList.remove('hidden');
          slide.classList.add('slide-enter');
        }} else {{
          slide.classList.add('hidden');
          slide.classList.remove('slide-enter');
        }}
      }});

      counter.textContent = `${{currentSlide < 10 ? '0' + currentSlide : currentSlide}} / ${{totalSlides < 10 ? '0' + totalSlides : totalSlides}}`;
      progressBar.style.width = `${{(currentSlide / totalSlides) * 100}}%`;

      // 更新圓點狀態
      const dots = dotsContainer.children;
      for (let i = 0; i < dots.length; i++) {{
        if (i + 1 === currentSlide) {{
          dots[i].className = 'w-6 h-2.5 rounded-full bg-emerald-400 transition-all duration-200';
        }} else {{
          dots[i].className = 'w-2.5 h-2.5 rounded-full bg-slate-700 hover:bg-slate-500 transition-all duration-200';
        }}
      }}
    }}

    function nextSlide() {{
      if (currentSlide < totalSlides) {{
        currentSlide++;
        updateView();
      }}
    }}

    function prevSlide() {{
      if (currentSlide > 1) {{
        currentSlide--;
        updateView();
      }}
    }}

    function goToSlide(n) {{
      if (n >= 1 && n <= totalSlides) {{
        currentSlide = n;
        updateView();
      }}
    }}

    function toggleFullscreen() {{
      if (!document.fullscreenElement) {{
        document.documentElement.requestFullscreen().catch(err => {{
          console.log(err.message);
        }});
      }} else {{
        if (document.exitFullscreen) {{
          document.exitFullscreen();
        }}
      }}
    }}

    // 鍵盤監聽
    window.addEventListener('keydown', (e) => {{
      if (e.key === 'ArrowRight' || e.key === 'PageDown' || e.key === ' ') {{
        e.preventDefault();
        nextSlide();
      }} else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {{
        e.preventDefault();
        prevSlide();
      }} else if (e.key === 'Home') {{
        e.preventDefault();
        goToSlide(1);
      }} else if (e.key === 'End') {{
        e.preventDefault();
        goToSlide(totalSlides);
      }}
    }});

    // 初始化
    updateView();
  </script>
</body>
</html>
'''

with open("航運行政Excel資料前處理教學簡報.html", "w", encoding="utf-8") as f:
    f.write(html_template)

print("Saved 航運行政Excel資料前處理教學簡報.html successfully!")
