import os
import random
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

os.makedirs("assets", exist_ok=True)
os.makedirs("data", exist_ok=True)

# 1. 建立 30 筆示範髒資料 (Shipping_Demo_Raw_30.xlsx)
raw_30_data = [
    # 提單號碼, 訂艙代號, 結關日期, 出發港(POL), 目的港(POD), 櫃號, 櫃型, 貨物重(KG), 運費(USD), 備註/狀態
    ["  MSKU-8291034 ", "BK001", "2024/05/01", " Kaohsiung ", "LAX", "TGHU1234567", "40HC", " 18,500 ", "$1,850", "準時"],
    ["MSKU-8291035", "BK002", "2024-05-02", "Keelung", "LGB ", "MSKU9876543", "20GP", "8200", "1200", " "],
    ["msku-8291036", "BK003", "2024.05.02", "Taichung", "RTM", "TEMU5544332", "40GP", "22400.5", "2,100", "危險品DG"],
    ["MSKU-8291034", "BK001", "2024/05/01", "Kaohsiung", "LAX", "TGHU1234567", "40HC", "18500", "1850", "重複重複"], # 完全重複/需去重
    ["OOLU-1102934", "BK004", "2024/05/03", "Kaohsiung", "HAM", "OOLU3322110", "40HC", "N/A", "2300", "待補重"], # 缺失/文字
    ["OOLU-1102935", "BK005", "2024/05/03", "Keelung", "SIN", "OOLU9988776", "20GP", "11000", "850", "OK"],
    ["EMCU-9923811", "BK006", "20240504", "Kaohsiung", "NYC", "EMCU1122334", "40HC", "19800", "3100", "特急件"],
    ["EMCU-9923812", "BK007", "2024/05/05", "Taichung", "HPH", "EMCU5566778", "20GP", "6500", "550", " "],
    ["CMAU-4402911", "BK008", "2024/05/06", "Keelung", "OAK", "CMAU2233445", "40HC", "21000", "2250", " "],
    ["CMAU-4402912", "BK009", "2024/05/06", "Kaohsiung", "BKK", "CMAU8899001", "20GP", "9300", "650", "待放行"],
    ["ONEU-7721029", "BK010", "2024/05/07", "Kaohsiung", "PUS", "ONEU1212121", "20GP", "14500", "450", "OK"],
    ["ONEU-7721030", "BK011", "2024-05-08", "Keelung", "TYO", "ONEU3434343", "40HC", "16200", "1100", " "],
    ["YMLU-5510293", "BK012", "2024/05/08", "Taichung", "DXB", "YMLU9090909", "40GP", "17800", "1950", "中轉卡達"],
    ["YMLU-5510294", "BK013", "2024/05/09", "Kaohsiung", "PKG", "YMLU8181818", "20GP", "10200", "700", " "],
    ["WHLC-3391024", "BK014", "2024/05/10", "Keelung", "MNL", "WHLU4545454", "20GP", "7800", "500", "電放提單"],
    ["WHLC-3391025", "BK015", "2024/05/10", "Taichung", "JKT", "WHLU6767676", "40HC", "23000", "1300", " "],
    ["  MSKU-8291040", "BK016", "2024/05/11", "Kaohsiung", "SHA", "MSKU7711223", "20GP", "12500", "350", " "],
    ["MSKU-8291041", "BK017", "2024/05/12", "Keelung", "SZX", "MSKU8822334", "40HC", "19100", "600", "急單"],
    ["OOLU-1102940", "BK018", "2024-05-12", "Kaohsiung", "MEL", "OOLU4455667", "40HC", "20500", "2100", "冷凍代用櫃"],
    ["OOLU-1102941", "BK019", "2024/05/13", "Taichung", "SYD", "OOLU7788990", "20GP", "8900", "1500", " "],
    ["EMCU-9923820", "BK020", "2024/05/14", "Kaohsiung", "SAV", "EMCU9900112", "40HC", "24500", "3400", "超重加價"],
    ["EMCU-9923821", "BK021", "2024/05/14", "Keelung", "CHS", "EMCU2233441", "40HC", "18200", "3250", " "],
    ["CMAU-4402920", "BK022", "2024.05.15", "Kaohsiung", "LEH", "CMAU5566771", "20GP", "11800", "1750", " "],
    ["CMAU-4402921", "BK023", "2024/05/15", "Taichung", "ZEE", "CMAU8899002", "40GP", "16500", "2050", " "],
    ["ONEU-7721035", "BK024", "2024/05/16", "Keelung", "KOB", "ONEU9911223", "20GP", "9500", "480", " "],
    ["ONEU-7721036", "BK025", "2024/05/16", "Kaohsiung", "YOK", "ONEU4433221", "40HC", "17300", "980", " "],
    ["YMLU-5510300", "BK026", "2024-05-17", "Taichung", "JED", "YMLU1155990", "40HC", "21500", "2400", " "],
    ["YMLU-5510301", "BK027", "2024/05/18", "Kaohsiung", "AQB", "YMLU2266001", "20GP", "13000", "1600", " "],
    ["WHLC-3391030", "BK028", "2024/05/19", "Keelung", "SUB", "WHLU3377112", "20GP", "10500", "800", " "],
    ["WHLC-3391031", "BK029", "2024/05/20", "Kaohsiung", "PEN", "WHLU4488223", "40HC", "19000", "1250", " "]
]

cols_30 = ["提單號碼(BL_NO)", "訂艙代號(Booking_ID)", "結關日期(Closing_Date)", "出發港(POL)", "目的港(POD)", "櫃號(Container_No)", "櫃型(Type)", "貨物毛重_KG(Weight)", "海運費_USD(Freight)", "狀態備註(Remark)"]
df_30 = pd.DataFrame(raw_30_data, columns=cols_30)
df_30.to_excel("data/Shipping_Demo_Raw_30.xlsx", index=False)
print("Saved data/Shipping_Demo_Raw_30.xlsx successfully!")

# 2. 建立 300 筆航運綜合實戰練習檔 (Shipping_Practice_300.xlsx)
random.seed(42)
carriers = ["MSKU", "OOLU", "EMCU", "CMAU", "ONEU", "YMLU", "WHLC", "HMMU"]
pols = ["Kaohsiung", "Keelung", "Taichung", "Taipei"]
pods = ["LAX", "LGB", "HAM", "RTM", "SIN", "NYC", "OAK", "BKK", "PUS", "TYO", "DXB", "PKG", "MNL", "JKT", "SHA", "SZX", "MEL", "SYD", "SAV", "CHS"]
ctypes = ["20GP", "40GP", "40HC", "45HC", "20RF", "40RF"]
remarks_pool = ["", "", "", "急單優先", "危險品DG第9類", "中轉貨載", "電放提單(Surrendered)", "待補領重單", "海關查驗(C3)", "冷凍控溫-18C", " "]

data_300 = []
for i in range(1, 301):
    c_prefix = random.choice(carriers)
    bl_no = f"{c_prefix}-{random.randint(1000000, 9999999)}"
    # 刻意製造髒數據：前綴空格、小寫、特殊格式
    r_type = random.random()
    if r_type < 0.1:
        bl_no = f"  {bl_no} "
    elif r_type < 0.18:
        bl_no = bl_no.lower()
        
    bk_id = f"BK{2024000+i}"
    
    # 日期雜亂格式
    m = random.randint(1, 12)
    d = random.randint(1, 28)
    date_style = random.choice(["slash", "dash", "dot", "compact"])
    if date_style == "slash":
        closing_date = f"2024/{m:02d}/{d:02d}"
    elif date_style == "dash":
        closing_date = f"2024-{m:02d}-{d:02d}"
    elif date_style == "dot":
        closing_date = f"2024.{m:02d}.{d:02d}"
    else:
        closing_date = f"2024{m:02d}{d:02d}"
        
    pol = random.choice(pols)
    if random.random() < 0.15:
        pol = f" {pol}  " # 夾雜前後空格
        
    pod = random.choice(pods)
    if random.random() < 0.1:
        pod = f"{pod} "
        
    cntr_no = f"{c_prefix}{random.randint(1000000, 9999999)}"
    ctype = random.choice(ctypes)
    
    # 重量：有些帶千分位逗號，有些帶小數點，有些為空或NA
    w_base = random.randint(4000, 26000)
    w_prob = random.random()
    if w_prob < 0.05:
        weight = "N/A"
    elif w_prob < 0.1:
        weight = ""
    elif w_prob < 0.35:
        weight = f" {w_base:,} "
    else:
        weight = str(w_base)
        
    # 運費：有些帶貨幣符號 $，有些有千分位逗號
    f_base = random.randint(350, 4500)
    f_prob = random.random()
    if f_prob < 0.2:
        freight = f"${f_base:,}"
    elif f_prob < 0.4:
        freight = f"{f_base:,}"
    elif f_prob < 0.45:
        freight = f" {f_base} "
    else:
        freight = str(f_base)
        
    remark = random.choice(remarks_pool)
    
    data_300.append([bl_no, bk_id, closing_date, pol, pod, cntr_no, ctype, weight, freight, remark])

# 注入重複資料 (約5筆完全重複)
for dup_idx in [10, 45, 92, 160, 230]:
    data_300.append(list(data_300[dup_idx]))

df_300 = pd.DataFrame(data_300, columns=cols_30)
df_300.to_excel("data/Shipping_Practice_300.xlsx", index=False)
print("Saved data/Shipping_Practice_300.xlsx successfully! Total rows:", len(df_300))
