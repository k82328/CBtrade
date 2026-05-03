#!/usr/bin/env python3
"""
台灣可轉換公司債（CB）批次資料爬蟲
資料來源：
  - 投資少數派 (thefew.tw)       → 發行資訊、到期日、餘額、轉換價等
  - 公開資訊觀測站 MOPS           → 停止轉換期間公告
  - TPEX 櫃買中心                 → CB 每日成交量
  - TDCC 集保結算所               → 剩餘張數（月報）

使用方式：
  pip install requests beautifulsoup4 pandas openpyxl lxml
  python cb_scraper.py

輸出：cb_detail_YYYYMMDD.xlsx
"""

import time
import re
import json
import requests
import pandas as pd
from datetime import datetime, date
from bs4 import BeautifulSoup

# ── 設定 ────────────────────────────────────────────────
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
}
DELAY = 1.5   # 每次請求間隔秒數（避免被封）
TIMEOUT = 15

# ── CB 代號清單（來自您的 CSV）──────────────────────────
CB_CODES = [
    "11011","12561","13166","13382","140201","140202","14363","14364",
    "14381","14664","14722","14723","14743","15142","15261","15364",
    "15601","15864","15865","15894","15952","15984","16095","17813",
    "17863","19094","20344","20662","21043","22013","22283","22331",
    "22362","22363","22471","22472","23291","23372","23383","23511",
    "23683","23742","24023","24271","24361","24394","24395","24423",
    "24424","24552","24553","24624","24642","24672","24673","24762",
    "24866","25283","25284","25303","25304","25483","25484","26107",
    "26417","26418","27271","27323","27324","27431","27451","27531",
    "27532","27551","27561","29061","30061","30112","30134","30165",
    "30321","30322","30336","30337","30371","30401","30454","30455",
    "30882","30922","30954","31222","31282","31311","31312","31351",
    "31382","31413","31672","31882","32071","32571","32608","32723",
    "32843","33035","33055","33121","33133","33225","33233","33245",
    "33246","33247","33466","33571","33621","33763","33881","33902",
    "34163","34793","34834","34843","35161","35181","35221","35265",
    "35266","35267","35431","35482","35483","35513","35642","35831",
    "35832","35871","35872","35914","36053","36173","36804","36841",
    "36872","36873","37153","37221","38001","38002","38011","38012",
    "38151","38241","38244","38391","38601","38621","38631","38651",
    "38681","38861","38862","38891","38921","39011","39101","39121",
    "39151","39201","39211","39231","39321","39361","39371","41011",
    "41031","42221","42312","43011","43063","43211","43453","43561",
    "43901","44011","44061","44271","44361","44622","44671","44851",
    "44901","44921","44981","45071","45201","45221","45311","45331",
    "45341","45361","45391","45401","45411","45421","45432","45441",
    "45521","45541","45551","45561","45651","45661","45671","45692",
    "45721","45771","45821","46061","46131","46231","46371","46381",
    "46421","46661","46671","46891","46941","46961","49211","49221",
    "49231","49471","50511","52131","54223","54581","54591","54801",
    "54921","55081","55101","55371","56101","56461","57011","57151",
    "57411","57511","57521","58111","58511","60041","60171","60531",
    "60561","60591","60651","60731","60761","60791","60841","60861",
    "61013","61043","61121","61131","61145","61153","61173","61261",
    "61263","61506","61565","61635","61777","61793","61794","61827",
    "61828","61843","61875","61894","61907","61973","61993","61994",
    "62075","62076","62093","62194","62433","62695","62696","62745",
    "62756","62822","62843","629010","64145","64251","64324","64422",
    "64423","64512","64693","64724","64773","64774","64775","65096",
    "65101","65152","65331","65381","65461","65472","65521","65761",
    "65841","65842","65851","65852","65914","66163","66241","66292",
    "66411","66451","66642","66702","66721","66722","66801","67062",
    "67152","67531","67681","67711","67961","67962","68041","68061",
    "68211","68231","68301","68351","68401","68431","68461","68462",
    "68541","68621","68631","68701","68702","68731","68732","68741",
    "68901","68902","68941","69041","69131","69571","69821","69822",
    "74022","75561","75562","77131","80212","80273","80282","80426",
    "80506","80872","80914","80924","81042","811210","81473","81551",
    "81714","81715","82102","82103","82551","82993","83742","84041",
    "84113","84221","84222","84312","84332","84333","84422","84623",
    "84662","84672","84673","84732","84781","84881","84891","89277",
    "89278","89964","98026","99063","99332","99353","99392","99412",
    "99587","99588",
]


# ════════════════════════════════════════════════════════
# 1. 投資少數派 thefew.tw — 主要發行資訊
# ════════════════════════════════════════════════════════
def fetch_thefew(cb_code: str) -> dict:
    """抓取投資少數派單檔 CB 詳細頁面"""
    url = f"https://thefew.tw/quote/{cb_code}"
    result = {"cb_code": cb_code, "source_thefew": url}
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "lxml")

        # 解析所有 <table> 內的 key-value 對
        for row in soup.find_all("tr"):
            cells = row.find_all("td")
            if len(cells) == 2:
                key = cells[0].get_text(strip=True)
                val = cells[1].get_text(strip=True)
                result[key] = val

    except Exception as e:
        result["thefew_error"] = str(e)
    return result


# ════════════════════════════════════════════════════════
# 2. TPEX 櫃買中心 — CB 每日成交量
#    https://www.tpex.org.tw/web/bond/CB/daily_trading/CB_quote.php
# ════════════════════════════════════════════════════════
def fetch_tpex_volume(cb_code: str, query_date: str = None) -> dict:
    """
    抓取單日 CB 成交量。
    query_date 格式：YYYY/MM/DD，預設為今天。
    """
    if query_date is None:
        query_date = date.today().strftime("%Y/%m/%d")

    url = "https://www.tpex.org.tw/web/bond/CB/daily_trading/CB_quote.php"
    params = {
        "l": "zh-tw",
        "d": query_date,
        "s": "0,asc",
    }
    result = {"cb_code": cb_code, "volume_date": query_date}
    try:
        r = requests.get(url, params=params, headers=HEADERS, timeout=TIMEOUT)
        r.raise_for_status()
        data = r.json()
        rows = data.get("aaData", [])
        for row in rows:
            # row[0] = CB代號, row[2] = 成交量(張)
            code = str(row[0]).strip()
            if code == cb_code:
                result["cb_volume_張"] = row[2].replace(",", "")
                result["cb_close_price"] = row[6].replace(",", "")
                result["cb_high"] = row[4].replace(",", "")
                result["cb_low"] = row[5].replace(",", "")
                result["cb_open"] = row[3].replace(",", "")
                break
        if "cb_volume_張" not in result:
            result["cb_volume_張"] = "無成交"
    except Exception as e:
        result["tpex_error"] = str(e)
    return result


# ════════════════════════════════════════════════════════
# 3. MOPS 公開資訊觀測站 — 停止轉換期間公告
#    使用 MOPS 重大訊息查詢 API
# ════════════════════════════════════════════════════════
def fetch_mops_suspension(stock_code: str) -> dict:
    """
    查詢該股票最近一筆停止轉換公告。
    stock_code = CB代號前4碼（母公司股票代號）
    """
    url = "https://mopsov.twse.com.tw/mops/web/t05st01"
    payload = {
        "encodeURIComponent": "1",
        "step": "1",
        "firstin": "1",
        "off": "1",
        "keyword4": "",
        "code1": "",
        "TYPEK2": "",
        "checkbtn": "",
        "queryName": "co_id",
        "inpuType": "co_id",
        "TYPEK": "all",
        "isnew": "false",
        "co_id": stock_code,
        "year": "",
        "month": "",
        "b_date": "",
        "e_date": "",
        "encodeURIComponent": "1",
        "step": "2",
    }
    result = {"stock_code": stock_code}
    try:
        r = requests.post(url, data=payload, headers=HEADERS, timeout=TIMEOUT)
        soup = BeautifulSoup(r.text, "lxml")
        # 找「停止轉換」相關公告
        rows = soup.find_all("tr")
        suspension_records = []
        for row in rows:
            text = row.get_text()
            if "停止" in text and ("轉換" in text or "轉(交)換" in text):
                cells = row.find_all("td")
                if cells:
                    suspension_records.append(" | ".join(c.get_text(strip=True) for c in cells[:4]))
        if suspension_records:
            result["停止轉換公告"] = suspension_records[0]  # 最新一筆
        else:
            result["停止轉換公告"] = "查無近期停止轉換公告"
    except Exception as e:
        result["mops_error"] = str(e)
    return result


# ════════════════════════════════════════════════════════
# 4. TDCC 集保 — 可轉換公司債月分析（剩餘張數更精確來源）
#    注意：TDCC 需要 POST 帶月份參數，較複雜，此為示範框架
# ════════════════════════════════════════════════════════
def fetch_tdcc_balance(cb_code: str) -> dict:
    """
    嘗試從集保月報取得流通在外張數。
    實務上 TDCC 以月報 PDF 為主，此處改抓 thefew 的餘額欄位。
    """
    # thefew 已含「最新餘額(百萬)」，轉換為張數
    return {"cb_code": cb_code, "note": "餘額已由 thefew 提供（百萬元）"}


# ════════════════════════════════════════════════════════
# 主流程：整合所有來源
# ════════════════════════════════════════════════════════
def scrape_all(cb_codes: list, max_count: int = None) -> pd.DataFrame:
    codes = cb_codes if max_count is None else cb_codes[:max_count]
    total = len(codes)
    records = []

    # 一次性抓取今日全部 CB 成交量（TPEX 一個 API 就全部回來）
    print("📡 抓取 TPEX 今日全部 CB 成交量...")
    tpex_all = {}
    try:
        today_str = date.today().strftime("%Y/%m/%d")
        url = "https://www.tpex.org.tw/web/bond/CB/daily_trading/CB_quote.php"
        params = {"l": "zh-tw", "d": today_str, "s": "0,asc"}
        r = requests.get(url, params=params, headers=HEADERS, timeout=TIMEOUT)
        data = r.json()
        for row in data.get("aaData", []):
            code = str(row[0]).strip()
            tpex_all[code] = {
                "CB現價_TPEX": row[6].replace(",", "") if len(row) > 6 else "",
                "CB開盤": row[3].replace(",", "") if len(row) > 3 else "",
                "CB最高": row[4].replace(",", "") if len(row) > 4 else "",
                "CB最低": row[5].replace(",", "") if len(row) > 5 else "",
                "CB成交量_張": row[2].replace(",", "") if len(row) > 2 else "",
                "CB成交金額_千元": row[7].replace(",", "") if len(row) > 7 else "",
            }
        print(f"   ✅ 取得 {len(tpex_all)} 筆 CB 行情資料")
    except Exception as e:
        print(f"   ⚠️  TPEX 抓取失敗：{e}")

    # 逐一抓 thefew 發行資訊
    for i, cb_code in enumerate(codes, 1):
        print(f"[{i:3d}/{total}] 抓取 {cb_code}...", end=" ")
        row_data = {"CB代號": cb_code}

        # --- thefew 發行資訊 ---
        tf = fetch_thefew(cb_code)
        field_map = {
            "可轉債名稱":           "CB名稱",
            "轉換標的名稱":         "標的名稱",
            "上市櫃別":             "上市櫃別",
            "擔保銀行/TCRI信用評等": "TCRI評等",
            "最新 CB 收盤價":       "CB收盤價",
            "轉換價值":             "CB理論價",
            "CBAS 權利金（百元報價）": "CBAS權利金",
            "CBAS 折現率":          "CBAS折現率",
            "轉換溢價率":           "轉換溢價率",
            "最新股票收盤價":       "股票現價",
            "目前轉換價":           "目前轉換價",
            "發行時轉換價":         "發行時轉換價",
            "發行價格":             "發行價格",
            "發行總額(百萬)":       "發行總額_百萬",
            "最新餘額(百萬)":       "最新餘額_百萬",
            "轉換比例":             "已轉換比例",
            "發行日":               "發行日",
            "到期日":               "到期日",
            "到期賣回價格":         "到期賣回價",
            "下次提前賣回日":       "下次賣回日",
            "下次提前賣回價格":     "下次賣回價",
        }
        for src_key, dst_key in field_map.items():
            row_data[dst_key] = tf.get(src_key, "")

        # 計算剩餘張數（1張 = 10萬元）
        try:
            bal_m = float(row_data.get("最新餘額_百萬", 0) or 0)
            row_data["剩餘張數"] = int(bal_m * 10)  # 百萬 × 10 = 張（以10萬/張計）
        except:
            row_data["剩餘張數"] = ""

        # --- TPEX 成交量 ---
        tpex_info = tpex_all.get(cb_code, {})
        row_data["CB成交量_張"] = tpex_info.get("CB成交量_張", "無成交")
        row_data["CB成交金額_千元"] = tpex_info.get("CB成交金額_千元", "")
        row_data["CB開盤"] = tpex_info.get("CB開盤", "")
        row_data["CB最高"] = tpex_info.get("CB最高", "")
        row_data["CB最低"] = tpex_info.get("CB最低", "")

        # --- 錯誤記錄 ---
        if "thefew_error" in tf:
            row_data["備註"] = f"thefew錯誤: {tf['thefew_error']}"

        records.append(row_data)
        status = "✅" if "thefew_error" not in tf else "⚠️"
        print(status)

        time.sleep(DELAY)

    df = pd.DataFrame(records)

    # 欄位排序
    col_order = [
        "CB代號", "CB名稱", "標的名稱", "上市櫃別",
        "發行日", "到期日", "發行總額_百萬", "最新餘額_百萬", "剩餘張數",
        "CB成交量_張", "CB成交金額_千元", "CB開盤", "CB最高", "CB最低",
        "目前轉換價", "發行時轉換價", "發行價格",
        "CB收盤價", "CB理論價", "CBAS權利金", "CBAS折現率", "轉換溢價率",
        "股票現價", "已轉換比例", "TCRI評等",
        "到期賣回價", "下次賣回日", "下次賣回價",
        "備註",
    ]
    existing = [c for c in col_order if c in df.columns]
    extra = [c for c in df.columns if c not in col_order]
    df = df[existing + extra]

    return df


# ════════════════════════════════════════════════════════
# 停止轉換期間：單獨查詢函式（較慢，可選用）
# ════════════════════════════════════════════════════════
def fetch_suspension_batch(stock_codes: list) -> dict:
    """
    批次查詢停止轉換期間。
    stock_codes = CB代號前4碼（母公司股票代號）的去重清單
    回傳 dict: {stock_code: 公告內容}
    """
    results = {}
    for code in stock_codes:
        print(f"   查停止轉換 {code}...", end=" ")
        info = fetch_mops_suspension(code)
        results[code] = info.get("停止轉換公告", "")
        print("✅")
        time.sleep(DELAY)
    return results


# ════════════════════════════════════════════════════════
# 輸出 Excel
# ════════════════════════════════════════════════════════
def save_excel(df: pd.DataFrame, filepath: str):
    with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="CB詳細資料")
        ws = writer.sheets["CB詳細資料"]

        # 自動欄寬
        for col in ws.columns:
            max_len = max((len(str(cell.value)) for cell in col if cell.value), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)

        # 凍結首列
        ws.freeze_panes = "A2"

    print(f"\n✅ 已儲存：{filepath}")


# ════════════════════════════════════════════════════════
# 執行
# ════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("=" * 60)
    print("  台灣 CB 批次資料爬蟲")
    print(f"  共 {len(CB_CODES)} 支 CB，預計耗時約 {len(CB_CODES) * DELAY / 60:.0f} 分鐘")
    print("=" * 60)

    # 若只想測試，改成 max_count=5
    df = scrape_all(CB_CODES, max_count=None)

    today = datetime.today().strftime("%Y%m%d")
    out_path = f"cb_detail_{today}.xlsx"
    save_excel(df, out_path)

    print(f"\n📊 摘要：")
    print(f"   總筆數：{len(df)}")
    print(f"   有成交：{(df['CB成交量_張'] != '無成交').sum()} 筆")
    print(f"   資料欄位：{list(df.columns)}")
