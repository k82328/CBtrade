#!/usr/bin/env python3
"""
CB + 股票報價抓取工具
用法：
    pip install requests openpyxl pandas
    python fetch_prices.py           # 抓報價 + 更新 HTML
    python fetch_prices.py --show    # 抓完後顯示結果
    python fetch_prices.py --cb-only
    python fetch_prices.py --stock-only
"""

import requests, json, time, os, sys, argparse
from datetime import datetime

CB_EXCEL_FILE = "cb_list.xlsx"
OUTPUT_FILE   = "prices.json"
HTML_FILE     = "cb_tracker.html"
BATCH_SIZE    = 50

headers  = {"User-Agent": "Mozilla/5.0"}
BASE_URL = "https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch={}&json=1&delay=0"

INDUSTRY = {
  "1101":"水泥","1256":"食品","1316":"化學","1338":"化學",
  "1402":"紡織","1436":"紡織","1438":"紡織","1466":"紡織",
  "1472":"紡織","1474":"紡織","1514":"電機","1526":"汽車",
  "1536":"汽車","1560":"機械","1586":"塑膠","1589":"其他",
  "1595":"其他","1598":"運動休閒","1609":"電線電纜",
  "1786":"生技醫療","1909":"造紙",
  "2034":"鋼鐵","2066":"鋼鐵","2104":"橡膠","2201":"汽車",
  "2228":"汽車零件","2233":"汽車零件","2236":"汽車零件",
  "2247":"汽車","2329":"半導體","2337":"半導體","2338":"半導體",
  "2351":"電子零組件","2368":"半導體","2374":"電腦周邊",
  "2402":"電子零組件","2427":"電子通路","2436":"半導體",
  "2439":"電子零組件","2442":"電子零組件","2455":"半導體",
  "2462":"電子零組件","2464":"電子零組件","2467":"半導體設備",
  "2476":"電子零組件","2486":"電子零組件","2528":"建設",
  "2530":"建設","2548":"建設","2610":"航運","2641":"航運",
  "2727":"餐飲","2732":"餐飲","2743":"觀光","2745":"觀光",
  "2753":"餐飲","2755":"餐飲","2756":"餐飲","2906":"百貨",
  "3006":"半導體","3011":"其他電子","3013":"其他電子",
  "3016":"半導體","3032":"電子通路","3033":"電子通路",
  "3037":"PCB","3040":"其他電子","3045":"電信",
  "3088":"其他電子","3092":"電子零組件","3095":"其他電子",
  "3122":"半導體","3128":"其他電子","3131":"半導體設備",
  "3135":"半導體","3138":"電子零組件","3141":"半導體",
  "3167":"半導體設備","3188":"其他電子","3207":"其他電子",
  "3257":"其他電子","3260":"電腦周邊","3272":"其他電子",
  "3284":"其他電子","3303":"光電","3305":"電子通路",
  "3312":"其他電子","3313":"其他電子","3322":"其他電子",
  "3323":"電子零組件","3324":"散熱","3346":"光電",
  "3357":"其他電子","3362":"光電","3376":"電子零組件",
  "3388":"電子通路","3390":"其他電子","3416":"其他電子",
  "3479":"其他電子","3483":"其他電子","3484":"其他電子",
  "3516":"其他電子","3518":"其他電子","3522":"其他電子",
  "3526":"電子零組件","3543":"其他電子","3548":"其他電子",
  "3551":"半導體設備","3564":"其他電子","3583":"半導體設備",
  "3587":"半導體設備","3591":"光電","3605":"電子零組件",
  "3617":"電源供應","3680":"半導體設備","3684":"生技",
  "3687":"電競","3689":"其他電子","3691":"太陽能",
  "3701":"電子通路","3702":"電子通路","3707":"半導體",
  "3708":"再生能源","3715":"其他電子","3717":"電子通路",
  "4113":"生技醫療","4123":"生技醫療","4129":"生技醫療",
  "4137":"生技醫療","4164":"生技醫療","4168":"生技醫療",
  "4190":"生技醫療","4416":"建設","4438":"其他","4439":"其他",
  "4442":"其他","4510":"機械","4542":"機械","4549":"其他電子",
  "4555":"機械","4558":"其他電子","4564":"其他電子",
  "4566":"機械","4569":"其他","4572":"其他電子","4581":"其他",
  "4722":"化學","4739":"電子零組件","4747":"生技",
  "4772":"化學","4906":"通訊網路","4916":"通訊網路",
  "4923":"其他電子","4967":"電腦周邊","4979":"光電",
  "5009":"鋼鐵","5201":"通訊網路","5209":"工程顧問",
  "5212":"其他電子","5230":"光電","5244":"其他電子",
  "5245":"半導體","5263":"觀光","5306":"機械",
  "5309":"電子通路","5321":"其他","5388":"通訊網路",
  "5434":"電子通路","5439":"生技","5450":"紡織",
  "5457":"電子零組件","5534":"建設","5608":"航運",
  "5905":"觀光",
  "6015":"證券","6101":"餐飲","6104":"半導體",
  "6111":"光電","6117":"電源供應","6125":"機械",
  "6126":"電子零組件","6150":"顯示卡","6156":"其他電子",
  "6163":"通訊網路","6177":"建設","6179":"其他電子",
  "6182":"半導體","6184":"有線電視","6187":"自動化",
  "6189":"電子通路","6190":"通訊網路","6197":"電子零組件",
  "6199":"生技","6207":"其他電子","6209":"光電",
  "6219":"建設","6243":"半導體","6269":"PCB",
  "6274":"PCB","6275":"其他電子","6282":"電源供應",
  "6284":"電子零組件","6290":"電子零組件","6414":"工業電腦",
  "6425":"其他","6432":"光電","6442":"光電",
  "6451":"半導體","6469":"藥局","6472":"生技醫療",
  "6477":"其他電子","6509":"化學","6510":"半導體設備",
  "6515":"半導體設備","6533":"半導體","6538":"機械",
  "6546":"其他電子","6547":"生技","6552":"其他電子",
  "6576":"生技","6584":"其他","6585":"化學",
  "6591":"其他電子","6616":"其他","6624":"其他",
  "6629":"其他","6641":"其他","6645":"生技",
  "6664":"其他電子","6670":"機械","6672":"電子零組件",
  "6680":"其他電子","6706":"其他電子","6715":"電子零組件",
  "6753":"造船","6768":"其他","6771":"環保",
  "6796":"生技醫療","6804":"其他電子","6806":"再生能源",
  "6821":"其他電子","6823":"其他電子","6830":"半導體設備",
  "6835":"其他","6840":"其他電子","6843":"其他電子",
  "6846":"再生能源","6854":"光電","6862":"其他",
  "6863":"其他","6870":"其他電子","6873":"再生能源",
  "6874":"其他電子","6890":"其他","6894":"其他電子",
  "6904":"其他電子","6913":"其他電子","6957":"其他",
  "6982":"機械",
  "7402":"電子零組件","7556":"其他","7713":"生技",
  "8021":"電子零組件","8027":"半導體","8028":"半導體",
  "8042":"電源供應","8050":"工業電腦","8087":"再生能源",
  "8091":"其他電子","8092":"其他電子","8104":"光電",
  "8112":"電子通路","8147":"其他電子","8155":"其他電子",
  "8171":"其他電子","8210":"伺服器","8255":"電子零組件",
  "8299":"半導體","8374":"電子通路","8404":"紡織",
  "8411":"其他","8422":"環保","8431":"其他電子",
  "8433":"其他電子","8442":"其他","8462":"健身",
  "8466":"其他","8467":"運動休閒","8473":"環保",
  "8478":"遊艇","8488":"其他","8927":"食品",
  "8996":"機械","9802":"紡織","9906":"其他",
  "9933":"工程","9935":"其他","9939":"包材",
  "9941":"租賃","9958":"再生能源",
}

def get_price(item):
    z = item.get("z", "-")
    y = item.get("y", "-")
    pz = item.get("pz", "-"); h = item.get("h", "-"); val = z if (z and z != "-") else (pz if (pz and pz != "-") else (h if (h and h != "-" and h != "0") else y))
    try: return round(float(val), 2)
    except: return None

def fetch_batch(codes, prefix):
    query = "|".join(f"{prefix}_{c}.tw" for c in codes)
    try:
        r = requests.get(BASE_URL.format(query), headers=headers, timeout=10)
        result = {}
        for item in r.json().get("msgArray", []):
            code = item.get("c", "")
            price = get_price(item)
            if code and price:
                result[code] = price
        return result
    except Exception as e:
        print(f"  警告: {e}")
        return {}

def load_cb_excel():
    if not os.path.exists(CB_EXCEL_FILE):
        print(f"找不到 {CB_EXCEL_FILE}")
        sys.exit(1)
    try:
        import pandas as pd
        df = pd.read_excel(CB_EXCEL_FILE, dtype=str)
        df.columns = df.columns.str.strip()
        df = df.rename(columns={"CB代號":"cb_code","股票代號":"stock","名稱":"name","轉換價":"conv"})
        df = df.dropna(subset=["cb_code","stock"])
        df["cb_code"] = df["cb_code"].str.strip()
        df["stock"]   = df["stock"].str.strip()
        df["name"]    = df["name"].fillna("").str.strip()
        df["conv"]    = pd.to_numeric(df["conv"], errors="coerce")
        records = df.to_dict("records")
        print(f"✓ 讀取 {CB_EXCEL_FILE}：共 {len(records)} 筆")
        return records
    except Exception as e:
        print(f"讀取失敗：{e}")
        sys.exit(1)

def fetch_all(codes, label, try_tse=True):
    prices = {}
    total  = len(codes)
    print(f"\n【{label}】{total} 筆...")
    for i in range(0, total, BATCH_SIZE):
        prices.update(fetch_batch(codes[i:i+BATCH_SIZE], "otc"))
        print(f"  {min(i+BATCH_SIZE,total)}/{total}...", end="\r")
        time.sleep(0.3)
    if try_tse:
        missing = [c for c in codes if c not in prices]
        if missing:
            print(f"  otc={len(prices)}, 改試tse ({len(missing)}筆)...")
            for i in range(0, len(missing), BATCH_SIZE):
                prices.update(fetch_batch(missing[i:i+BATCH_SIZE], "tse"))
                time.sleep(0.3)
    failed = [c for c in codes if c not in prices]
    print(f"  完成！{len(prices)}/{total}" + (f"，失敗：{failed}" if failed else "   "))
    return prices

def generate_html(records, cb_prices, stock_prices, update_time):
    """產生完整 HTML，RAW 資料從 Excel 清單來"""
    # 產生 RAW 陣列
    raw_lines = []
    for r in records:
        cb_code = r["cb_code"]
        stock   = r["stock"]
        name    = r["name"]
        conv    = r["conv"]
        cb_p    = cb_prices.get(cb_code, "null")
        sp      = stock_prices.get(stock, "null")
        conv_js = conv if conv == conv else "null"  # NaN check
        raw_lines.append(f'["{cb_code}","{stock}","{name}",{conv_js},{cb_p},{sp}]')
    raw_js = "[\n" + ",\n".join(raw_lines) + "\n]"

    industry_js = json.dumps(INDUSTRY, ensure_ascii=False)

    html = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>CB 可轉債追蹤工具</title>
<style>
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; font-size: 13px; background: #f5f5f5; color: #222; }}
.top {{ background: #fff; border-bottom: 1px solid #ddd; padding: 12px 16px; display: flex; align-items: center; gap: 12px; flex-wrap: wrap; position: sticky; top: 0; z-index: 10; }}
.top h1 {{ font-size: 15px; font-weight: 600; margin-right: 4px; }}
.summary {{ display: flex; gap: 8px; flex-wrap: wrap; padding: 10px 16px; background: #f9f9f9; border-bottom: 1px solid #eee; }}
.stat {{ background: #fff; border: 1px solid #e0e0e0; border-radius: 6px; padding: 6px 14px; }}
.stat-l {{ font-size: 10px; color: #888; }}
.stat-v {{ font-size: 18px; font-weight: 600; }}
.toolbar {{ display: flex; gap: 8px; flex-wrap: wrap; padding: 10px 16px; align-items: center; border-bottom: 1px solid #eee; background: #fff; }}
input[type=text], select {{ padding: 5px 8px; border: 1px solid #ccc; border-radius: 5px; font-size: 12px; }}
button {{ padding: 5px 12px; border: 1px solid #bbb; border-radius: 5px; background: #fff; cursor: pointer; font-size: 12px; }}
button:hover {{ background: #f0f0f0; }}
button.primary {{ background: #1a73e8; color: #fff; border-color: #1a73e8; }}
button.primary:hover {{ background: #1558c0; }}
.tbl-wrap {{ overflow-x: auto; }}
table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
thead th {{ position: sticky; top: 0; background: #fafafa; border-bottom: 2px solid #ddd; padding: 7px 8px; text-align: right; font-weight: 600; font-size: 11px; color: #555; white-space: nowrap; }}
thead th:nth-child(-n+3) {{ text-align: left; }}
tbody tr:hover {{ background: #f0f6ff; }}
tbody tr:nth-child(even) {{ background: #fafafa; }}
tbody tr:nth-child(even):hover {{ background: #f0f6ff; }}
tbody tr.row-checked {{ background: #e4e4e4 !important; }}
tbody tr.row-checked:hover {{ background: #d8d8d8 !important; }}
input.note-in {{ background: none; border: none; border-bottom: 1px dashed #ccc; width: 100%; min-width: 90px; font-size: 12px; padding: 1px 2px; outline: none; color: #222; font-family: inherit; }}
input.note-in:focus {{ border-bottom-color: #1a73e8; background: #f0f6ff; }}
tbody tr.row-checked {{ background: #e4e4e4 !important; }}
tbody tr.row-checked:hover {{ background: #d8d8d8 !important; }}
input.note-in {{ background: none; border: none; border-bottom: 1px dashed #ccc; width: 100%; min-width: 90px; font-size: 12px; padding: 1px 2px; outline: none; color: #222; font-family: inherit; }}
tbody td {{ padding: 5px 8px; border-bottom: 1px solid #eee; text-align: right; }}
tbody td:nth-child(-n+3) {{ text-align: left; }}
.cb-name {{ font-weight: 500; }}
.badge {{ font-size: 11px; border-radius: 3px; padding: 1px 5px; display: inline-block; font-weight: 600; }}
.b-deep {{ background: #fde8e8; color: #c0392b; }}
.b-low  {{ background: #fef3cd; color: #856404; }}
.b-ok   {{ background: #d4edda; color: #155724; }}
.na     {{ color: #bbb; }}
.auto   {{ color: #1a73e8; font-weight: 500; }}
input.cell-in {{ background: none; border: 1px solid transparent; border-radius: 4px; width: 72px; text-align: right; font-size: 12px; padding: 3px 5px; outline: none; color: #222; transition: all .15s; }}
input.cell-in:hover {{ border-color: #ccc; background: #fff; }}
input.cell-in:focus {{ border-color: #1a73e8; background: #f0f6ff; box-shadow: 0 0 0 2px rgba(26,115,232,.15); }}
input.cell-in.has-val {{ color: #1a73e8; font-weight: 500; }}
.modal-bg {{ display: none; position: fixed; top:0;left:0;right:0;bottom:0; background: rgba(0,0,0,.4); z-index:100; align-items: center; justify-content: center; }}
.modal-bg.show {{ display: flex; }}
.modal {{ background: #fff; border-radius: 10px; padding: 24px; width: 480px; max-width: 95vw; }}
.modal h2 {{ font-size: 15px; margin-bottom: 12px; }}
.modal textarea {{ width: 100%; height: 200px; border: 1px solid #ccc; border-radius: 5px; padding: 8px; font-size: 12px; font-family: monospace; }}
.modal .actions {{ margin-top: 12px; display: flex; gap: 8px; justify-content: flex-end; }}
.status {{ font-size: 11px; color: #888; margin-left: auto; }}
</style>
</head>
<body>

<div class="top">
  <h1>CB 可轉債追蹤</h1>
  <button class="primary" onclick="openModal()">匯入報價 JSON</button>
  <button onclick="exportCSV()">匯出 CSV</button>
  <span id="lastUpdate" class="status">資料更新時間：{update_time}</span>
</div>

<div class="summary">
  <div class="stat"><div class="stat-l">總 CB 數</div><div class="stat-v" id="s-total">—</div></div>
  <div class="stat"><div class="stat-l">有股價</div><div class="stat-v" style="color:#1a73e8" id="s-price">—</div></div>
  <div class="stat"><div class="stat-l">有 CB 現價</div><div class="stat-v" id="s-cb">—</div></div>
  <div class="stat"><div class="stat-l">深折價 &lt;80%</div><div class="stat-v" style="color:#c0392b" id="s-deep">—</div></div>
  <div class="stat"><div class="stat-l">折價 80~100%</div><div class="stat-v" style="color:#856404" id="s-low">—</div></div>
  <div class="stat"><div class="stat-l">溢價 &gt;100%</div><div class="stat-v" style="color:#155724" id="s-prem">—</div></div>
</div>

<div class="toolbar">
  <input type="text" id="searchInput" placeholder="搜尋代碼/名稱..." style="width:160px" oninput="filterSearch=this.value;render()">
  <select onchange="filterPct=this.value;render()">
    <option value="all">全部溢價</option>
    <option value="deep">深折價 &lt;80%</option>
    <option value="low">折價 80–100%</option>
    <option value="prem">溢價 &gt;100%</option>
    <option value="noprice">無股價</option>
    <option value="hascb">有 CB 現價</option>
  </select>
  <select onchange="filterIndustry=this.value;render()" id="industrySelect">
    <option value="">全部產業</option>
  </select>
  <select onchange="sortBy=this.value;render()">
    <option value="">排序：預設</option>
    <option value="cbpct">CB/理論 升冪</option>
    <option value="cbpct_desc">CB/理論 降冪</option>
    <option value="spct">股/轉換 升冪</option>
    <option value="spct_desc">股/轉換 降冪</option>
  </select>
  <span id="countLabel" class="status"></span>
</div>

<div class="tbl-wrap">
<table>
<thead><tr>
  <th style="width:32px;text-align:center">✓</th>
  <th style="width:70px;text-align:left">CB代號</th>
  <th style="width:68px;text-align:left">股票</th>
  <th style="width:115px;text-align:left">名稱</th>
  <th style="width:80px;text-align:left">產業</th>
  <th style="width:64px">轉換價</th>
  <th style="width:76px">股票現價</th>
  <th style="width:76px">CB理論價</th>
  <th style="width:76px">CB現價</th>
  <th style="width:72px">CB/理論</th>
  <th style="width:70px">股/轉換</th>
  <th style="min-width:110px;text-align:left">備註</th>
</tr></thead>
<tbody id="tbody"></tbody>
</table>
</div>

<div class="modal-bg" id="modal">
  <div class="modal">
    <h2>匯入報價 JSON</h2>
    <p style="font-size:12px;color:#666;margin-bottom:8px">將 <code>prices.json</code> 內容貼入：</p>
    <textarea id="jsonInput" placeholder='{{"stock_prices": {{"1101": 24.4}}, "cb_prices": {{"11011": 99.0}}}}'></textarea>
    <div style="margin-top:10px;display:flex;gap:16px;font-size:12px;">
      <label><input type="radio" name="importMode" value="overwrite" checked> 全部覆蓋</label>
      <label><input type="radio" name="importMode" value="fill"> 只補空白</label>
    </div>
    <div class="actions">
      <button onclick="closeModal()">取消</button>
      <button class="primary" onclick="applyJSON()">套用</button>
    </div>
  </div>
</div>

<script>
const RAW = {raw_js};
const INDUSTRY = {industry_js};

let cbData = RAW.map(r => ({{
  cbCode: r[0], stock: r[1], name: r[2], conv: r[3],
  cbPrice: r[4] || null, stockPrice: r[5] || null, fetched: r[5]!=null
}}));

let filterPct='all', filterSearch='', sortBy='', filterIndustry='';

(function(){{
  const industries = [...new Set(Object.values(INDUSTRY))].sort();
  const sel = document.getElementById('industrySelect');
  industries.forEach(ind => {{
    const opt = document.createElement('option');
    opt.value = ind; opt.textContent = ind;
    sel.appendChild(opt);
  }});
}})();

function T(sp, cv) {{ return sp&&cv ? sp/cv*100 : null; }}
function fmt(v, d=2) {{ return v==null||isNaN(v)?'':v.toFixed(d); }}
function badge(r) {{
  if (!r) return '<span class="na">—</span>';
  const c = r<80?'b-deep':r<100?'b-low':'b-ok';
  return `<span class="badge ${{c}}">${{r.toFixed(1)}}%</span>`;
}}

function getFiltered() {{
  let arr = cbData.filter(d => {{
    if (filterIndustry && (INDUSTRY[d.stock]||'') !== filterIndustry) return false;
    const q = filterSearch.toLowerCase();
    if (q && !d.cbCode.includes(q) && !d.stock.includes(q) && !d.name.includes(q)) return false;
    const t = T(d.stockPrice, d.conv);
    const r = (d.cbPrice&&t) ? d.cbPrice/t*100 : null;
    if (filterPct==='noprice') return !d.stockPrice;
    if (filterPct==='hascb') return !!d.cbPrice;
    if (filterPct==='deep') return r!==null&&r<80;
    if (filterPct==='low') return r!==null&&r>=80&&r<100;
    if (filterPct==='prem') return r!==null&&r>=100;
    return true;
  }});
  if (sortBy) {{
    const asc = !sortBy.endsWith('_desc');
    const key = sortBy.replace('_desc','');
    arr.sort((a,b) => {{
      const ta=T(a.stockPrice,a.conv), tb=T(b.stockPrice,b.conv);
      let va=null,vb=null;
      if (key==='cbpct') {{ va=(a.cbPrice&&ta)?a.cbPrice/ta*100:null; vb=(b.cbPrice&&tb)?b.cbPrice/tb*100:null; }}
      if (key==='spct')  {{ va=(a.stockPrice&&a.conv)?a.stockPrice/a.conv*100:null; vb=(b.stockPrice&&b.conv)?b.stockPrice/b.conv*100:null; }}
      if (va==null&&vb==null) return 0;
      if (va==null) return 1; if (vb==null) return -1;
      return asc ? va-vb : vb-va;
    }});
  }}
  return arr;
}}


// localStorage：儲存打勾 & 備註
const LS_KEY = 'cb_user_meta';
let userMeta = {{}};
try {{ userMeta = JSON.parse(localStorage.getItem(LS_KEY) || '{{}}'); }} catch(e) {{}}
function setMeta(cbCode, field, val) {{
  if (!userMeta[cbCode]) userMeta[cbCode] = {{}};
  userMeta[cbCode][field] = val;
  try {{ localStorage.setItem(LS_KEY, JSON.stringify(userMeta)); }} catch(e) {{}}
}}

function render() {{
  const filtered = getFiltered();
  document.getElementById('tbody').innerHTML = filtered.map(d => {{
    const idx = cbData.indexOf(d);
    const t = T(d.stockPrice, d.conv);
    const cr = (d.cbPrice&&t)?d.cbPrice/t*100:null;
    const sr = (d.stockPrice&&d.conv)?d.stockPrice/d.conv*100:null;
    const spClass = d.stockPrice ? 'has-val' : '';
    const isChk = userMeta[d.cbCode]?.checked || false;
    const noteVal = (userMeta[d.cbCode]?.note || '').replace(/"/g,'&quot;');
    return `<tr data-idx="${{idx}}" class="${{isChk?'row-checked':''}}">
      <td style="text-align:center"><input type="checkbox" ${{isChk?'checked':''}} onchange="setMeta('${{d.cbCode}}','checked',this.checked);render()"></td>
      <td><input class="cell-in" type="text" value="${{d.cbCode}}" style="width:68px;color:#999;font-size:11px" onblur="cbData[${{idx}}].cbCode=this.value"></td>
      <td><input class="cell-in" type="text" value="${{d.stock}}" style="width:58px;color:#999;font-size:11px" onblur="cbData[${{idx}}].stock=this.value;refreshRow(${{idx}})"></td>
      <td class="cb-name">${{d.name}}</td>
      <td style="font-size:11px;color:#555">${{INDUSTRY[d.stock]||'—'}}</td>
      <td>${{fmt(d.conv)}}</td>
      <td><input class="cell-in ${{spClass}}" type="number" step="0.01" value="${{d.stockPrice??''}}" placeholder="—" onblur="upStock(${{idx}},this.value)"></td>
      <td style="color:#888">${{t?fmt(t):'<span class=na>—</span>'}}</td>
      <td><input class="cell-in" type="number" step="0.05" value="${{d.cbPrice??''}}" placeholder="—" onblur="upCB(${{idx}},this.value)"></td>
      <td>${{badge(cr)}}</td>
      <td>${{badge(sr)}}</td>
      <td><input class="note-in" type="text" value="${{noteVal}}" placeholder="備註…" onchange="setMeta('${{d.cbCode}}','note',this.value)"></td>
    </tr>`;
  }}).join('');
  document.getElementById('countLabel').textContent = `顯示 ${{filtered.length}} / ${{cbData.length}}`;
  updateSummary();
}}

function upStock(idx, val) {{
  cbData[idx].stockPrice = val===''||isNaN(+val) ? null : +val;
  cbData[idx].fetched = false;
  refreshRow(idx); updateSummary();
}}
function upCB(idx, val) {{
  cbData[idx].cbPrice = val===''||isNaN(+val) ? null : +val;
  refreshRow(idx); updateSummary();
}}
function refreshRow(idx) {{
  const d = cbData[idx];
  const t = T(d.stockPrice, d.conv);
  const cr = (d.cbPrice&&t)?d.cbPrice/t*100:null;
  const sr = (d.stockPrice&&d.conv)?d.stockPrice/d.conv*100:null;
  const row = document.querySelector(`tr[data-idx="${{idx}}"]`);
  if (!row) return;
  const tds = row.querySelectorAll('td');
  tds[6].innerHTML = t ? `<span style="color:#888">${{fmt(t)}}</span>` : '<span class="na">—</span>';
  tds[8].innerHTML = badge(cr);
  tds[9].innerHTML = badge(sr);
  tr.className = (userMeta[d.cbCode]?.checked) ? 'row-checked' : '';
  const spInput = tds[5].querySelector('input');
  if (spInput) spInput.className = 'cell-in' + (d.stockPrice?' has-val':'');
}}
function updateSummary() {{
  const wp = cbData.filter(d=>d.stockPrice).length;
  const wc = cbData.filter(d=>d.cbPrice).length;
  const deep = cbData.filter(d=>{{const t=T(d.stockPrice,d.conv);return d.cbPrice&&t&&d.cbPrice/t*100<80;}}).length;
  const low  = cbData.filter(d=>{{const t=T(d.stockPrice,d.conv);return d.cbPrice&&t&&(r=>r>=80&&r<100)(d.cbPrice/t*100);}}).length;
  const prem = cbData.filter(d=>{{const t=T(d.stockPrice,d.conv);return d.cbPrice&&t&&d.cbPrice/t*100>=100;}}).length;
  document.getElementById('s-total').textContent = cbData.length;
  document.getElementById('s-price').textContent = wp;
  document.getElementById('s-cb').textContent = wc;
  document.getElementById('s-deep').textContent = deep;
  document.getElementById('s-low').textContent = low;
  document.getElementById('s-prem').textContent = prem;
}}
function openModal() {{ document.getElementById('modal').classList.add('show'); }}
function closeModal() {{ document.getElementById('modal').classList.remove('show'); }}
function applyJSON() {{
  try {{
    const json = JSON.parse(document.getElementById('jsonInput').value.trim());
    const overwrite = document.querySelector('input[name=importMode]:checked').value === 'overwrite';
    let sCount=0, cbCount=0;
    const stockMap = json.stock_prices || json;
    const cbMap    = json.cb_prices || {{}};
    cbData.forEach(d => {{
      const sp = stockMap[d.stock];
      if (sp && sp>0 && (overwrite||!d.stockPrice)) {{ d.stockPrice=sp; d.fetched=true; sCount++; }}
      const cp = cbMap[d.cbCode];
      if (cp && cp>0 && (overwrite||!d.cbPrice)) {{ d.cbPrice=cp; cbCount++; }}
    }});
    document.getElementById('lastUpdate').textContent =
      `${{overwrite?'覆蓋':'補入'}} 股價${{sCount}}筆/CB現價${{cbCount}}筆（${{new Date().toLocaleTimeString('zh-TW')}}）`;
    closeModal(); render();
  }} catch(e) {{ alert('JSON 格式錯誤！'); }}
}}
function exportCSV() {{
  const rows = [['CB代號','股票','名稱','產業','轉換價','股票現價','CB理論價','CB現價','CB/理論%','股/轉換%']];
  cbData.forEach(d => {{
    const t = T(d.stockPrice, d.conv);
    rows.push([d.cbCode,d.stock,d.name,INDUSTRY[d.stock]||'',d.conv,
      d.stockPrice??'', t?t.toFixed(2):'', d.cbPrice??'',
      (d.cbPrice&&t)?(d.cbPrice/t*100).toFixed(2):'',
      (d.stockPrice&&d.conv)?(d.stockPrice/d.conv*100).toFixed(2):''
    ]);
  }});
  const a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob(['\uFEFF'+rows.map(r=>r.join(',')).join('\\n')],{{type:'text/csv;charset=utf-8'}}));
  a.download='cb_tracker.csv'; a.click();
}}
render();
</script>
</body>
</html>"""
    return html

def save_tracker_json(records, cb_prices, stock_prices):
    INDUSTRY = {
      "1101":"水泥","1256":"食品","1316":"化學","1338":"化學","1402":"紡織",
      "1436":"紡織","1438":"紡織","1466":"紡織","1472":"紡織","1474":"紡織",
      "1514":"電機","1526":"汽車","1536":"汽車","1560":"機械","1586":"塑膠",
      "1595":"其他","1598":"運動休閒","1609":"電線電纜","1786":"生技醫療",
      "1909":"造紙","2034":"鋼鐵","2066":"鋼鐵","2104":"橡膠","2201":"汽車",
      "2228":"汽車零件","2233":"汽車零件","2236":"汽車零件","2247":"汽車",
      "2329":"半導體","2337":"半導體","2338":"半導體","2351":"電子零組件",
      "2368":"半導體","2374":"電腦周邊","2402":"電子零組件","2427":"電子通路",
      "2436":"半導體","2439":"電子零組件","2442":"電子零組件","2455":"半導體",
      "2462":"電子零組件","2464":"電子零組件","2467":"半導體設備","2476":"電子零組件",
      "2486":"電子零組件","2528":"建設","2530":"建設","2548":"建設","2610":"航運",
      "2641":"航運","2727":"餐飲","2732":"餐飲","2743":"觀光","2745":"觀光",
      "2753":"餐飲","2755":"餐飲","2756":"餐飲","2906":"百貨","3006":"半導體",
      "3011":"其他電子","3013":"其他電子","3016":"半導體","3032":"電子通路",
      "3033":"電子通路","3037":"PCB","3040":"其他電子","3045":"電信",
      "3088":"其他電子","3092":"電子零組件","3095":"其他電子","3122":"半導體",
      "3128":"其他電子","3131":"半導體設備","3135":"半導體","3138":"電子零組件",
      "3141":"半導體","3167":"半導體設備","3188":"其他電子","3207":"其他電子",
      "3257":"其他電子","3260":"電腦周邊","3272":"其他電子","3284":"其他電子",
      "3303":"光電","3305":"電子通路","3312":"其他電子","3313":"其他電子",
      "3322":"其他電子","3323":"電子零組件","3324":"散熱","3346":"光電",
      "3357":"其他電子","3362":"光電","3376":"電子零組件","3388":"電子通路",
      "3390":"其他電子","3416":"其他電子","3479":"其他電子","3483":"其他電子",
      "3484":"其他電子","3516":"其他電子","3518":"其他電子","3522":"其他電子",
      "3526":"電子零組件","3543":"其他電子","3548":"其他電子","3551":"半導體設備",
      "3564":"其他電子","3583":"半導體設備","3587":"半導體設備","3591":"光電",
      "3605":"電子零組件","3617":"電源供應","3680":"半導體設備","3684":"生技",
      "3687":"電競","3715":"PCB","3722":"電信","3800":"PCB","3801":"工業電腦",
      "3815":"光電","3824":"半導體","3839":"顯示器","3860":"半導體","3862":"半導體",
      "3863":"其他電子","3865":"再生能源","3868":"伺服器","3886":"半導體",
      "3889":"通訊網路","3892":"其他","3901":"再生能源","3910":"機械",
      "3912":"工業電腦","3915":"其他電子","3920":"其他電子","3921":"散熱",
      "3923":"電子零組件","3932":"其他電子","3936":"其他電子","3937":"半導體",
      "4101":"生技","4103":"半導體","4222":"工業區","4231":"電子零組件",
      "4301":"其他電子","4306":"包材","4321":"機械","4345":"生技","4356":"其他",
      "4390":"生技醫療","4401":"航運","4406":"其他電子","4427":"其他電子",
      "4436":"其他電子","4462":"電子零組件","4467":"其他","4485":"其他",
      "4490":"其他電子","4492":"其他電子","4498":"觀光","4507":"機械",
      "4520":"生技","4522":"通訊","4531":"化學","4533":"再生能源","4534":"電機",
      "4536":"電子零組件","4539":"電機","4540":"電線電纜","4541":"半導體設備",
      "4542":"光電","4543":"電子零組件","4544":"其他","4552":"半導體",
      "4554":"電子零組件","4555":"半導體","4556":"其他電子","4565":"化學",
      "4566":"通訊網路","4567":"半導體","4569":"建設","4572":"其他電子",
      "4577":"半導體","4582":"PCB","4606":"通訊","4613":"生技","4623":"電子零組件",
      "4637":"其他電子","4638":"生技","4642":"航空","4666":"其他電子",
      "4667":"通訊網路","4689":"半導體","4694":"其他電子","4696":"建設",
      "4921":"通訊網路","4922":"半導體","4923":"半導體","4947":"其他電子",
      "5051":"光電","5213":"汽車零件","5422":"半導體","5458":"觀光",
      "5459":"半導體","5480":"通訊網路","5492":"其他電子","5508":"建設",
      "5510":"建設","5537":"建設","5610":"其他電子","5646":"電子零組件",
      "5701":"觀光","5715":"餐飲","5741":"建設","5751":"餐飲","5752":"其他電子",
      "5811":"金融","5851":"其他","6004":"光電","6017":"再生能源",
      "6053":"其他電子","6056":"半導體","6059":"機械","6065":"其他電子",
      "6073":"農業","6076":"機械","6079":"其他","6084":"電機","6086":"其他電子",
      "6101":"其他電子","6104":"半導體","6112":"半導體","6113":"其他電子",
      "6114":"半導體設備","6115":"電子零組件","6117":"其他電子","6126":"電子零組件",
      "6150":"顯示卡","6156":"其他電子","6163":"通訊網路","6177":"建設",
      "6179":"其他電子","6182":"半導體","6184":"有線電視","6187":"自動化",
      "6189":"電子通路","6190":"通訊網路","6197":"電子零組件","6199":"生技",
      "6207":"其他電子","6209":"光電","6219":"建設","6243":"半導體","6269":"PCB",
      "6274":"PCB","6275":"其他電子","6282":"電源供應","6284":"電子零組件",
      "6290":"電子零組件","6414":"工業電腦","6425":"其他","6432":"光電",
      "6442":"光電","6451":"半導體","6469":"藥局","6472":"生技醫療",
      "6477":"其他電子","6509":"化學","6510":"半導體設備","6515":"半導體設備",
      "6533":"半導體","6538":"機械","6546":"其他電子","6547":"生技",
      "6552":"其他電子","6576":"生技","6584":"其他","6585":"化學",
      "6591":"其他電子","6616":"其他","6624":"其他","6629":"其他",
      "6641":"其他","6645":"生技","6664":"其他電子","6670":"機械",
      "6672":"電子零組件","6680":"其他電子","6706":"其他電子","6715":"電子零組件",
      "6753":"造船","6768":"其他","6771":"環保","6796":"生技醫療",
      "6804":"其他電子","6806":"再生能源","6821":"其他電子","6823":"其他電子",
      "6830":"半導體設備","6835":"其他","6840":"其他電子","6843":"其他電子",
      "6846":"再生能源","6854":"光電","6862":"其他","6863":"其他",
      "6870":"其他電子","6873":"再生能源","6874":"其他電子","6890":"其他",
      "6894":"其他電子","6904":"其他電子","6913":"其他電子","6957":"其他",
      "6982":"機械","7402":"電子零組件","7556":"其他","7713":"生技",
      "8021":"電子零組件","8027":"半導體","8028":"半導體","8042":"電源供應",
      "8050":"工業電腦","8087":"再生能源","8091":"其他電子","8092":"其他電子",
      "8104":"光電","8112":"電子通路","8147":"其他電子","8155":"其他電子",
      "8171":"其他電子","8210":"伺服器","8255":"電子零組件","8299":"半導體",
      "8374":"電子通路","8404":"紡織","8411":"其他","8422":"環保",
      "8431":"其他電子","8433":"其他電子","8442":"其他","8462":"健身",
      "8466":"其他","8467":"運動休閒","8473":"環保","8478":"遊艇",
      "8488":"其他","8927":"食品","8996":"機械","9802":"紡織","9906":"其他",
      "9933":"工程","9935":"其他","9939":"包材","9941":"租賃","9958":"再生能源",
    }
    data = []
    for r in records:
        cb_code = r["cb_code"]
        stock   = r["stock"]
        conv    = r["conv"]
        cb_p    = cb_prices.get(cb_code)
        sp      = stock_prices.get(stock)
        th      = round(sp/conv*100,2) if sp and conv and conv>0 else None
        cbth    = round(cb_p/th*100,2) if cb_p and th and th>0 else None
        stk2    = round(sp/conv*100,2) if sp and conv and conv>0 else None
        data.append({"cb":cb_code,"stk":stock,"name":r["name"],
            "ind":INDUSTRY.get(stock,"—"),
            "conv":float(conv) if conv==conv else None,
            "cbp":cb_p,"sp":sp,"th":th,"cbth":cbth,"stk2":stk2})
    out = {"updated":datetime.now().strftime("%Y/%m/%d %H:%M"),
           "date":datetime.now().strftime("%Y/%m/%d"),"data":data}
    with open("cb_tracker_latest.json","w",encoding="utf-8") as f:
        json.dump(out,f,ensure_ascii=False,indent=2)
    print(f"✓ 儲存 cb_tracker_latest.json（{len(data)} 筆）")

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--show",       action="store_true")
    parser.add_argument("--cb-only",    action="store_true")
    parser.add_argument("--stock-only", action="store_true")
    args = parser.parse_args()

    print("="*50)
    print(f"  CB報價工具  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*50)

    records     = load_cb_excel()
    cb_codes    = [r["cb_code"] for r in records]
    stock_codes = sorted(set(r["stock"] for r in records))

    cb_prices    = {}
    stock_prices = {}

    if not args.stock_only:
        cb_prices = fetch_all(cb_codes, "抓CB現價", try_tse=False)
    if not args.cb_only:
        stock_prices = fetch_all(stock_codes, "抓股票現價", try_tse=True)

    # 儲存 prices.json
    result = {"stock_prices": stock_prices, "cb_prices": cb_prices}
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"\n✓ 儲存 {OUTPUT_FILE}")

    # 產生更新的 HTML
    update_time = datetime.now().strftime('%Y-%m-%d %H:%M')
    html = generate_html(records, cb_prices, stock_prices, update_time)
    with open(HTML_FILE, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✓ 更新 {HTML_FILE}（{len(records)} 筆CB，股價{len(stock_prices)}筆，CB現價{len(cb_prices)}筆）")

    if args.show:
        print(f"\n股票 {len(stock_prices)} 筆 / CB {len(cb_prices)} 筆")

    save_tracker_json(records, cb_prices, stock_prices)
    update_google_sheet(records, cb_prices)
    print("="*50)
    print("完成！直接打開 cb_tracker.html 即可，不需要貼 JSON。")
    print("="*50)


def update_google_sheet(records, cb_prices):
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        import os

        SHEET_ID  = "1BVNzkag83tgyzdv4FQ8qRRVBpmOrYM8k4v0va0P_-rk"
        CREDS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "credentials.json")
        SHEET_NAME = "CB持股"

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = Credentials.from_service_account_file(CREDS_FILE, scopes=scopes)
        gc = gspread.authorize(creds)
        ws = gc.open_by_key(SHEET_ID).worksheet(SHEET_NAME)

        all_values = ws.get_all_values()
        updates = []
        for i, row in enumerate(all_values[2:], start=3):
            if not row or not row[0]:
                continue
            cb_code = str(row[0]).strip()
            price = cb_prices.get(cb_code)
            if price is not None:
                updates.append({"range": f"I{i}", "values": [[price]]})

        if updates:
            ws.batch_update(updates)
            print(f"✓ 已更新 Google Sheets（{len(updates)} 筆現價）")
        else:
            print("⚠ Google Sheets：找不到對應的 CB 代號")

    except Exception as e:
        print(f"⚠ Google Sheets 更新失敗：{e}")




# ── 額外輸出 JSON 供網頁使用 ──────────────────────────

if __name__ == "__main__":
    main()


