import pandas as pd
import requests, re, time
from openpyxl import Workbook
from bs4 import BeautifulSoup

# -------------------------------------------------------
# 公司代碼與對應名稱
# -------------------------------------------------------
TICKERS = {
    "AA": {"name": "Alcoa", "url": "https://stockanalysis.com/stocks/aa/financials/ratios/"},
    "RIO": {"name": "Rio Tinto", "url": "https://stockanalysis.com/stocks/rio/financials/ratios/"},
    "NHY": {"name": "Norsk Hydro", "url": "https://stockanalysis.com/quote/osl/NHY/financials/ratios/"},  # ✅ 歐洲市場
    "RS": {"name": "Reliance Steel & Aluminum", "url": "https://stockanalysis.com/stocks/rs/financials/ratios/"},
    "KALU": {"name": "Kaiser Aluminum", "url": "https://stockanalysis.com/stocks/kalu/financials/ratios/"},
    "RYI": {"name": "Ryerson Holding", "url": "https://stockanalysis.com/stocks/ryi/financials/ratios/"},
    # -------- Suppliers --------
    "ULTR": {"name": "Ultra Clean Holdings", "url": "https://stockanalysis.com/stocks/uctt/financials/ratios/", "category": "Supplier"},
    "FOX": {"name": "Foxconn", "url": "https://stockanalysis.com/stocks/hnhaf/financials/ratios/", "category": "Supplier"},
    "FERRO": {"name": "Ferrotec Holdings", "url": "https://stockanalysis.com/stocks/frtcf/financials/ratios/", "category": "Supplier"},
    "BHE": {"name": "Benchmark Electronics", "url": "https://stockanalysis.com/stocks/bhe/financials/ratios/", "category": "Supplier"},
    "CLS": {"name": "Celestica", "url": "https://stockanalysis.com/stocks/clst/financials/ratios/", "category": "Supplier"},
    "FLEX": {"name": "Flex Ltd", "url": "https://stockanalysis.com/stocks/flex/financials/ratios/", "category": "Supplier"},
    "MKS": {"name": "MKS Instruments", "url": "https://stockanalysis.com/stocks/mksi/financials/ratios/", "category": "Supplier"}
}

TARGET = {
    "EBITDA": "EBITDA",
    "Debt": "Debt / Equity Ratio",
    "Inventory Turnover": "Inventory Turnover",
    "Current Ratio": "Current Ratio"
}


# -------------------------------------------------------
# 財報比率爬取（支援 quote/osl）
# -------------------------------------------------------
def fetch_ratios(symbol, url):
    headers = {"User-Agent": "Mozilla/5.0"}
    html = None

    # Retry up to 5 times
    for attempt in range(5):
        try:
            r = requests.get(url, headers=headers, timeout=25)
            if r.status_code == 200 and "<table" in r.text:
                html = r.text
                break
            print(f"⚠️ {symbol}: 第 {attempt+1} 次嘗試失敗 {url}")
        except Exception as e:
            print(f"⚠️ {symbol}: 嘗試失敗 ({e})")
        time.sleep(5)

    if not html:
        print(f"❌ {symbol}: 所有頁面都無法取得表格")
        return None

    # 嘗試用 pandas 讀取
    try:
        tables = pd.read_html(html)
    except Exception:
        tables = []

    # 如果 pandas 抓不到，用 BeautifulSoup
    if not tables:
        soup = BeautifulSoup(html, "html.parser")
        raw_table = soup.find("table")
        tables = [pd.read_html(str(raw_table))[0]] if raw_table else []

    if not tables:
        print(f"⚠️ {symbol}: 找不到表格")
        return None

    df = tables[0].copy()

    # 壓平多層標題
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [
            " ".join([str(c) for c in col if c and c != "nan"]).strip()
            for col in df.columns
        ]

    df.rename(columns={df.columns[0]: "Metric"}, inplace=True)

    # 篩選目標欄位
    df = df[df["Metric"].str.contains("|".join(TARGET.keys()), case=False, na=False)]
    df["Metric"] = df["Metric"].apply(
        lambda x: next((v for k, v in TARGET.items() if k.lower() in x.lower()), x)
    )

    # 轉置表格
    df = df.set_index("Metric").T.reset_index().rename(columns={"index": "Date_1"})

   # 日期清理（保持 YYYY/MM/DD）
    from datetime import datetime
    
    def clean_date(x):
        x = str(x)
        
        # 嘗試解析完整英文日期（例: Oct 25 2025）
        m = re.search(r"([A-Za-z]{3,9}\s\d{1,2}\s\d{4})", x)
        if m:
            try:
                return pd.to_datetime(m.group(1)).strftime("%Y/%m/%d")
            except:
                pass
    
        today_str = datetime.today().strftime("%Y/%m/%d")
        m = re.search(r"(\d{4})", x)
        
        # 若是"Current"、"TTM"、"Oct"、"Sep"等 → 使用今天日期
        if any(k in x for k in ["Current", "TTM", "Oct", "Sep"]):
            return today_str
        elif m:
            return f"{m.group(1)}/12/31"
        else:
            return today_str

    df["Date_1"] = df["Date_1"].apply(clean_date)
    df = df.loc[:, ~df.columns.duplicated()].fillna("")
    return df


# -------------------------------------------------------
# 抓取 Z/F Score（支援 quote/osl/NHY）
# -------------------------------------------------------
def fetch_scores(symbol):
    if symbol == "NHY":
        url = "https://stockanalysis.com/quote/osl/NHY/statistics/"
    else:
        url = f"https://stockanalysis.com/stocks/{symbol.lower()}/statistics/"

    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        r = requests.get(url, headers=headers, timeout=20)
        if r.status_code != 200:
            return {"Altman Z-Score": "", "Piotroski F-Score": ""}
        df = pd.concat(pd.read_html(r.text), ignore_index=True)
        df.columns = ["Metric", "Value"]
        z = df[df["Metric"].str.contains("Altman Z", na=False)]["Value"].values
        f = df[df["Metric"].str.contains("Piotroski F", na=False)]["Value"].values
        return {
            "Altman Z-Score": z[0] if len(z) else "",
            "Piotroski F-Score": f[0] if len(f) else ""
        }
    except Exception:
        return {"Altman Z-Score": "", "Piotroski F-Score": ""}


# -------------------------------------------------------
# 寫入 Excel
# -------------------------------------------------------
wb = Workbook()
wb.remove(wb.active)

for t, info in TICKERS.items():
    print(f"🔍 抓取 {info['name']} ({t}) ...")
    ratios = fetch_ratios(t, info["url"])
    scores = fetch_scores(t)

    if ratios is None or ratios.empty:
        ratios = pd.DataFrame(columns=["Date_1", "EBITDA", "Debt / Equity Ratio",
                                       "Inventory Turnover", "Current Ratio"])
        print(f"⚠️ {info['name']}: 無資料，建立空白頁。")

    ratios["Ticker"] = t
    ratios["Altman Z-Score"] = scores.get("Altman Z-Score", "")
    ratios["Piotroski F-Score"] = scores.get("Piotroski F-Score", "")

    # 固定欄位順序
    final_cols = [
        "Date_1", "EBITDA", "Debt / Equity Ratio",
        "Inventory Turnover", "Current Ratio",
        "Ticker", "Altman Z-Score", "Piotroski F-Score"
    ]
    ratios = ratios[[c for c in final_cols if c in ratios.columns]]

    sheet = wb.create_sheet(title=info["name"][:30])
    sheet.append(ratios.columns.tolist())
    for row in ratios.itertuples(index=False):
        sheet.append(["" if pd.isna(x) else x for x in row])

    print(f"✅ {info['name']} 完成")

wb.save("Stock_Risk_Scores.xlsx")
print("✅ 已輸出 Stock_Risk_Scores.xlsx ✅")
