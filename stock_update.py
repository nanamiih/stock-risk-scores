# ============================================
# stock_update.py
# 自動抓取財務比率 + Z/F Score 並輸出 Excel
# 適用於 GitHub Actions 無 Notebook 環境
# ============================================

import os
import sys
import re
import requests
import pandas as pd
from bs4 import BeautifulSoup
from openpyxl import Workbook

# ---- 自動安裝必要套件 (僅第一次需要) ----
def ensure_packages():
    try:
        import lxml, html5lib, openpyxl, bs4, requests, pandas
    except ImportError:
        os.system("pip install pandas requests lxml html5lib beautifulsoup4 openpyxl")

ensure_packages()

# ---- 公司清單 ----
TICKERS = {
    "AA": "Alcoa",
    "RIO": "Rio Tinto",
    "NHYDY": "Norsk Hydro",
    "RS": "Reliance Steel & Aluminum",
    "KALU": "Kaiser Aluminum",
    "RYI": "Ryerson Holding"
}

# ---- 指標 ----
TARGET = {
    "Current Ratio": "Current Ratio",
    "Debt": "Debt / Equity Ratio",
    "EBITDA": "EBITDA",
    "Free Cash Flow": "Free Cash Flow (Millions)",
    "Inventory Turnover": "Inventory Turnover",
    "Net Income": "Net Income (Millions)"
}

# ---- 抓取財務比率 ----
def fetch_ratios(symbol):
    url = f"https://stockanalysis.com/stocks/{symbol.lower()}/financials/ratios/"
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        html = requests.get(url, headers=headers).text
        tables = pd.read_html(html)
    except Exception as e:
        print(f"⚠️ {symbol} 抓取失敗: {e}")
        return None

    if not tables:
        print(f"⚠️ {symbol} 找不到表格")
        return None

    df = tables[0]
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [
            " ".join([str(c) for c in col if c and c != 'nan']).strip()
            for col in df.columns
        ]
    df.rename(columns={df.columns[0]: "Metric"}, inplace=True)
    df = df[df["Metric"].str.contains("|".join(TARGET.keys()), case=False, na=False)]
    df["Metric"] = df["Metric"].apply(
        lambda x: next((v for k, v in TARGET.items() if k.lower() in x.lower()), x)
    )
    df = df.replace(["Upgrade", "-", "—"], pd.NA)
    df = df.set_index("Metric").T.reset_index().rename(columns={"index": "Date_1"})
    df["Date_1"] = df["Date_1"].apply(lambda x: re.sub(r"[\(\)'\"]+", "", str(x)).strip())
    df["Ticker"] = symbol
    return df.fillna("")

# ---- 抓取 Z / F Score ----
def fetch_scores(symbol):
    url = f"https://stockanalysis.com/stocks/{symbol.lower()}/statistics/"
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        html = requests.get(url, headers=headers).text
        df = pd.concat(pd.read_html(html), ignore_index=True)
    except Exception:
        return {"Altman Z-Score": "", "Piotroski F-Score": ""}

    df.columns = ["Metric", "Value"]
    z = df[df["Metric"].str.contains("Altman Z", na=False)]["Value"].values
    f = df[df["Metric"].str.contains("Piotroski F", na=False)]["Value"].values
    return {
        "Altman Z-Score": z[0] if len(z) else "",
        "Piotroski F-Score": f[0] if len(f) else ""
    }

# ---- 寫入 Excel ----
wb = Workbook()
wb.remove(wb.active)

for t, name in TICKERS.items():
    print(f"🔍 抓取 {name} ({t}) ...")
    ratios = fetch_ratios(t)
    scores = fetch_scores(t)

    if ratios is None:
        print(f"⚠️ {name} 無資料，略過")
        continue

    sheet = wb.create_sheet(title=name[:30])
    sheet.append(["Altman Z-Score", scores["Altman Z-Score"]])
    sheet.append(["Piotroski F-Score", scores["Piotroski F-Score"]])
    sheet.append([])

    # 寫入表格內容
    clean_df = pd.DataFrame(ratios).fillna("")
    sheet.append(clean_df.columns.tolist())
    for row in clean_df.itertuples(index=False):
        sheet.append(row)

    print(f"✅ {name} 完成")

wb.save("Stock_Risk_Scores.xlsx")
print("✅ 已輸出 Stock_Risk_Scores.xlsx")
