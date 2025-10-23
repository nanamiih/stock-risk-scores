import pandas as pd
import requests, re
from openpyxl import Workbook

# -------------------------------------------------------
# 公司代碼與名稱（含 Norsk Hydro）
# -------------------------------------------------------
TICKERS = {
    "AA": "Alcoa",
    "RIO": "Rio Tinto",
    "OSL:NHY": "Norsk Hydro",  # ✅ 正確代碼
    "RS": "Reliance Steel & Aluminum",
    "KALU": "Kaiser Aluminum",
    "RYI": "Ryerson Holding"
}

# 想抓的比率
TARGET = {
    "EBITDA": "EBITDA",
    "Debt": "Debt / Equity Ratio",
    "Inventory Turnover": "Inventory Turnover",
    "Current Ratio": "Current Ratio"
}

# -------------------------------------------------------
# 將 ticker 轉成 stockanalysis 網址格式
# -------------------------------------------------------
def format_symbol(symbol):
    # "OSL:NHY" → "osl-nhy"
    return symbol.lower().replace(":", "-")

# -------------------------------------------------------
# 抓財報比率（annual）
# -------------------------------------------------------
def fetch_ratios(symbol):
    url = f"https://stockanalysis.com/stocks/{format_symbol(symbol)}/financials/ratios/"
    headers = {"User-Agent": "Mozilla/5.0"}
    r = requests.get(url, headers=headers)
    if r.status_code != 200:
        print(f"⚠️ {symbol}: 無法連線 ({r.status_code})")
        return None

    try:
        tables = pd.read_html(r.text)
    except Exception as e:
        print(f"⚠️ {symbol}: 無法解析表格 ({e})")
        return None

    if not tables:
        print(f"⚠️ {symbol}: 找不到表格")
        return None

    df = tables[0]
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [
            " ".join([str(c) for c in col if c and c != "nan"]).strip()
            for col in df.columns
        ]

    df.rename(columns={df.columns[0]: "Metric"}, inplace=True)
    df = df[df["Metric"].str.contains("|".join(TARGET.keys()), case=False, na=False)]
    df["Metric"] = df["Metric"].apply(
        lambda x: next((v for k, v in TARGET.items() if k.lower() in x.lower()), x)
    )

    # 清理並轉置
    df = df.replace(["Upgrade", "-", "—"], pd.NA)
    df = df.set_index("Metric").T.reset_index().rename(columns={"index": "Date_1"})
    df["Date_1"] = df["Date_1"].apply(lambda x: re.sub(r"[^\w\s\-\/]", "", str(x)).strip())

    df = df[["Date_1"] + list(TARGET.values())]
    return df

# -------------------------------------------------------
# 抓 Z / F Score
# -------------------------------------------------------
def fetch_scores(symbol):
    url = f"https://stockanalysis.com/stocks/{format_symbol(symbol)}/statistics/"
    headers = {"User-Agent": "Mozilla/5.0"}
    r = requests.get(url, headers=headers)
    if r.status_code != 200:
        return {"Altman Z-Score": "", "Piotroski F-Score": ""}

    try:
        tables = pd.read_html(r.text)
        df = pd.concat(tables, ignore_index=True)
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
# 寫入 Excel（每家公司一個工作表）
# -------------------------------------------------------
wb = Workbook()
wb.remove(wb.active)

for t, name in TICKERS.items():
    print(f"🔍 抓取 {name} ({t}) ...")
    ratios = fetch_ratios(t)
    scores = fetch_scores(t)

    if ratios is None:
        print(f"⚠️ {name} 無資料，略過")
        continue

    # 加上 Z/F Score 與 Ticker
    ratios["Ticker"] = t
    ratios["Altman Z-Score"] = scores.get("Altman Z-Score", "")
    ratios["Piotroski F-Score"] = scores.get("Piotroski F-Score", "")

    # 固定欄位順序
    final_cols = [
        "Date_1", "EBITDA", "Debt / Equity Ratio", "Inventory Turnover",
        "Current Ratio", "Ticker", "Altman Z-Score", "Piotroski F-Score"
    ]
    ratios = ratios[[c for c in final_cols if c in ratios.columns]]

    # 寫入工作表
    sheet = wb.create_sheet(title=name[:30])
    sheet.append(ratios.columns.tolist())

    for row in ratios.itertuples(index=False):
        # ⚙️ 關鍵修正：清除 NA 避免 openpyxl crash
        clean_row = [("" if pd.isna(x) else x) for x in row]
        sheet.append(clean_row)

    print(f"✅ {name} 完成")

wb.save("Stock_Risk_Scores.xlsx")
print("✅ 已輸出 Stock_Risk_Scores.xlsx ✅")
