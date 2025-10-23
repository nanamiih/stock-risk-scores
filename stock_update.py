import pandas as pd
import requests, re, time
from openpyxl import Workbook

# -------------------------------------------------------
# 公司代碼與名稱
# -------------------------------------------------------
TICKERS = {
    "AA": "Alcoa",
    "RIO": "Rio Tinto",
    "NHYDY": "Norsk Hydro",  # ✅ 使用美股 ADR 代碼
    "RS": "Reliance Steel & Aluminum",
    "KALU": "Kaiser Aluminum",
    "RYI": "Ryerson Holding"
}

TARGET = {
    "EBITDA": "EBITDA",
    "Debt": "Debt / Equity Ratio",
    "Inventory Turnover": "Inventory Turnover",
    "Current Ratio": "Current Ratio"
}

# -------------------------------------------------------
# 統一轉換代碼格式
# -------------------------------------------------------
def format_symbol(symbol):
    return symbol.lower().replace(":", "-")

# -------------------------------------------------------
# 抓取財報比率（含防斷線重試 + 日期清理）
# -------------------------------------------------------
def fetch_ratios(symbol):
    url = f"https://stockanalysis.com/stocks/{format_symbol(symbol)}/financials/ratios/"
    headers = {"User-Agent": "Mozilla/5.0"}

    for attempt in range(3):  # 最多重試三次
        try:
            r = requests.get(url, headers=headers, timeout=20)
            if r.status_code == 200:
                break
        except Exception:
            time.sleep(3)
    else:
        print(f"⚠️ {symbol}: 無法連線")
        return None

    tables = pd.read_html(r.text)
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

    # 第一欄改名
    df.rename(columns={df.columns[0]: "Metric"}, inplace=True)

    # 篩選指標
    df = df[df["Metric"].str.contains("|".join(TARGET.keys()), case=False, na=False)]
    df["Metric"] = df["Metric"].apply(
        lambda x: next((v for k, v in TARGET.items() if k.lower() in x.lower()), x)
    )

    # 清理空白與符號
    df = df.replace(["Upgrade", "-", "—"], pd.NA)

    # 轉置
    df = df.set_index("Metric").T.reset_index().rename(columns={"index": "Date_1"})

    # 日期格式清理：只保留 YYYY/MM/DD
    df["Date_1"] = df["Date_1"].apply(lambda x: re.findall(r"\d{4}.*\d{2,}", str(x)))
    df["Date_1"] = df["Date_1"].apply(
        lambda x: x[0] if x else ""
    )
    df["Date_1"] = df["Date_1"].apply(lambda x: re.sub(r"[^0-9/]", "", x).strip())

    # 移除重複欄位，只保留第一個
    df = df.loc[:, ~df.columns.duplicated()]

    df = df.fillna("")
    return df

# -------------------------------------------------------
# 抓取 Z/F Score
# -------------------------------------------------------
def fetch_scores(symbol):
    url = f"https://stockanalysis.com/stocks/{format_symbol(symbol)}/statistics/"
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

for t, name in TICKERS.items():
    print(f"🔍 抓取 {name} ({t}) ...")
    ratios = fetch_ratios(t)
    scores = fetch_scores(t)

    if ratios is None:
        print(f"⚠️ {name} 無資料，略過")
        continue

    ratios["Ticker"] = t
    ratios["Altman Z-Score"] = scores.get("Altman Z-Score", "")
    ratios["Piotroski F-Score"] = scores.get("Piotroski F-Score", "")

    final_cols = [
        "Date_1", "EBITDA", "Debt / Equity Ratio",
        "Inventory Turnover", "Current Ratio",
        "Ticker", "Altman Z-Score", "Piotroski F-Score"
    ]
    ratios = ratios[[c for c in final_cols if c in ratios.columns]]

    sheet = wb.create_sheet(title=name[:30])
    sheet.append(ratios.columns.tolist())
    for row in ratios.itertuples(index=False):
        sheet.append(["" if pd.isna(x) else x for x in row])

    print(f"✅ {name} 完成")

wb.save("Stock_Risk_Scores.xlsx")
print("✅ 已輸出 Stock_Risk_Scores.xlsx ✅")
