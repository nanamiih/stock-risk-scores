import pandas as pd
import requests, re, time
from openpyxl import Workbook

# -------------------------------------------------------
# 公司代碼與名稱
# -------------------------------------------------------
TICKERS = {
    "AA": "Alcoa",
    "RIO": "Rio Tinto",
    "NHYDY": "Norsk Hydro",  # ✅ 使用 ADR 代碼（美股）
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
# 抓取財報比率
# -------------------------------------------------------
def fetch_ratios(symbol):
    url = f"https://stockanalysis.com/stocks/{symbol.lower()}/financials/ratios/"
    headers = {"User-Agent": "Mozilla/5.0"}

    # retry 機制（最多三次）
    for attempt in range(3):
        try:
            r = requests.get(url, headers=headers, timeout=20)
            if r.status_code == 200:
                break
        except Exception:
            time.sleep(3)
    else:
        print(f"⚠️ {symbol}: 無法連線")
        return None

    try:
        tables = pd.read_html(r.text)
    except Exception:
        print(f"⚠️ {symbol}: 找不到表格")
        return None

    if not tables:
        print(f"⚠️ {symbol}: 找不到表格")
        return None

    df = tables[0].copy()

    # 壓平多層標題
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [" ".join([str(c) for c in col if c and c != "nan"]).strip()
                      for col in df.columns]

    df.rename(columns={df.columns[0]: "Metric"}, inplace=True)

    # 篩選欄位
    df = df[df["Metric"].str.contains("|".join(TARGET.keys()), case=False, na=False)]
    df["Metric"] = df["Metric"].apply(
        lambda x: next((v for k, v in TARGET.items() if k.lower() in x.lower()), x)
    )

    # 轉置
    df = df.set_index("Metric").T.reset_index().rename(columns={"index": "Date_1"})

    # ✅ 日期清理：只保留 YYYY/MM/DD 格式
    def clean_date(x):
        x = str(x)
        # 嘗試找 YYYY-MM-DD 或 YYYY年MM月DD日 或 Dec 31 2024
        m = re.search(r"(\d{4}[-/]\d{1,2}[-/]\d{1,2})", x)
        if m:
            return m.group(1).replace("-", "/")
        m = re.search(r"([A-Za-z]{3,9}\s\d{1,2}\s\d{4})", x)
        if m:
            try:
                return pd.to_datetime(m.group(1)).strftime("%Y/%m/%d")
            except:
                pass
        m = re.search(r"(\d{4})", x)
        return f"{m.group(1)}/12/31" if m else ""

    df["Date_1"] = df["Date_1"].apply(clean_date)
    df = df.loc[:, ~df.columns.duplicated()].fillna("")

    return df


# -------------------------------------------------------
# 抓取 Z/F Score
# -------------------------------------------------------
def fetch_scores(symbol):
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

    # 固定欄位順序
    final_cols = [
        "Date_1", "EBITDA", "Debt / Equity Ratio",
        "Inventory Turnover", "Current Ratio",
        "Ticker", "Altman Z-Score", "Piotroski F-Score"
    ]
    ratios = ratios[[c for c in final_cols if c in ratios.columns]]

    # 寫入工作表
    sheet = wb.create_sheet(title=name[:30])
    sheet.append(ratios.columns.tolist())
    for row in ratios.itertuples(index=False):
        sheet.append(["" if pd.isna(x) else x for x in row])

    print(f"✅ {name} 完成")

wb.save("Stock_Risk_Scores.xlsx")
print("✅ 已輸出 Stock_Risk_Scores.xlsx ✅")
