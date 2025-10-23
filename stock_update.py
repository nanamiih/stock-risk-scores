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

    html = None
    for attempt in range(3):
        try:
            r = requests.get(url, headers=headers, timeout=25)
            if r.status_code == 200 and "<table" in r.text:
                html = r.text
                break
            else:
                print(f"⚠️ {symbol}: 第 {attempt+1} 次嘗試失敗，等待重試...")
                time.sleep(5)
        except Exception as e:
            print(f"⚠️ {symbol}: 嘗試失敗 {e}")
            time.sleep(5)

    if not html:
        print(f"❌ {symbol}: 連線三次仍失敗，略過")
        return None

    # 先嘗試 pandas 解析
    try:
        tables = pd.read_html(html)
    except Exception:
        tables = []

    # 若 pandas 沒抓到，用 BeautifulSoup 補抓
    if not tables:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html, "html.parser")
        raw_table = soup.find("table")
        if raw_table:
            tables = [pd.read_html(str(raw_table))[0]]

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

    # 篩選欄位
    df = df[df["Metric"].str.contains("|".join(TARGET.keys()), case=False, na=False)]
    df["Metric"] = df["Metric"].apply(
        lambda x: next((v for k, v in TARGET.items() if k.lower() in x.lower()), x)
    )

    # 轉置
    df = df.set_index("Metric").T.reset_index().rename(columns={"index": "Date_1"})

    # 日期清理
    def clean_date(x):
        x = str(x)
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
