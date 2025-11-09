import pandas as pd
import requests, re, time, os
from bs4 import BeautifulSoup
from datetime import datetime

# -------------------------------------------------------
# 公司代碼與分類
# -------------------------------------------------------
TICKERS = {
    #mills
    "AA": {"name": "Alcoa", "url": "https://stockanalysis.com/stocks/aa/financials/ratios/", "category": "mills"},
    "RIO": {"name": "Rio Tinto", "url": "https://stockanalysis.com/stocks/rio/financials/ratios/", "category": "mills"},
    "NHY": {"name": "Norsk Hydro", "url": "https://stockanalysis.com/quote/osl/NHY/financials/ratios/", "category": "mills"},

    #distributors
    "RS": {"name": "Reliance", "url": "https://stockanalysis.com/stocks/rs/financials/ratios/", "category": "distributor"},
    "KALU": {"name": "Kaiser", "url": "https://stockanalysis.com/stocks/kalu/financials/ratios/", "category": "distributor"},
    "RYI": {"name": "Ryerson", "url": "https://stockanalysis.com/stocks/ryi/financials/ratios/", "category": "distributor"},
    "BVB:ALR": {"name": "Alro Steel", "url": "https://stockanalysis.com/quote/bvb/alr/financials/", "category": "distributor"},

    #supplier
    "SEOJIN": {"name": "Seojin", "url": "https://stockanalysis.com/stocks/seojin/financials/ratios/", "category": "supplier"},
    "ULTR": {"name": "Ultra", "url": "https://stockanalysis.com/stocks/uctt/financials/ratios/", "category": "supplier"},
    "FOX": {"name": "Foxconn", "url": "https://stockanalysis.com/stocks/hnhaf/financials/ratios/", "category": "supplier"},
    "FERRO": {"name": "Ferrotec", "url": "https://stockanalysis.com/stocks/frtcf/financials/ratios/", "category": "supplier"},
    "BHE": {"name": "Benchmark", "url": "https://stockanalysis.com/stocks/bhe/financials/ratios/", "category": "supplier"},
    "CLS": {"name": "Celestica", "url": "https://stockanalysis.com/stocks/cls/financials/ratios/", "category": "supplier"},
    "JABIL": {"name": "Jabil", "url": "https://stockanalysis.com/stocks/jbl/financials/ratios/", "category": "supplier"},
    "FLEX": {"name": "Flex", "url": "https://stockanalysis.com/stocks/flex/financials/ratios/", "category": "supplier"},
    "MKS": {"name": "MKS", "url": "https://stockanalysis.com/stocks/mksi/financials/ratios/", "category": "supplier"},
    "ATLAS": {"name": "Atlas Tech", "url": "https://stockanalysis.com/stocks/atlas/financials/ratios/", "category": "supplier"},
}

TARGET = {
    "EBITDA": "EBITDA",
    "Debt": "Debt / Equity Ratio",
    "Inventory Turnover": "Inventory Turnover",
    "Current Ratio": "Current Ratio"
}


# -------------------------------------------------------
# 讀取財報比率
# -------------------------------------------------------
def fetch_ratios(symbol, url):
    headers = {"User-Agent": "Mozilla/5.0"}
    html = None

    for attempt in range(5):
        try:
            r = requests.get(url, headers=headers, timeout=25)
            if r.status_code == 200 and "<table" in r.text:
                html = r.text
                break
        except Exception:
            pass
        time.sleep(3)

    # 若 ratios 抓不到，自動換成 /financials/
    if not html and "/ratios/" in url:
        alt_url = url.replace("/ratios/", "/")
        try:
            r = requests.get(alt_url, headers=headers, timeout=25)
            if r.status_code == 200 and "<table" in r.text:
                html = r.text
        except Exception:
            pass

    if not html:
        return None

    try:
        tables = pd.read_html(html)
    except Exception:
        tables = []

    if not tables:
        soup = BeautifulSoup(html, "html.parser")
        raw_table = soup.find("table")
        tables = [pd.read_html(str(raw_table))[0]] if raw_table else []

    if not tables:
        return None

    df = tables[0].copy()
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [" ".join([str(c) for c in col if c and c != "nan"]).strip() for col in df.columns]

    df.rename(columns={df.columns[0]: "Metric"}, inplace=True)
    df = df[df["Metric"].str.contains("|".join(TARGET.keys()), case=False, na=False)]
    df["Metric"] = df["Metric"].apply(lambda x: next((v for k, v in TARGET.items() if k.lower() in x.lower()), x))
    df = df.set_index("Metric").T.reset_index().rename(columns={"index": "Date_1"})

    def clean_date(x):
        x = str(x)
        m = re.search(r"([A-Za-z]{3,9}\s\d{1,2}\s\d{4})", x)
        if m:
            try:
                return pd.to_datetime(m.group(1)).strftime("%Y/%m/%d")
            except:
                pass
        today_str = datetime.today().strftime("%Y/%m/%d")
        if any(k in x for k in ["Current", "TTM", "Oct", "Sep"]):
            return today_str
        m = re.search(r"(\d{4})", x)
        if m:
            return f"{m.group(1)}/12/31"
        return today_str

    df["Date_1"] = df["Date_1"].apply(clean_date)
    df = df.loc[:, ~df.columns.duplicated()].fillna("")
    return df


# -------------------------------------------------------
# 抓取 Z/F Score
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
# 主程式：整合成單一表 + 輸出
# -------------------------------------------------------
all_data = []

for t, info in TICKERS.items():
    print(f"🔍 抓取 {info['name']} ({t}) ...")
    ratios = fetch_ratios(t, info["url"])
    scores = fetch_scores(t)

    if ratios is None or ratios.empty:
        print(f"⚠️ {info['name']} ({t}) 沒抓到資料")
        ratios = pd.DataFrame(columns=["Date_1", "EBITDA", "Debt / Equity Ratio", "Inventory Turnover", "Current Ratio"])

    ratios["Ticker"] = t
    ratios["Altman Z-Score"] = scores.get("Altman Z-Score", "")
    ratios["Piotroski F-Score"] = scores.get("Piotroski F-Score", "")
    ratios["Category"] = info["category"]
    all_data.append(ratios)

final_df = pd.concat(all_data, ignore_index=True)

print("\n📊 抓取完成，以下是各公司資料筆數：")
for t in final_df["Ticker"].unique():
    count = len(final_df[final_df["Ticker"] == t])
    print(f" - {t}: {count} rows")

# 🔹 移除任何含有 "Upgrade" 的列
final_df = final_df[~final_df.apply(lambda row: row.astype(str).str.contains("Upgrade", case=False).any(), axis=1)]

# 🔹 固定欄位順序
final_cols = ["Date_1", "EBITDA", "Debt / Equity Ratio", "Inventory Turnover",
              "Current Ratio", "Ticker", "Altman Z-Score", "Piotroski F-Score", "Category"]
final_df = final_df[[c for c in final_cols if c in final_df.columns]]

# -------------------------------------------------------
# 輸出 Excel
# -------------------------------------------------------
output_file = "Stock_Risk_Scores.xlsx"
final_df.to_excel(output_file, index=False)
print(f"\n✅ 已輸出乾淨版 Stock_Risk_Scores.xlsx（無 Upgrade 列）")
print("📁 輸出位置：", os.path.abspath(output_file))
