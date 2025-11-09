import pandas as pd
import requests, re, time
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

    # 若抓不到 ratios，自動改抓 financials
    for attempt in range(5):
        try:
            r = requests.get(url, headers=headers, timeout=25)
            if r.status_code == 200 and "<table" in r.text:
                html = r.text
                break
        except Exception:
            pass
        time.sleep(3)

    if not html and "ratios" in url:
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
