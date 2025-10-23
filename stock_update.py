#!/usr/bin/env python
# coding: utf-8

# In[1]:


get_ipython().run_line_magic('pip', 'install pandas requests lxml openpyxl')


# In[11]:


import os
os.system("pip install html5lib lxml beautifulsoup4")



# In[14]:


import pandas as pd
import requests, re
from openpyxl import Workbook
from bs4 import BeautifulSoup

# -------------------------------------------------------
# 追蹤的公司（用 StockAnalysis 上的股票代碼）
# -------------------------------------------------------
TICKERS = {
    "AA": "Alcoa",
    "RIO": "Rio Tinto",
    "NHYDY": "Norsk Hydro",
    "RS": "Reliance Steel & Aluminum",
    "KALU": "Kaiser Aluminum",
    "RYI": "Ryerson Holding"
}

# 想抓的比率
TARGET = {
    "Current Ratio": "Current Ratio",
    "Debt": "Debt / Equity Ratio",
    "EBITDA": "EBITDA",
    "Free Cash Flow": "Free Cash Flow (Millions)",
    "Inventory Turnover": "Inventory Turnover",
    "Net Income": "Net Income (Millions)"
}

# -------------------------------------------------------
# 抓取財務比率頁面（非 API）
# -------------------------------------------------------
def fetch_ratios(symbol):
    url = f"https://stockanalysis.com/stocks/{symbol.lower()}/financials/ratios/"
    headers = {"User-Agent": "Mozilla/5.0"}
    r = requests.get(url, headers=headers)
    if r.status_code != 200:
        print(f"⚠️ {symbol}: 無法連線 ({r.status_code})")
        return None

    soup = BeautifulSoup(r.text, "html.parser")
    tables = pd.read_html(str(soup))
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

    # 清理與轉置
    df = df.replace(["Upgrade", "-", "—"], pd.NA)
    df = df.set_index("Metric").T.reset_index().rename(columns={"index": "Date_1"})
    df["Date_1"] = df["Date_1"].apply(lambda x: re.sub(r"[\(\)'\"]+", "", str(x)).strip())
    df["Ticker"] = symbol
    return df.fillna("")

# -------------------------------------------------------
# 抓取 Z-Score / F-Score
# -------------------------------------------------------
def fetch_scores(symbol):
    url = f"https://stockanalysis.com/stocks/{symbol.lower()}/statistics/"
    headers = {"User-Agent": "Mozilla/5.0"}
    r = requests.get(url, headers=headers)
    if r.status_code != 200:
        print(f"⚠️ {symbol}: 無法抓取統計資料 ({r.status_code})")
        return {"Altman Z-Score": "", "Piotroski F-Score": ""}
    try:
        tables = pd.read_html(r.text)
    except Exception:
        return {"Altman Z-Score": "", "Piotroski F-Score": ""}
    if not tables:
        return {"Altman Z-Score": "", "Piotroski F-Score": ""}
    df = pd.concat(tables, ignore_index=True)
    df.columns = ["Metric", "Value"]
    z = df[df["Metric"].str.contains("Altman Z", na=False)]["Value"].values
    f = df[df["Metric"].str.contains("Piotroski F", na=False)]["Value"].values
    return {
        "Altman Z-Score": z[0] if len(z) else "",
        "Piotroski F-Score": f[0] if len(f) else ""
    }

# -------------------------------------------------------
# 寫入 Excel（每家公司一個工作表）
# -------------------------------------------------------
wb = Workbook()
wb.remove(wb.active)

for t, name in TICKERS.items():
    print(f"🔍 抓取 {name} ({t}) ...")
    try:
        ratios = fetch_ratios(t)
        scores = fetch_scores(t)
    except Exception as e:
        print(f"⚠️ {name} 抓取失敗: {e}")
        continue

    if ratios is None:
        print(f"⚠️ {name} 無資料，略過")
        continue

    # 寫入資料
    sheet = wb.create_sheet(title=name[:30])
    sheet.append(["Altman Z-Score", scores["Altman Z-Score"]])
    sheet.append(["Piotroski F-Score", scores["Piotroski F-Score"]])
    sheet.append([])

    clean_df = pd.DataFrame(ratios).fillna("")
    sheet.append(clean_df.columns.tolist())
    for row in clean_df.itertuples(index=False):
        sheet.append(row)

    print(f"✅ {name} 完成")

wb.save("Stock_Risk_Scores.xlsx")
print("✅ 已輸出 Stock_Risk_Scores.xlsx")


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[8]:


import pandas as pd
import requests, re
from datetime import date

# 想追蹤的公司
TICKERS = ["AA", "RIO", "NUE"]

# 想抓的比率
TARGET = {
    "Current Ratio": "Current Ratio",
    "Debt": "Debt / Equity Ratio",
    "EBITDA": "EBITDA",
    "Free Cash Flow": "Free Cash Flow (Millions)",
    "Inventory Turnover": "Inventory Turnover",
    "Net Income": "Net Income (Millions)"
}

# 抓季度財務比率
def fetch_quarterly_ratios(symbol):
    url = f"https://stockanalysis.com/stocks/{symbol.lower()}/financials/ratios/quarterly/"
    html = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}).text
    tables = pd.read_html(html, header=0)
    if not tables:
        return None
    df = tables[0]

    # 若是多層欄位先壓平
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [' '.join([str(c) for c in col if c and c != 'nan']).strip() for col in df.columns]

    df.rename(columns={df.columns[0]: "Metric"}, inplace=True)
    df = df[df["Metric"].str.contains("|".join(TARGET.keys()), case=False, na=False)]
    df["Metric"] = df["Metric"].apply(lambda x: next((v for k, v in TARGET.items() if k.lower() in x.lower()), x))
    df = df.replace(["Upgrade", "-", "—"], pd.NA)

    # 轉置
    df = df.set_index("Metric").T.reset_index().rename(columns={"index": "Date_1"})
    df["Date_1"] = df["Date_1"].apply(lambda x: re.sub(r"[\(\)'\"]+", "", str(x)).split(",")[-1].strip())
    df["Ticker"] = symbol
    return df

# 抓 Z/F Score
def fetch_scores(symbol):
    url = f"https://stockanalysis.com/stocks/{symbol.lower()}/statistics/"
    html = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}).text
    df = pd.concat(pd.read_html(html), ignore_index=True)
    df.columns = ["Metric", "Value"]
    z = df[df["Metric"].str.contains("Altman Z", na=False)]["Value"].values
    f = df[df["Metric"].str.contains("Piotroski F", na=False)]["Value"].values
    return {
        "Altman Z-Score 的平均": z[0] if len(z) else None,
        "Piotroski F-Score 的平均": f[0] if len(f) else None
    }

# 主程式
writer = pd.ExcelWriter("stock_data_quarterly.xlsx", engine="openpyxl")

for t in TICKERS:
    print(f"🔍 抓取 {t} (Quarterly) ...")
    ratios = fetch_quarterly_ratios(t)
    scores = fetch_scores(t)
    if ratios is not None:
        # 讓 Z/F 分數只顯示一次
        z_score = scores.get("Altman Z-Score 的平均")
        f_score = scores.get("Piotroski F-Score 的平均")

        # 建立一行的 summary DataFrame
        summary = pd.DataFrame({
            "Altman Z-Score 的平均": [z_score],
            "Piotroski F-Score 的平均": [f_score]
        })

        # 將 summary 寫在最上方，財務資料寫在下面
        summary.to_excel(writer, sheet_name=t, index=False, startrow=0)
        ratios.to_excel(writer, sheet_name=t, index=False, startrow=3)
        print(f"✅ {t} 完成，共 {len(ratios)} 期")
    else:
        print(f"⚠️ {t} 抓取失敗")

writer.close()
print("✅ 已輸出 stock_data_quarterly.xlsx (每個公司一個工作表)")


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:




