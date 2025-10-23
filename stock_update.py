import pandas as pd
import requests, re
from openpyxl import Workbook

# -------------------------------------------------------
# 公司代碼與名稱
# -------------------------------------------------------
TICKERS = {
    "AA": "Alcoa",
    "RIO": "Rio Tinto",
    "OSL:NHY": "Norsk Hydro",        # ✅ 正確代碼
    "RS": "Reliance Steel & Aluminum",
    "KALU": "Kaiser Aluminum",
    "RYI": "Ryerson Holding"
}

# -------------------------------------------------------
# 欄位名稱對應
# -------------------------------------------------------
TARGET = {
    "EBITDA": "EBITDA",
    "Debt": "Debt / Equity Ratio",
    "Inventory Turnover": "Inventory Turnover",
    "Current Ratio": "Current Ratio"
}

# -------------------------------------------------------
# 把代碼轉成 stockanalysis 用的網址格式，例如：
# OSL:NHY → osl-nhy
# -------------------------------------------------------
def format_symbol(symbol):
    return symbol.lower
