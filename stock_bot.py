import yfinance as yf
import requests
import os
import datetime
import pandas as pd
import mplfinance as mpf
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

# === 1. 從 GitHub 保險箱抓取機密 ===
TG_TOKEN = os.environ.get("TG_TOKEN")
TG_CHAT_ID = os.environ.get("TG_CHAT_ID")
EMAIL_USER = os.environ.get("EMAIL_USER")
EMAIL_PASS = os.environ.get("EMAIL_PASS")
EMAIL_TO = os.environ.get("EMAIL_TO")
FINMIND_TOKEN = os.environ.get("FINMIND_TOKEN") # 👈 Phase 2 新增：籌碼大門金鑰

# === 2. 專屬通訊錄 (四層 VLAN 網域版) ===
STOCK_DICT = {
    "🛡️ 核心持股 (重倉伺服器)": {
        "3037.TW": "欣興 (ABF載板)"
    },
    "🔥 潛力種子 (高頻寬觀察區)": {
        "3163.TW": "波若威 (光通訊)",
        "5388.TW": "中磊 (網通設備)",
        "3714.TW": "富采 (LED光電)"
    },
    "👀 常態觀察區 (例行監控節點)": {
        "2330.TW": "台積電",
        "2317.TW": "鴻海",
        "0050.TW": "元大台灣50",
        "009816.TW": "凱基台灣TOP50",
        "8431.TWO": "匯鑽科",
        "AAPL": "蘋果 (Apple)",
        "NVDA": "輝達 (NVIDIA)"
    },
    "💾 記憶體族群 (美光連動網域)": {
        "MU"  : "美光 (Micron)", 
        "2408.TW": "南亞科 (DRAM製造)",
        "2344.TW": "華邦電 (利基記憶體)",
        "6239.TW": "力成 (記憶體封測)", 
        "3260.TW": "威剛 (記憶體模組)", 
        "8299.TW": "群聯 (Flash控制IC)", 
        "4967.TW": "十銓 (電競模組)"
    }
}

# === 3. 籌碼資料抓取模組 (FinMind API) ===
def get_trust_data(symbol):
    # 如果沒有 Token 或不是台股，直接跳過防呆
    if not FINMIND_TOKEN or (".TW" not in symbol and ".TWO" not in symbol):
        return 0
    
    stock_id = symbol.replace(".TW", "").replace(".TWO", "")
    # 抓取過去 10 天的資料，確保能取到最近的 3 個交易日
    start_date = (datetime.datetime.now() - datetime.timedelta(days=10)).strftime("%Y-%m-%d")
    
    url = "https://api.finmindtrade.com/api/v4/data"
    params = {
        "dataset": "TaiwanStockInstitutionalInvestorsBuySell",
        "data_id": stock_id,
        "start_date": start_date,
        "token": FINMIND_TOKEN
    }
    
    try:
        resp = requests.get(url, params=params, timeout=10)
        data = resp.json()
        if data.get("msg") == "success" and "data" in data:
            df = pd.DataFrame(data["data"])
            if df.empty: return 0
            
            # 篩選投信 (Investment_Trust)
            trust_df = df[df['name'] == 'Investment_Trust'].copy()
            if trust_df.empty: return 0
            
            # 依日期排序並取最近 3 筆 (近 3 日)
            trust_df = trust_df.sort_values(by='date')
            last_3 = trust_df.tail(3)
            
            # 計算近 3 日買賣超張數 (API 給的是「股」數，除以 1000 變為「張」)
            net_buy_shares = last_3['buy'].sum() - last_3['sell'].sum()
            net_buy_lots = net_buy_shares / 1000
            return round(net_buy_lots)
    except Exception as e:
        print(f"[{symbol}] FinMind 籌碼撈取失敗: {e}")
        return 0
    return 0

# === 4. 分析、繪圖與情報抓取模組 ===
def get_analysis_and_draw_chart(symbol, name):
    try:
        stock = yf.Ticker(symbol)
        hist = stock.history(period="6mo")
        if len(hist) < 30: return None
        
        chart_filename = f"{symbol}_chart.png"
        mpf.plot(hist, type='candle', style='yahoo', volume=True, 
                 mav=(5, 20), title=f"{name} ({symbol})", 
                 savefig=chart_filename)
        
        hist['5MA'] = hist['Close'].rolling(window=5).mean()
        hist['20MA'] = hist['Close'].rolling(window=20).mean()
        hist['5VMA'] = hist['Volume'].rolling(
