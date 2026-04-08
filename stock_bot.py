import yfinance as yf
import requests
import os
import datetime
import pandas as pd
import numpy as np
import csv
import mplfinance as mpf
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication

# === 1. 機密環境變數 ===
TG_TOKEN = os.environ.get("TG_TOKEN")
TG_CHAT_ID = os.environ.get("TG_CHAT_ID")
EMAIL_USER = os.environ.get("EMAIL_USER")
EMAIL_PASS = os.environ.get("EMAIL_PASS")
EMAIL_TO = os.environ.get("EMAIL_TO")
FINMIND_TOKEN = os.environ.get("FINMIND_TOKEN")

# === 2. 專屬通訊錄 (18檔監控名單) ===
STOCK_DICT = {
    "🛡️ 核心持股 (重倉伺服器)": {"3037.TW": "欣興 (ABF載板)"},
    "🔥 潛力種子 (高頻寬觀察區)": {"3163.TW": "波若威", "5388.TW": "中磊", "3714.TW": "富采"},
    "👀 常態觀察區 (例行監控節點)": {"2330.TW": "台積電", "2317.TW": "鴻海", "0050.TW": "元大台灣50"},
    "💾 記憶體族群 (美光連動網域)": {"MU": "美光", "2408.TW": "南亞科", "3260.TW": "威剛", "8299.TW": "群聯", "AAPL": "蘋果", "NVDA": "輝達"}
}

# === 3. 持久化日誌 ===
def write_noc_log(date, symbol, name, close_price, rsi, vol_status, status, alert, predict, chip_signal):
    log_filename = "noc_trading_log.csv"
    file_exists = os.path.exists(log_filename)
    with open(log_filename, mode='a', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["日期", "代號", "名稱", "收盤價", "RSI", "量能狀態", "趨勢狀態", "戰場預判", "籌碼訊號", "行動指令"])
        writer.writerow([date, symbol, name, f"{close_price:.2f}", f"{rsi:.2f}", vol_status, status, predict, chip_signal, alert])

# === 4. 🔌 FinMind API 串接模組 ===
def get_finmind_chip_data(symbol, start_date_str):
    if not FINMIND_TOKEN:
        return pd.DataFrame()
    
    fm_symbol = symbol.replace(".TW", "").replace(".TWO", "")
    if not fm_symbol.isdigit():
        return pd.DataFrame()

    url = "https://api.finmindtrade.com/api/v4/data"
    params = {
        "dataset": "TaiwanStockInstitutionalInvestorsBuySell",
        "data_id": fm_symbol,
        "start_date": start_date_str,
        "token": FINMIND_TOKEN
    }
    
    try:
        r = requests.get(url, params=params, timeout=10)
        data = r.json()
        if data.get("msg") == "success" and len(data.get("data", [])) > 0:
            df = pd.DataFrame(data["data"])
            df['net_buy'] = df['buy'] - df['sell']
            
            df['type'] = 'Other'
            df.loc[df['name'].str.contains('外資'), 'type'] = 'Foreign_Inv'
            df.loc[df['name'].str.contains('投信'), 'type'] = 'Trust_Inv'
            df.loc[df['name'].str.contains('自營商'), 'type'] = 'Dealer_Inv'
            
            pivot_df = df.groupby(['date', 'type'])['net_buy'].sum().unstack(fill_value=0).reset_index()
            
            for col in ['Foreign_Inv', 'Trust_Inv', 'Dealer_Inv']:
                if col not in pivot_df.columns: pivot_df[col] = 0
                    
            pivot_df['Date'] = pd.to_datetime(pivot_df['date']).dt.date
            pivot_df.set_index('Date', inplace=True)
            return pivot_df[['Foreign_Inv', 'Trust_Inv', 'Dealer_Inv']]
    except Exception as e:
        print(f"[{symbol}] FinMind 籌碼撈取失敗: {e}")
        
    return pd.DataFrame()

# === 5. Phase 2 籌碼運算模組 ===
def calculate_chip_signals(hist: pd.DataFrame) -> pd.DataFrame:
    required_chip_cols = ['Foreign_Inv', 'Trust_Inv', 'Dealer_Inv']
    hist['Chip_Status'] = "無資料"
    
    if all(col in hist.columns for col in required_chip_cols):
        hist['Total_Institutional'] = hist['Foreign_Inv'] + hist['Trust_Inv'] + hist['Dealer_Inv']
        hist['Foreign_Buy_Flag'] = (hist['Foreign_Inv'] > 0).astype(int)
        hist['Trust_Buy_Flag'] = (hist['Trust_Inv'] > 0).astype(int)
        hist['Trust_Buy_Days_5d'] = hist['Trust_Buy_Flag'].rolling(window=5).sum()
        
        hist['Signal_CoBuy'] = (hist['Foreign_Inv'] > 0) & (hist['Trust_Inv'] > 0)
        hist['Signal_Trust_Trend'] = (hist['Trust_Buy_Days_5d'] >= 4) & (hist['Trust_Buy_Flag'] == 1)
        
        conditions = [
            (hist['Signal_CoBuy'] == True),
            (hist['Signal_Trust_Trend'] == True),
            (hist['Total_Institutional'] > 0)
        ]
        choices = ["🤝 土洋齊買", "🏦 投信連買", "📈 法人偏多"]
        hist['Chip_Status'] = np.select(conditions, choices, default="➖ 中性/偏空")
        
    return hist

# === 6. 分析與預判模組 ===
def get_analysis_and_chart(symbol, name):
    try:
        stock = yf.Ticker(symbol)
        hist = stock.history(period="6mo")
        if len(hist) < 30: return None

        hist['Date_Key'] = hist.index.strftime('%Y-%m-%d')
        
        if FINMIND_TOKEN and (".TW" in symbol or ".TWO" in symbol):
            start_date_str = (datetime.datetime.now() - datetime.timedelta(days=180)).strftime("%Y-%m-%d")
            chip_df = get_finmind_chip_data(symbol, start_date_str)
            if not chip_df.empty:
                # 確保 index 是 date object 以便與 hist 的 index 比較或合併
                hist.index = hist.index.date
                hist = hist.join(chip_df, how='left')
                hist.fillna({'Foreign_Inv': 0, 'Trust_Inv': 0, 'Dealer_Inv': 0}, inplace=True)

        hist = calculate_chip_signals(hist)

        hist['5MA'] = hist['Close'].rolling(window=5).mean()
        hist['20MA'] = hist['Close'].rolling(window=20).mean()
        hist['5VMA'] = hist['Volume'].rolling(window=5).mean()
        
        delta = hist['Close'].diff()
        gain = delta.clip(lower=0); loss = -1 * delta.clip(upper=0)
        ema_gain = gain.ewm(com=13, adjust=False).mean()
        ema_loss = loss.ewm(com=13, adjust=False).mean()
        hist['RSI'] = 100 - (100 / (1 + (ema_gain / ema_loss)))

        hist['EMA12'] = hist['Close'].ewm(span=12, adjust=False).mean()
        hist['EMA26'] = hist['Close'].ewm(span=26, adjust=False).mean()
        hist['MACD'] = hist['EMA12'] - hist['EMA26']
        hist['Signal'] = hist['MACD'].ewm(span=9, adjust=False).mean()
        hist['MACD_Hist'] = hist['MACD'] - hist['Signal']

        hist['STD20'] = hist['Close'].rolling(window=20).std()
        hist['BB_Upper'] = hist['20MA'] + (2 * hist['STD20'])
        hist['BB_Lower'] = hist['20MA'] - (2 * hist['STD20'])
        hist['BB_Width'] = (hist['BB_Upper'] - hist['BB_Lower']) / hist['20MA']

        chart_file = f"{symbol}_chart.png"
        
        # 準備繪圖用的資料格式
        plot_df = hist.copy()
        plot_df.index = pd.to_datetime(plot_df.index)
        
        mpf.plot(plot_df[-60:], type='candle', style='yahoo', volume=True, 
                 mav=(5, 20), title=f"{name} ({symbol})", savefig=chart_file)

        return hist, chart_file, stock.news[0] if stock.news else None
    except Exception as e:
        print(f"[{symbol}] 分析過程發生錯誤: {e}")
        return None

# === 7. 發送模組 (Telegram + Email) ===
def send_telegram(msg):
    if TG_TOKEN and TG_CHAT_ID:
        requests.post(f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage", 
                      json={"chat_id": TG_CHAT_ID, "text": msg, "disable_web_page_preview": True})

def send_email_report(subject, text_body, chart_files):
    if not EMAIL_USER or not EMAIL_PASS or not EMAIL_TO:
        print("未設定 Email 環境變數，略過發送。")
        return
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_USER
        msg['To'] = EMAIL_TO
        msg['Subject'] = subject
        msg.attach(MIMEText(text_body, 'plain'))
        
        for chart in chart_files:
            if os.path.exists(chart):
                with open(chart, 'rb') as f:
                    # 這行是剛剛出錯的地方，現在已經
