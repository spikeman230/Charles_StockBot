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

# === 1. 機密環境變數 ===
TG_TOKEN = os.environ.get("TG_TOKEN")
TG_CHAT_ID = os.environ.get("TG_CHAT_ID")
EMAIL_USER = os.environ.get("EMAIL_USER")  # 你的 Gmail
EMAIL_PASS = os.environ.get("EMAIL_PASS")  # Gmail 應用程式密碼 (16碼)
EMAIL_TO = os.environ.get("EMAIL_TO")      # 收件者信箱

# === 2. 專屬通訊錄 (18檔監控名單) ===
STOCK_DICT = {
    "🛡️ 核心持股 (重倉伺服器)": {"3037.TW": "欣興 (ABF載板)"},
    "🔥 潛力種子 (高頻寬觀察區)": {"3163.TW": "波若威", "5388.TW": "中磊", "3714.TW": "富采"},
    "👀 常態觀察區 (例行監控節點)": {"2330.TW": "台積電", "2317.TW": "鴻海", "0050.TW": "元大台灣50"},
    "💾 記憶體族群 (美光連動網域)": {"MU": "美光", "2408.TW": "南亞科", "3260.TW": "威剛", "8299.TW": "群聯"}
}

# === 3. 持久化日誌 ===
def write_noc_log(date, symbol, name, close_price, rsi, vol_status, status, alert, predict):
    log_filename = "noc_trading_log.csv"
    file_exists = os.path.exists(log_filename)
    with open(log_filename, mode='a', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["日期", "代號", "名稱", "收盤價", "RSI", "量能狀態", "趨勢狀態", "戰場預判", "行動指令"])
        writer.writerow([date, symbol, name, f"{close_price:.2f}", f"{rsi:.2f}", vol_status, status, predict, alert])

# === 4. 分析與預判模組 ===
def get_analysis_and_chart(symbol, name):
    try:
        stock = yf.Ticker(symbol)
        hist = stock.history(period="6mo")
        if len(hist) < 30: return None

        # 基本 MA 與 Volume
        hist['5MA'] = hist['Close'].rolling(window=5).mean()
        hist['20MA'] = hist['Close'].rolling(window=20).mean()
        hist['5VMA'] = hist['Volume'].rolling(window=5).mean()
        
        # RSI
        delta = hist['Close'].diff()
        gain = delta.clip(lower=0); loss = -1 * delta.clip(upper=0)
        ema_gain = gain.ewm(com=13, adjust=False).mean()
        ema_loss = loss.ewm(com=13, adjust=False).mean()
        hist['RSI'] = 100 - (100 / (1 + (ema_gain / ema_loss)))

        # 🔮 MACD 預判動能
        hist['EMA12'] = hist['Close'].ewm(span=12, adjust=False).mean()
        hist['EMA26'] = hist['Close'].ewm(span=26, adjust=False).mean()
        hist['MACD'] = hist['EMA12'] - hist['EMA26']
        hist['Signal'] = hist['MACD'].ewm(span=9, adjust=False).mean()
        hist['MACD_Hist'] = hist['MACD'] - hist['Signal']

        # 🔮 布林通道壓縮 (BB Squeeze)
        hist['STD20'] = hist['Close'].rolling(window=20).std()
        hist['BB_Upper'] = hist['20MA'] + (2 * hist['STD20'])
        hist['BB_Lower'] = hist['20MA'] - (2 * hist['STD20'])
        hist['BB_Width'] = (hist['BB_Upper'] - hist['BB_Lower']) / hist['20MA']

        # 繪圖 (存檔稍後給 Email 用)
        chart_file = f"{symbol}_chart.png"
        mpf.plot(hist[-60:], type='candle', style='yahoo', volume=True, 
                 mav=(5, 20), title=f"{name} ({symbol})", savefig=chart_file)

        return hist, chart_file, stock.news[0] if stock.news else None
    except: return None

# === 5. 發送模組 (Telegram + Email) ===
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
        
        # 夾帶所有股票的 K 線圖
        for chart in chart_files:
            if os.path.exists(chart):
                with open(chart, 'rb') as f:
                    img = MIMEImage(f.read(), name=os.path.basename(chart))
                    msg.attach(img)
        
        # 👇 新增：夾帶 CSV 日誌檔 (確保在雲端執行也能留存紀錄) 👇
        log_file = "noc_trading_log.csv"
        if os.path.exists(log_file):
            with open(log_file, 'r', encoding='utf-8-sig') as f:
                csv_part = MIMEText(f.read(), 'csv', 'utf-8-sig')
                csv_part.add_header('Content-Disposition', f'attachment; filename="{log_file}"')
                msg.attach(csv_part)
        
        # 連線發信
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(EMAIL_USER, EMAIL_PASS)
            server.send_message(msg)
        print("✅ Email 戰報 (含圖表與 CSV 日誌) 發送成功！")
    except Exception as e:
        print(f"❌ Email 發送失敗: {e}")

# === 6. 主程式執行 ===
if __name__ == "__main__":
    tw_tz = datetime.timezone(datetime.timedelta(hours=8))
    curr_date = datetime.datetime.now(tw_tz).date()
    curr_time = datetime.datetime.now(tw_tz).strftime("%Y-%m-%d %H:%M:%S")
    
    msg_list = []
    generated_charts = []
    has_data = False

    print(f"[{curr_time}] NOC 戰情室 v3.0 (預判版) 啟動...")

    for cat, stocks in STOCK_DICT.items():
        msg_list.append(f"━━━━━━━━━━━━━━\n📂 【{cat}】\n━━━━━━━━━━━━━━\n")
        for sym, name in stocks.items():
            res = get_analysis_and_chart(sym, name)
            if not res: continue
            hist, chart_file, news = res
            
            td = hist.iloc[-1]; yd = hist.iloc[-2]; yyd = hist.iloc[-3]
            last_date = hist.index[-1].date()
            
            # 台股休市防呆 (美股不影響)
            if last_date != curr_date and ".TW" in sym: continue
            
            has_data = True
            generated_charts.append(chart_file)

            # --- 基本狀態判
