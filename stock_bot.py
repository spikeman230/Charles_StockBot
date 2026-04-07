import yfinance as yf
import requests
import os
import datetime
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

# === 3. 分析、繪圖與情報抓取模組 ===
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
        hist['5VMA'] = hist['Volume'].rolling(window=5).mean()
        
        delta = hist['Close'].diff()
        gain = delta.clip(lower=0)
        loss = -1 * delta.clip(upper=0)
        ema_gain = gain.ewm(com=13, adjust=False).mean()
        ema_loss = loss.ewm(com=13, adjust=False).mean()
        rs = ema_gain / ema_loss
        hist['RSI'] = 100 - (100 / (1 + rs))
        
        td = hist.iloc[-1]
        yd = hist.iloc[-2]
        last_trade_date = hist.index[-1].date()
        
        # 抓取新聞情報
        latest_news = None
        try:
            news_list = stock.news
            if news_list and len(news_list) > 0:
                latest_news = news_list[0]
        except:
            pass 
        
        return td, yd, last_trade_date, chart_filename, latest_news
    except Exception as e:
        print(f"[{symbol}] 分析失敗: {e}")
        return None

# === 4. Telegram 發送模組 ===
def send_telegram_msg(msg):
    url = f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage"
    payload = {"chat_id": TG_CHAT_ID, "text": msg, "disable_web_page_preview": True}
    requests.post(url, json=payload)

# === 5. Email 發送模組 (若不需要可略過設定) ===
def send_email_report(subject, text_body, image_files):
    msg = MIMEMultipart()
    msg
