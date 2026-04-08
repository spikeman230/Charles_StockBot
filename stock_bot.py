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
EMAIL_PASS = os.environ.get("EMAIL_PASS")  # Gmail 應用程式密碼
EMAIL_TO = os.environ.get("EMAIL_TO")      # 收件者信箱

# === 2. 專屬通訊錄 ===
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

        # 🔮 新增：MACD 預判動能
        hist['EMA12'] = hist['Close'].ewm(span=12, adjust=False).mean()
        hist['EMA26'] = hist['Close'].ewm(span=26, adjust=False).mean()
        hist['MACD'] = hist['EMA12'] - hist['EMA26']
        hist['Signal'] = hist['MACD'].ewm(span=9, adjust=False).mean()
        hist['MACD_Hist'] = hist['MACD'] - hist['Signal']

        # 🔮 新增：布林通道壓縮 (BB Squeeze)
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
        
        for chart in chart_files:
            if os.path.exists(chart):
                with open(chart, 'rb') as f:
                    img = MIMEImage(f.read(), name=os.path.basename(chart))
                    msg.attach(img)
        
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(EMAIL_USER, EMAIL_PASS)
            server.send_message(msg)
        print("✅ Email 戰報夾帶圖表發送成功！")
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
            if last_date != curr_date and ".TW" in sym: continue
            
            has_data = True
            generated_charts.append(chart_file)

            # --- 基本狀態判斷 ---
            vol_today = td['Volume']; vma5 = td['5VMA']
            vol_status = "📈 出量" if vol_today > vma5 * 1.2 else "📉 量縮" if vol_today < vma5 * 0.8 else "➖ 量平"
            trend_status = "🔥 多頭" if td['Close'] > td['5MA'] > td['20MA'] else "🧊 空頭" if td['Close'] < td['5MA'] < td['20MA'] else "🔄 盤整"

            # --- 🔮 預判雷達邏輯 (Pre-cognition) ---
            predict_msg = "無特殊徵兆"
            
            # 1. 布林壓縮預警 (寬度小於 8%)
            if td['BB_Width'] < 0.08:
                predict_msg = "⚠️【大變盤預警】布林通道極度壓縮，即將表態，請盯緊方向！"
            # 2. 底部背離預判 (價格跌破5MA，但MACD綠柱連續三天縮短)
            elif td['Close'] < td['5MA'] and yyd['MACD_Hist'] < yd['MACD_Hist'] < td['MACD_Hist'] < 0:
                predict_msg = "📈【築底預判】空方動能衰退(MACD收斂)，醞釀反彈契機！"
            # 3. 高檔背離預判 (價格創新高，但MACD紅柱開始縮短)
            elif td['Close'] > td['5MA'] and yyd['MACD_Hist'] > yd['MACD_Hist'] > td['MACD_Hist'] > 0:
                predict_msg = "📉【見頂預判】多方動能衰退，小心假突破真倒貨！"

            # --- 🛡️ 行動指令邏輯 (維持紀律) ---
            if td['Close'] < td['5MA'] < td['20MA'] and vol_today > vma5 * 1.2:
                alert = "💀【強制退場】大單狂砸！立刻清倉保命！"
            elif yd['Close'] < yd['5MA'] and td['Close'] > td['5MA'] and vol_today > vma5 * 1.2:
                alert = "🚀【強烈買進】出量站回5日線！立刻進場！"
            elif td['RSI'] > 80:
                alert = "💰【獲利了結】RSI過熱，分批獲利入袋！"
            else:
                alert = "✅【持股續抱】順勢操作，等待訊號。"

            # 寫入日誌
            write_noc_log(curr_date, sym, name, td['Close'], td['RSI'], vol_status, trend_status, predict_msg, alert)

            # 排版 (Telegram & Email 共用)
            stock_msg = f"🔸 {name} ({sym})\n"
            stock_msg += f"   現價: {td['Close']:.2f} | RSI: {td['RSI']:.1f}\n"
            stock_msg += f"   狀態: {trend_status} | {vol_status}\n"
            stock_msg += f"   🔮 預判: {predict_msg}\n"
            stock_msg += f"   👉 指令: {alert}\n\n"
            msg_list.append(stock_msg)

    # --- 戰報發送 ---
    if has_data:
        final_text = f"📡 【老網管 NOC 指揮中心：綜合戰報】\n📅 時間：{curr_time}\n━━━━━━━━━━━━━━\n" + "".join(msg_list)
        
        send_telegram(final_text)
        send_email_report(f"NOC 戰情報告 {curr_date}", final_text, generated_charts)
        
        # 清理產生出來的圖片垃圾
        for chart in generated_charts:
            if os.path.exists(chart): os.remove(chart)
    else:
        print("休市，伺服器待命。")
