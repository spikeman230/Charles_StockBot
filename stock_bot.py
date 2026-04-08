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

# === 4. FinMind API 串接 ===
def get_finmind_chip_data(symbol, start_date_str):
    if not FINMIND_TOKEN: return pd.DataFrame()
    fm_symbol = symbol.replace(".TW", "").replace(".TWO", "")
    if not fm_symbol.isdigit(): return pd.DataFrame()
    url = "https://api.finmindtrade.com/api/v4/data"
    params = {"dataset": "TaiwanStockInstitutionalInvestorsBuySell", "data_id": fm_symbol, "start_date": start_date_str, "token": FINMIND_TOKEN}
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
    except: pass
    return pd.DataFrame()

# === 5. 分析模組 (含狙擊模式邏輯) ===
def get_analysis_and_chart(symbol, name):
    try:
        stock = yf.Ticker(symbol)
        hist = stock.history(period="6mo")
        if len(hist) < 30: return None

        # 基礎指標與 MA
        hist['5MA'] = hist['Close'].rolling(5).mean()
        hist['20MA'] = hist['Close'].rolling(20).mean()
        hist['5VMA'] = hist['Volume'].rolling(5).mean()
        
        # MACD
        hist['EMA12'] = hist['Close'].ewm(span=12, adjust=False).mean()
        hist['EMA26'] = hist['Close'].ewm(span=26, adjust=False).mean()
        hist['MACD'] = hist['EMA12'] - hist['EMA26']
        hist['Signal'] = hist['MACD'].ewm(span=9, adjust=False).mean()
        hist['MACD_Hist'] = hist['MACD'] - hist['Signal']

        # 布林通道
        hist['STD20'] = hist['Close'].rolling(20).std()
        hist['BB_Width'] = (hist['20MA']*4*hist['STD20'])/hist['20MA'] # 簡化
        hist['BB_Width'] = (hist['Close'].rolling(20).std() * 4) / hist['20MA']

        # RSI
        delta = hist['Close'].diff()
        gain = delta.clip(lower=0); loss = -1 * delta.clip(upper=0)
        hist['RSI'] = 100 - (100 / (1 + (gain.ewm(com=13).mean() / loss.ewm(com=13).mean())))

        # 🔌 籌碼合併
        hist['Date_Key'] = hist.index.date
        if FINMIND_TOKEN and (".TW" in symbol or ".TWO" in symbol):
            start_str = (datetime.datetime.now() - datetime.timedelta(days=180)).strftime("%Y-%m-%d")
            chip = get_finmind_chip_data(symbol, start_str)
            if not chip.empty:
                hist = hist.merge(chip, left_on='Date_Key', right_index=True, how='left').fillna(0)

        # ---------------------------------------------------------
        # 🎯 核心：狙擊模式偵測
        # ---------------------------------------------------------
        # 定義築底：價格在5MA下且MACD綠柱連三天收斂 (MACD_Hist < 0 且連續上升)
        hist['Is_Bottoming'] = (hist['Close'] < hist['5MA']) & \
                               (hist['MACD_Hist'].shift(2) < hist['MACD_Hist'].shift(1)) & \
                               (hist['MACD_Hist'].shift(1) < hist['MACD_Hist']) & \
                               (hist['MACD_Hist'] < 0)
        
        # 檢查過去 3 天(含今天) 是否有築底訊號
        hist['Recent_Bottoming'] = hist['Is_Bottoming'].rolling(window=3).max().astype(bool)

        chart_file = f"{symbol}_chart.png"
        mpf.plot(hist[-60:], type='candle', style='yahoo', volume=True, mav=(5, 20), savefig=chart_file)
        return hist, chart_file
    except: return None

# === 6. 發送模組 (修復 CSV 夾帶) ===
def send_reports(subject, msg, charts):
    # Telegram
    if TG_TOKEN: requests.post(f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage", json={"chat_id": TG_CHAT_ID, "text": msg})
    # Email
    if EMAIL_USER and EMAIL_PASS:
        try:
            m = MIMEMultipart(); m['From'] = EMAIL_USER; m['To'] = EMAIL_TO; m['Subject'] = subject
            m.attach(MIMEText(msg, 'plain'))
            for c in charts:
                if os.path.exists(c):
                    with open(c, 'rb') as f: m.attach(MIMEImage(f.read(), name=os.path.basename(c)))
            if os.path.exists("noc_trading_log.csv"):
                with open("noc_trading_log.csv", 'rb') as f:
                    part = MIMEApplication(f.read(), Name="noc_trading_log.csv")
                    part.add_header('Content-Disposition', 'attachment; filename="noc_trading_log.csv"')
                    m.attach(part)
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
                s.login(EMAIL_USER, EMAIL_PASS); s.send_message(m)
        except Exception as e: print(f"Email Error: {e}")

# === 7. 主程式 ===
if __name__ == "__main__":
    tw_tz = datetime.timezone(datetime.timedelta(hours=8))
    curr_date = datetime.datetime.now(tw_tz).date()
    curr_time = datetime.datetime.now(tw_tz).strftime("%Y-%m-%d %H:%M:%S")
    msg_list = []; charts = []; has_data = False

    for cat, stocks in STOCK_DICT.items():
        msg_list.append(f"━━━━━━━━━━━━━━\n📂 【{cat}】")
        for sym, name in stocks.items():
            res = get_analysis_and_chart(sym, name)
            if not res: continue
            hist, chart_file = res
            td = hist.iloc[-1]; yd = hist.iloc[-2]
            if td.name.date() != curr_date and ".TW" in sym: continue
            has_data = True; charts.append(chart_file)

            # 狀態判斷
            vol_status = "📈 出量" if td['Volume'] > td['5VMA']*1.2 else "📉 量縮"
            trend = "🔥 多頭" if td['Close'] > td['5MA'] > td['20MA'] else "🧊 空頭" if td['Close'] < td['5MA'] < td['20MA'] else "🔄 盤整"
            
            # 🔮 預判
            predict = "⚠️ 大變盤預警" if td['BB_Width'] < 0.08 else "築底醞釀中" if td['Is_Bottoming'] else "無特殊徵兆"
            
            # 🛡️ 狙擊指令邏輯
            # 條件：近 3 日有築底訊號 + 今日帶量突破 5MA
            is_breakout = (yd['Close'] < yd['5MA']) and (td['Close'] > td['5MA']) and (td['Volume'] > td['5VMA']*1.1)
            
            if td['Recent_Bottoming'] and is_breakout:
                alert = "🚀【狙擊模式：強烈買進】底部完成且帶量突破！"
            elif td['RSI'] > 80:
                alert = "💰【獲利了結】短線過熱"
            elif td['Close'] < td['5MA'] < td['20MA']:
                alert = "💀【強制退場】空頭確認"
            else:
                alert = "✅【持股續抱】順勢操作"

            # 寫入日誌
            write_noc_log(curr_date, sym, name, td['Close'], td['RSI'], vol_status, trend, alert, predict, "N/A")
            
            msg_list.append(f"🔸 {name} ({sym})\n   現價: {td['Close']:.2f} | RSI: {td['RSI']:.1f}\n   狀態: {trend} | {vol_status}\n   🔮 預判: {predict}\n   👉 指令: {alert}\n")

    if has_data:
        final_msg = f"📡 【NOC 戰情室 v4.0：嚴格狙擊版】\n📅 時間：{curr_time}\n" + "\n".join(msg_list)
        send_reports(f"NOC 戰情報告 {curr_date}", final_msg, charts)
        for c in charts: 
            if os.path.exists(c): os.remove(c)
