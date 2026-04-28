import yfinance as yf
import requests
import os
import datetime
import pandas as pd
import numpy as np
import csv
import json
import math
import re
import mplfinance as mpf
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication

# === 1. 機密環境變數 (完整保留) ===
TG_TOKEN = os.environ.get("TG_TOKEN")
TG_CHAT_ID = os.environ.get("TG_CHAT_ID")
EMAIL_USER = os.environ.get("EMAIL_USER")
EMAIL_PASS = os.environ.get("EMAIL_PASS")
EMAIL_TO = os.environ.get("EMAIL_TO")
FINMIND_TOKEN = os.environ.get("FINMIND_TOKEN")
TRELLO_KEY = os.getenv('TRELLO_KEY')
TRELLO_TOKEN = os.getenv('TRELLO_TOKEN')
TRELLO_BOARD_ID = os.getenv('TRELLO_BOARD_ID')

# === 1.1 量化風控參數 (v8.3 原始參數回歸) ===
TOTAL_CAPITAL = 1000000 
RISK_PER_TRADE = 0.02    
ATR_MULTIPLIER = 2.0     
YOY_EXPLOSION_PCT = 50.0 
PE_LIMIT = 40.0          
SILENT_MODE = True       

# === 2. 🌟 Trello 雲端控制引擎 (完整 573 行版邏輯) ===
def fetch_trello_deployment():
    trello_dict = {}; my_portfolio = {}
    if not TRELLO_KEY or not TRELLO_TOKEN or not TRELLO_BOARD_ID:
        print("⚠️ 未偵測到 Trello 金鑰，將啟用預設部署。")
        return None, None
    url = f"https://api.trello.com/1/boards/{TRELLO_BOARD_ID}/lists"
    params = {"cards": "open", "key": TRELLO_KEY, "token": TRELLO_TOKEN}
    try:
        response = requests.get(url, params=params, timeout=15)
        if response.status_code == 200:
            lists_data = response.json()
            for lst in lists_data:
                list_name = lst['name']
                cards = lst.get('cards', [])
                if "庫存" in list_name:
                    for card in cards:
                        raw_name = card['name'].strip()
                        ticker_match = re.match(r'^[A-Za-z0-9.]+', raw_name)
                        symbol = ticker_match.group() if ticker_match else raw_name
                        name = raw_name[len(symbol):].strip() if ticker_match else symbol
                        desc = card.get('desc', '')
                        buy_price = 0.0; shares = 1000
                        p_match = re.search(r"成本[：:]\s*([0-9.]+)", desc)
                        s_match = re.search(r"股數[：:]\s*([0-9]+)", desc)
                        if p_match: buy_price = float(p_match.group(1))
                        if s_match: shares = int(s_match.group(1))
                        my_portfolio[symbol] = {"name": name if name else symbol, "buy_price": buy_price, "shares": shares}
                else:
                    stock_list = {}
                    for card in cards:
                        raw_name = card['name'].strip()
                        ticker_match = re.match(r'^[A-Za-z0-9.]+', raw_name)
                        symbol = ticker_match.group() if ticker_match else raw_name
                        name = raw_name[len(symbol):].strip() if ticker_match else symbol
                        stock_list[symbol] = name if name else symbol
                    if stock_list: trello_dict[list_name] = stock_list
            return trello_dict, my_portfolio
    except Exception as e: print(f"⚠️ Trello 連線失敗: {e}")
    return None, None

# === 3. 核心功能: 發報機 (✅ 確保 send_reports 完整且可見) ===
def send_reports(subject, text_body, chart_files):
    if TG_TOKEN and TG_CHAT_ID:
        try:
            requests.post(f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage", 
                          json={"chat_id": TG_CHAT_ID, "text": text_body, "disable_web_page_preview": True}, timeout=15)
        except Exception as e: print(f"❌ Telegram 發送失敗: {e}")
    
    if EMAIL_USER and EMAIL_PASS and EMAIL_TO:
        try:
            msg = MIMEMultipart()
            msg['From'] = EMAIL_USER; msg['To'] = EMAIL_TO; msg['Subject'] = subject
            msg.attach(MIMEText(text_body, 'plain'))
            for chart in chart_files:
                if os.path.exists(chart):
                    with open(chart, 'rb') as f: msg.attach(MIMEImage(f.read(), name=os.path.basename(chart)))
            log_file = "noc_trading_log.csv"
            if os.path.exists(log_file):
                with open(log_file, 'rb') as f:
                    csv_part = MIMEApplication(f.read(), Name=log_file)
                    csv_part.add_header('Content-Disposition', f'attachment; filename="{log_file}"')
                    msg.attach(csv_part)
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
                server.login(EMAIL_USER, EMAIL_PASS); server.send_message(msg)
        except Exception as e: print(f"❌ Email 發送失敗: {e}")

# === 4. 環境感知與數據引擎 (完整回歸) ===
def get_market_regime():
    try:
        twii = yf.Ticker("^TWII").history(period="1mo")
        twii['20MA'] = twii['Close'].rolling(20).mean()
        status = "🟢 多頭格局 (站上月線)" if twii['Close'].iloc[-1] > twii['20MA'].iloc[-1] else "🔴 空頭警戒 (跌破月線)"
        return twii['Close'].iloc[-1] > twii['20MA'].iloc[-1], status
    except: return True, "🟡 大盤狀態未知"

def get_revenue_yoy(symbol):
    if not FINMIND_TOKEN: return "N/A"
    match = re.search(r'\d+', symbol)
    if not match: return "N/A"
    try:
        url = "https://api.finmindtrade.com/api/v4/data"
        params = {"dataset": "TaiwanStockMonthRevenue", "data_id": match.group(), "start_date": (datetime.datetime.now() - datetime.timedelta(days=400)).strftime("%Y-%m-%d"), "token": FINMIND_TOKEN}
        r = requests.get(url, params=params, timeout=10)
        df = pd.DataFrame(r.json()["data"])
        latest = df.iloc[-1]; prev = df[(df['revenue_year'] == latest['revenue_year']-1) & (df['revenue_month'] == latest['revenue_month'])].iloc[-1]
        return float(((latest['revenue'] - prev['revenue']) / prev['revenue']) * 100)
    except: return "N/A"

def get_pe_ratio(symbol):
    try:
        info = yf.Ticker(symbol).info
        return info.get('trailingPE', info.get('forwardPE', "N/A"))
    except: return "N/A"

# === 5. 分析核心 (加回 KD/RSI/MACD/ATR 完整計算) ===
def get_analysis_and_chart(symbol, name):
    try:
        stock = yf.Ticker(symbol)
        hist = stock.history(period="8mo").dropna(subset=['Close'])
        if len(hist) < 40: return None
        # MA 計算
        hist['5MA'] = hist['Close'].rolling(5).mean(); hist['20MA'] = hist['Close'].rolling(20).mean()
        hist['5VMA'] = hist['Volume'].rolling(5).mean()
        # ATR 計算
        h_l = hist['High'] - hist['Low']; h_pc = abs(hist['High'] - hist['Close'].shift(1)); l_pc = abs(hist['Low'] - hist['Close'].shift(1))
        hist['ATR'] = pd.concat([h_l, h_pc, l_pc], axis=1).max(axis=1).rolling(14).mean()
        # RSI 計算
        delta = hist['Close'].diff(); gain = delta.clip(lower=0); loss = -delta.clip(upper=0)
        hist['RSI'] = 100 - (100 / (1 + (gain.ewm(com=13).mean() / loss.ewm(com=13).mean())))
        # KD 計算
        l9 = hist['Low'].rolling(9).min(); h9 = hist['High'].rolling(9).max()
        hist['K'] = ((hist['Close']-l9)/(h9-l9)*100).ewm(com=2).mean(); hist['D'] = hist['K'].ewm(com=2).mean()
        # MACD 計算
        hist['MACD'] = hist['Close'].ewm(span=12).mean() - hist['Close'].ewm(span=26).mean()
        hist['MACD_Hist'] = hist['MACD'] - hist['MACD'].ewm(span=9).mean()
        # 影線比
        hist['Shadow_Ratio'] = (hist['High'] - hist[['Open','Close']].max(axis=1)) / (hist['High']-hist['Low']).replace(0, 0.001)
        # 繪圖
        cf = f"{symbol}_chart.png"
        mc = mpf.make_marketcolors(up='red', down='green', edge='black', wick='black', volume='gray')
        try: s = mpf.make_mpf_style(base_mpf_style='yahoo', marketcolors=mc)
        except: s = mpf.make_mpf_style(marketcolors=mc)
        mpf.plot(hist[-60:], type='candle', style=s, volume=True, mav=(5, 20), savefig=cf)
        return hist, cf
    except: return None

# === 8. 主程式執行 (恢復 143633.JPG 格式並完整顯示數據) ===
if __name__ == "__main__":
    tw_tz = datetime.timezone(datetime.timedelta(hours=8))
    curr_time = datetime.datetime.now(tw_tz).strftime("%Y-%m-%d %H:%M:%S")
    msg_list = []; gen_charts = []; has_data = False
    
    TRELLO_DICT, MY_PORTFOLIO = fetch_trello_deployment()
    is_bull, m_status = get_market_regime()
    msg_list.append(f"🌐 【大盤風向】: {m_status}\n")

    # A. 💼 實體持股櫃 (顯示損益與 ATR 防線)
    if MY_PORTFOLIO:
        msg_list.append("━━━━━━━━━━━━━━\n💼 【庫存機櫃 (實體持股)】\n━━━━━━━━━━━━━━\n")
        for sym, data in MY_PORTFOLIO.items():
            res = get_analysis_and_chart(sym, data['name'])
            if not res: continue
            hist, chart = res; td = hist.iloc[-1]; has_data = True; gen_charts.append(chart)
            roi = ((td['Close'] - data['buy_price']) / data['buy_price']) * 100
            # ATR 動態防守
            stop_line = td['Close'] - (td['ATR'] * ATR_MULTIPLIER)
            pnl_tag = "🔥獲利巡航" if roi > 0 else "🩸破網" if td['Close'] < stop_line else "🟡暫時浮虧"
            msg_list.append(f"⚙️ {data['name']} ({sym})\n   損益: {roi:+.2f}% | 乖離: {((td['Close']-td['20MA'])/td['20MA']*100):+.1f}%\n   👉 {pnl_tag} (防線:{stop_line:.1f})\n\n")

    # B. 📂 觀察戰區 (顯示 KD、RSI、預判)
    for cat, stocks in (TRELLO_DICT.items() if TRELLO_DICT else []):
        msg_list.append(f"━━━━━━━━━━━━━━\n📂 【{cat}】\n━━━━━━━━━━━━━━\n")
        for sym, name in stocks.items():
            res = get_analysis_and_chart(sym, name)
            if not res: continue
            hist, chart = res; td = hist.iloc[-1]; has_data = True; gen_charts.append(chart)
            yoy = get_revenue_yoy(sym); pe = get_pe_ratio(sym)
            predict = "🚀動能轉強" if td['MACD_Hist'] > 0 and td['Close'] > td['5MA'] else "➖盤整"
            if td['Shadow_Ratio'] > 0.5: predict = "⚠️避雷針陷阱"
            msg_list.append(f"⚙️ {name} ({sym})\n   現價: {td['Close']:.2f} | K:{td['K']:.0f} D:{td['D']:.0f}\n   指標: RSI:{td['RSI']:.1f} | YoY:{yoy if isinstance(yoy, str) else f'{yoy:+.0f}%'}\n   👉 指令: {'🚀進場' if td['K']>td['D'] and td['Close']>td['5MA'] else '✅續抱'} | {predict}\n\n")

    if has_data:
        send_reports(f"NOC 戰報", f"📡 NOC v8.6\n📅 {curr_time}\n" + "".join(msg_list), gen_charts)
        for c in gen_charts: 
            if os.path.exists(c): os.remove(c)
