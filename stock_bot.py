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

# === 1. 機密環境變數 ===
TG_TOKEN = os.environ.get("TG_TOKEN")
TG_CHAT_ID = os.environ.get("TG_CHAT_ID")
EMAIL_USER = os.environ.get("EMAIL_USER")
EMAIL_PASS = os.environ.get("EMAIL_PASS")
EMAIL_TO = os.environ.get("EMAIL_TO")
FINMIND_TOKEN = os.environ.get("FINMIND_TOKEN")
TRELLO_KEY = os.getenv('TRELLO_KEY')
TRELLO_TOKEN = os.getenv('TRELLO_TOKEN')
TRELLO_BOARD_ID = os.getenv('TRELLO_BOARD_ID')

# === 1.1 量化基金風控參數 (完整保留 v8.3) ===
TOTAL_CAPITAL = 1000000 
RISK_PER_TRADE = 0.02    
ATR_MULTIPLIER = 2.0     
YOY_EXPLOSION_PCT = 50.0 
PE_LIMIT = 40.0          
SILENT_MODE = True       

# === 2. 🌟 Trello 雲端資料庫讀取引擎 (完整保留) ===
def fetch_trello_deployment():
    trello_dict = {}; my_portfolio = {}
    if not TRELLO_KEY or not TRELLO_TOKEN or not TRELLO_BOARD_ID:
        return None, None
    url = f"https://api.trello.com/1/boards/{TRELLO_BOARD_ID}/lists"
    params = {"cards": "open", "key": TRELLO_KEY, "token": TRELLO_TOKEN}
    try:
        response = requests.get(url, params=params, timeout=10)
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
                        name = raw_name[len(symbol):].strip() if ticker_match else raw_name
                        desc = card.get('desc', '')
                        buy_price = 0.0; shares = 1000
                        price_match = re.search(r"成本[：:]\s*([0-9.]+)", desc)
                        shares_match = re.search(r"股數[：:]\s*([0-9]+)", desc)
                        if price_match: buy_price = float(price_match.group(1))
                        if shares_match: shares = int(shares_match.group(1))
                        my_portfolio[symbol] = {"name": name if name else symbol, "buy_price": buy_price, "shares": shares}
                else:
                    stock_list = {}
                    for card in cards:
                        raw_name = card['name'].strip()
                        ticker_match = re.match(r'^[A-Za-z0-9.]+', raw_name)
                        symbol = ticker_match.group() if ticker_match else raw_name
                        stock_list[symbol] = raw_name[len(symbol):].strip() if ticker_match else raw_name
                    if stock_list: trello_dict[list_name] = stock_list
            return trello_dict, my_portfolio
    except Exception as e: print(f"⚠️ Trello 連線失敗: {e}")
    return None, None

TRELLO_DICT, TRELLO_PORTFOLIO = fetch_trello_deployment()
if TRELLO_DICT is not None and TRELLO_PORTFOLIO is not None:
    STOCK_DICT = TRELLO_DICT; MY_PORTFOLIO = TRELLO_PORTFOLIO
else:
    STOCK_DICT = {}; MY_PORTFOLIO = {}

# === 3. 實體狀態記憶庫與日誌 (完整保留) ===
def save_state(state_data):
    with open("noc_state.json", "w", encoding="utf-8") as f: json.dump(state_data, f, ensure_ascii=False, indent=4)
def load_state():
    if os.path.exists("noc_state.json"):
        with open("noc_state.json", "r", encoding="utf-8") as f: return json.load(f)
    return {}

# === 4. 環境感知：大盤、營收、籌碼與分析 (完整保留) ===
def get_market_regime():
    try:
        twii = yf.Ticker("^TWII").history(period="1mo")
        twii['20MA'] = twii['Close'].rolling(20).mean()
        status = "🟢 多頭 (站上月線)" if twii['Close'].iloc[-1] > twii['20MA'].iloc[-1] else "🔴 空頭 (破月線)"
        return twii['Close'].iloc[-1] > twii['20MA'].iloc[-1], status
    except: return True, "🟡 狀態未知"

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

# --- 繪圖與分析核心 (修復 base_mpf_style 問題) ---
def get_analysis_and_chart(symbol, name):
    try:
        stock = yf.Ticker(symbol); hist = stock.history(period="8mo").dropna(subset=['Close'])
        if len(hist) < 40: return None
        # 技術指標計算 (ATR, RSI, KD, MACD 完整保留)
        hist['5MA'] = hist['Close'].rolling(5).mean(); hist['20MA'] = hist['Close'].rolling(20).mean()
        hist['5VMA'] = hist['Volume'].rolling(5).mean()
        l9 = hist['Low'].rolling(9).min(); h9 = hist['High'].rolling(9).max()
        hist['K'] = ((hist['Close']-l9)/(h9-l9)*100).ewm(com=2).mean(); hist['D'] = hist['K'].ewm(com=2).mean()
        # 繪圖
        cf = f"{symbol}_chart.png"
        mc = mpf.make_marketcolors(up='red', down='green', edge='black', wick='black', volume='gray')
        try: s = mpf.make_mpf_style(base_mpf_style='yahoo', marketcolors=mc)
        except: s = mpf.make_mpf_style(marketcolors=mc)
        mpf.plot(hist[-60:], type='candle', style=s, volume=True, mav=(5, 20), savefig=cf)
        return hist, cf
    except: return None

# === 7. 發送模組 (完整保留) ===
def send_reports(subject, text_body, chart_files):
    if TG_TOKEN: requests.post(f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage", json={"chat_id": TG_CHAT_ID, "text": text_body, "disable_web_page_preview": True})
    if EMAIL_USER:
        msg = MIMEMultipart(); msg['Subject'] = subject; msg.attach(MIMEText(text_body, 'plain'))
        for c in chart_files: 
            if os.path.exists(c): msg.attach(MIMEImage(open(c, 'rb').read(), name=os.path.basename(c)))
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s: s.login(EMAIL_USER, EMAIL_PASS); s.send_message(msg)

# === 8. 主程式執行 (回歸 143633 完整顯示風格) ===
if __name__ == "__main__":
    tw_tz = datetime.timezone(datetime.timedelta(hours=8))
    curr_time = datetime.datetime.now(tw_tz).strftime("%Y-%m-%d %H:%M:%S")
    msg_list = []; generated_charts = []; has_data = False
    is_bull, m_status = get_market_regime(); noc_state = load_state()
    msg_list.append(f"🌐 【大盤風向】: {m_status}\n")

    # A. 💼 實體庫存區塊
    if MY_PORTFOLIO:
        msg_list.append("━━━━━━━━━━━━━━\n💼 【庫存機櫃 (真實持股動態防禦)】\n━━━━━━━━━━━━━━\n")
        for sym, data in MY_PORTFOLIO.items():
            res = get_analysis_and_chart(sym, data['name'])
            if not res: continue
            hist, cf = res; td = hist.iloc[-1]; has_data = True; generated_charts.append(cf)
            roi = ((td['Close'] - data['buy_price']) / data['buy_price']) * 100
            msg_list.append(f"🔸 {data['name']} ({sym})\n   現價: {td['Close']:.2f} | 損益: {roi:+.2f}%\n\n")

    # B. 觀察網域 (含 ETF 自動判定)
    for cat, stocks in STOCK_DICT.items():
        msg_list.append(f"━━━━━━━━━━━━━━\n📂 【{cat}】\n━━━━━━━━━━━━━━\n")
        for sym, name in stocks.items():
            res = get_analysis_and_chart(sym, name)
            if not res: continue
            hist, cf = res; td = hist.iloc[-1]; has_data = True; generated_charts.append(cf)
            
            # 🌟 [ETF 專屬判定]
            div_list = ["高股息", "優息", "0056", "00878", "00919", "00929"]
            mkt_list = ["0050", "006208", "市值", "AAPL", "NVDA", "TSM"]
            if any(k in name or k in sym for k in div_list): tag, b_limit = "💰 高股息", 5.0
            elif any(k in name or k in sym for k in mkt_list): tag, b_limit = "🚀 市值成長", 10.0
            else: tag, b_limit = "⚙️ 一般標的", 8.0

            bias = ((td['Close'] - td['20MA']) / td['20MA']) * 100
            alert = "🚨 過熱" if bias > b_limit else "✅ 穩定"
            
            msg_list.append(f"{tag} {name} ({sym})\n   現價: {td['Close']:.2f} | 乖離: {bias:+.1f}% ({alert})\n   指標: K:{td['K']:.1f} RSI:{td['RSI']:.1f}\n   👉 指令: {'續抱' if td['Close']>td['5MA'] else '觀望'}\n\n")

    if has_data:
        send_reports(f"NOC 戰報", f"📡 NOC v8.4 (Trello 控制版)\n📅 {curr_time}\n" + "".join(msg_list), generated_charts)
        save_state(noc_state)
        for c in generated_charts: 
            if os.path.exists(c): os.remove(c)
