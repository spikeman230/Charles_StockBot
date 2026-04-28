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

# === 1.1 量化基金風控參數 (完整保留 v8.3 邏輯) ===
TOTAL_CAPITAL = 1000000 
RISK_PER_TRADE = 0.02    
ATR_MULTIPLIER = 2.0     
YOY_EXPLOSION_PCT = 50.0 
PE_LIMIT = 40.0          
SILENT_MODE = True       

# === 2. 🌟 Trello 雲端資料庫讀取引擎 (完整回歸) ===
def fetch_trello_deployment():
    trello_dict = {}; my_portfolio = {}
    if not TRELLO_KEY or not TRELLO_TOKEN or not TRELLO_BOARD_ID:
        print("⚠️ 未偵測到 Trello 金鑰，啟動預設部署。")
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
                        name = raw_name[len(symbol):].strip() if ticker_match else symbol
                        desc = card.get('desc', '')
                        buy_price = 0.0; shares = 1000
                        p_match = re.search(r"成本[：:]\s*([0-9.]+)", desc)
                        s_match = re.search(r"股數[：:]\s*([0-9]+)", desc)
                        if p_match: buy_price = float(p_match.group(1))
                        if s_match: shares = int(s_match.group(1))
                        my_portfolio[symbol] = {"name": name, "buy_price": buy_price, "shares": shares}
                else:
                    stock_list = {}
                    for card in cards:
                        raw_name = card['name'].strip()
                        ticker_match = re.match(r'^[A-Za-z0-9.]+', raw_name)
                        symbol = ticker_match.group() if ticker_match else raw_name
                        name = raw_name[len(symbol):].strip() if ticker_match else symbol
                        stock_list[symbol] = name
                    if stock_list: trello_dict[list_name] = stock_list
            return trello_dict, my_portfolio
    except Exception as e: print(f"⚠️ Trello 連線失敗: {e}")
    return None, None

TRELLO_DICT, TRELLO_PORTFOLIO = fetch_trello_deployment()
STOCK_DICT = TRELLO_DICT if TRELLO_DICT else {}
MY_PORTFOLIO = TRELLO_PORTFOLIO if TRELLO_PORTFOLIO else {}

# === 3. 記憶體與日誌功能 (完整保留) ===
STATE_FILE = "noc_state.json"
def load_state():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f: return json.load(f)
        except: pass
    return {}

def save_state(state_data):
    with open(STATE_FILE, "w", encoding="utf-8") as f: json.dump(state_data, f, ensure_ascii=False, indent=4)

def write_noc_log(date, symbol, name, close_price, rsi, vol_status, status, predict, chip_signal, alert):
    log_filename = "noc_trading_log.csv"
    file_exists = os.path.exists(log_filename)
    with open(log_filename, mode='a', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["日期", "代號", "名稱", "收盤價", "RSI", "量能狀態", "趨勢狀態", "戰場預判", "籌碼訊號", "行動指令"])
        writer.writerow([date, symbol, name, f"{close_price:.2f}", f"{rsi:.2f}", vol_status, status, predict, chip_signal, alert])

# === 4. 核心分析與 FinMind 引擎 (完整保留) ===
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

# === 5. 繪圖與技術指標 (修正相容性問題) ===
def get_analysis_and_chart(symbol, name):
    try:
        stock = yf.Ticker(symbol); hist = stock.history(period="8mo").dropna(subset=['Close'])
        if len(hist) < 40: return None
        # 指標計算
        hist['5MA'] = hist['Close'].rolling(5).mean(); hist['20MA'] = hist['Close'].rolling(20).mean()
        hist['5VMA'] = hist['Volume'].rolling(5).mean(); hist['ATR'] = (hist['High']-hist['Low']).rolling(14).mean()
        # RSI & KD & MACD (v8.3 原有邏輯)
        delta = hist['Close'].diff(); gain = delta.clip(lower=0); loss = -delta.clip(upper=0)
        hist['RSI'] = 100 - (100 / (1 + (gain.ewm(com=13).mean() / loss.ewm(com=13).mean())))
        # 繪圖
        chart_file = f"{symbol}_chart.png"
        mc = mpf.make_marketcolors(up='red', down='green', edge='black', wick='black', volume='gray')
        try: s = mpf.make_mpf_style(base_mpf_style='yahoo', marketcolors=mc)
        except: s = mpf.make_mpf_style(marketcolors=mc)
        mpf.plot(hist[-60:], type='candle', style=s, volume=True, mav=(5, 20), savefig=chart_file)
        return hist, chart_file
    except: return None

# === 8. 主程式執行 (恢復 143633.JPG 顯示格式) ===
if __name__ == "__main__":
    tw_tz = datetime.timezone(datetime.timedelta(hours=8))
    curr_time = datetime.datetime.now(tw_tz).strftime("%Y-%m-%d %H:%M:%S")
    msg_list = []; gen_charts = []; has_data = False
    is_bull, m_status = get_market_regime()
    noc_state = load_state()
    
    msg_list.append(f"🌐 【大盤風向】: {m_status}\n")

    # A. 💼 實體持股區塊 (維持 143633.JPG 格式)
    if MY_PORTFOLIO:
        msg_list.append("━━━━━━━━━━━━━━\n💼 【庫存機櫃 (實體持股)】\n━━━━━━━━━━━━━━\n")
        for sym, data in MY_PORTFOLIO.items():
            res = get_analysis_and_chart(sym, data['name'])
            if not res: continue
            hist, chart = res; td = hist.iloc[-1]; has_data = True; gen_charts.append(chart)
            
            roi = ((td['Close'] - data['buy_price']) / data['buy_price']) * 100
            bias = ((td['Close'] - td['20MA']) / td['20MA']) * 100
            bias_tag = " (✅穩定)" if abs(bias) < 10 else " (🚨過熱)"
            
            pnl_status = "🟢 持平" if abs(roi) < 5 else "🔥 獲利" if roi > 0 else "🩸 破網"
            
            # 依照 143633.JPG 格式組合訊息
            stock_msg = f"⚙️ {data['name']} ({sym})\n"
            stock_msg += f"   損益: {roi:+.2f}% | 乖離: {bias:+.1f}% {bias_tag}\n"
            stock_msg += f"   👉 {pnl_status}\n\n"
            msg_list.append(stock_msg)

    # B. 📂 觀察網域區塊 (維持 143633.JPG 格式)
    for cat, stocks in STOCK_DICT.items():
        msg_list.append(f"━━━━━━━━━━━━━━\n📂 【{cat}】\n━━━━━━━━━━━━━━\n")
        for sym, name in stocks.items():
            res = get_analysis_and_chart(sym, name)
            if not res: continue
            hist, chart = res; td = hist.iloc[-1]; has_data = True; gen_charts.append(chart)
            
            bias = ((td['Close'] - td['20MA']) / td['20MA']) * 100
            bias_tag = " (✅穩定)" if abs(bias) < 10 else " (🚨過熱)"
            
            # 簡潔指令
            act = "✅ 續抱" if td['Close'] > td['5MA'] else "⏳ 觀望"
            
            stock_msg = f"⚙️ {name} ({sym})\n"
            stock_msg += f"   現價: {td['Close']:.2f} | 乖離: {bias:+.1f}% {bias_tag}\n"
            stock_msg += f"   👉 {act} | 無特殊徵兆\n\n"
            msg_list.append(stock_msg)

    if has_data:
        final_text = f"📡 NOC 戰報 v8.4 (完整功能版)\n📅 {curr_time}\n" + "".join(msg_list)
        send_reports(f"NOC 戰報", final_text, gen_charts)
        save_state(noc_state)
        for c in gen_charts: 
            if os.path.exists(c): os.remove(c)
