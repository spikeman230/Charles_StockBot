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

# === 1.1 風控參數 ===
TOTAL_CAPITAL = 1000000 
RISK_PER_TRADE = 0.02    
ATR_MULTIPLIER = 2.0     
YOY_EXPLOSION_PCT = 50.0 
PE_LIMIT = 40.0          
SILENT_MODE = True       

# === 2. Trello 資料讀取引擎 ===
def fetch_trello_deployment():
    trello_dict = {}; my_portfolio = {}
    if not TRELLO_KEY or not TRELLO_TOKEN or not TRELLO_BOARD_ID: return None, None
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
                        name = raw_name[len(symbol):].strip() if ticker_match else raw_name
                        stock_list[symbol] = name if name else symbol
                    if stock_list: trello_dict[list_name] = stock_list
            return trello_dict, my_portfolio
    except: pass
    return None, None

TRELLO_DICT, TRELLO_PORTFOLIO = fetch_trello_deployment()
STOCK_DICT = TRELLO_DICT if TRELLO_DICT else {}
MY_PORTFOLIO = TRELLO_PORTFOLIO if TRELLO_PORTFOLIO else {}

# === 3. 狀態與日誌 ===
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

# === 4. 環境感知與分析 ===
def get_market_regime():
    try:
        twii = yf.Ticker("^TWII").history(period="1mo")
        twii['20MA'] = twii['Close'].rolling(20).mean()
        status = "🟢 多頭" if twii['Close'].iloc[-1] > twii['20MA'].iloc[-1] else "🔴 空頭"
        return twii['Close'].iloc[-1] > twii['20MA'].iloc[-1], status
    except: return True, "🟡 未知"

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
        return info.get('trailingPE', "N/A")
    except: return "N/A"

def get_analysis_and_chart(symbol, name):
    try:
        stock = yf.Ticker(symbol); hist = stock.history(period="8mo").dropna()
        if len(hist) < 40: return None
        # 簡易指標計算
        hist['5MA'] = hist['Close'].rolling(5).mean(); hist['20MA'] = hist['Close'].rolling(20).mean()
        hist['5VMA'] = hist['Volume'].rolling(5).mean(); hist['ATR'] = (hist['High']-hist['Low']).rolling(14).mean()
        # RSI
        delta = hist['Close'].diff(); gain = delta.clip(lower=0); loss = -delta.clip(upper=0)
        hist['RSI'] = 100 - (100 / (1 + (gain.ewm(com=13).mean() / loss.ewm(com=13).mean())))
        # Sniper Signal
        hist['20_High'] = hist['High'].rolling(20).max().shift(1)
        hist['Is_Breakout'] = (hist['Close'] > hist['5MA']) & (hist['Volume'] > hist['5VMA']*1.2)
        
        chart_file = f"{symbol}_chart.png"
        mc = mpf.make_marketcolors(up='red', down='green', edge='black', wick='black', volume='gray')
        s = mpf.make_mpf_style(base_mpf_style='yahoo', marketcolors=mc)
        mpf.plot(hist[-60:], type='candle', style=s, volume=True, mav=(5, 20), savefig=chart_file)
        return hist, chart_file
    except: return None

def send_reports(subject, text_body, chart_files):
    if TG_TOKEN and TG_CHAT_ID:
        requests.post(f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage", json={"chat_id": TG_CHAT_ID, "text": text_body, "disable_web_page_preview": True})

# === 8. 主程式執行 (Trello 完全對齊版) ===
if __name__ == "__main__":
    tw_tz = datetime.timezone(datetime.timedelta(hours=8))
    curr_time = datetime.datetime.now(tw_tz).strftime("%Y-%m-%d %H:%M:%S")
    msg_list = []; gen_charts = []; has_data = False
    is_bull, m_status = get_market_regime()
    noc_state = load_state()
    
    msg_list.append(f"🌐 【大盤風向】: {m_status}\n")

    # A. 💼 實體庫存 (與 Trello 庫存清單同步)
    if MY_PORTFOLIO:
        msg_list.append("■■■ 💼 庫存機櫃 (動態防禦) ■■■\n")
        for sym, data in MY_PORTFOLIO.items():
            res = get_analysis_and_chart(sym, data['name'])
            if not res: continue
            hist, chart = res; td = hist.iloc[-1]; has_data = True; gen_charts.append(chart)
            
            roi = ((td['Close'] - data['buy_price']) / data['buy_price']) * 100
            stop_line = td['Close'] - (td['ATR'] * ATR_MULTIPLIER)
            
            pnl_tag = "🚀獲利巡航" if roi > 0 else "🟡暫時浮虧"
            if td['Close'] < stop_line: pnl_tag = "🚨拔線離場"
            
            msg_list.append(f"🔹 {data['name']} ({sym}): {td['Close']:.2f}\n└ 損益:{roi:+.1f}% | {pnl_tag} (防線:{stop_line:.1f})\n")

    # B. 觀察網域 (與 Trello 其他清單同步)
    for cat, stocks in STOCK_DICT.items():
        msg_list.append(f"■■■ {cat} ■■■\n")
        for sym, name in stocks.items():
            res = get_analysis_and_chart(sym, name)
            if not res: continue
            hist, chart = res; td = hist.iloc[-1]; has_data = True; gen_charts.append(chart)
            
            yoy = get_revenue_yoy(sym); pe = get_pe_ratio(sym)
            yoy_str = f"{yoy:+.0f}%" if isinstance(yoy, float) else "N/A"
            
            # 指令判定
            bias = ((td['Close'] - td['20MA']) / td['20MA']) * 100
            act = "🚀進場" if td['Is_Breakout'] and bias < 10 else "✅續抱"
            if td['RSI'] > 75: act = "⚠️過熱"
            
            msg_list.append(f"🔸 {name} ({sym}): {td['Close']:.2f}\n└ {act} | RSI:{td['RSI']:.0f} | YoY:{yoy_str}\n")

    if has_data:
        final_text = f"📡 NOC 終極戰情室 v8.4 (Trello 同步版)\n📅 {curr_time}\n" + "".join(msg_list)
        send_reports(f"NOC 戰報", final_text, gen_charts)
        save_state(noc_state)
        for c in gen_charts: 
            if os.path.exists(c): os.remove(c)
