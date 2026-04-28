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

# === 1. 環境變數與風控參數 (完整保留 v8.3) ===
TG_TOKEN = os.environ.get("TG_TOKEN")
TG_CHAT_ID = os.environ.get("TG_CHAT_ID")
EMAIL_USER = os.environ.get("EMAIL_USER")
EMAIL_PASS = os.environ.get("EMAIL_PASS")
EMAIL_TO = os.environ.get("EMAIL_TO")
FINMIND_TOKEN = os.environ.get("FINMIND_TOKEN")
TRELLO_KEY = os.getenv('TRELLO_KEY')
TRELLO_TOKEN = os.getenv('TRELLO_TOKEN')
TRELLO_BOARD_ID = os.getenv('TRELLO_BOARD_ID')

TOTAL_CAPITAL = 1000000 
RISK_PER_TRADE = 0.02    
ATR_MULTIPLIER = 2.0     
YOY_EXPLOSION_PCT = 50.0 
PE_LIMIT = 40.0          
SILENT_MODE = True       

# === 2. Trello 讀取與發報函式 (✅ 確保發報函式存在) ===
def fetch_trello_deployment():
    trello_dict = {}; my_portfolio = {}
    if not TRELLO_KEY: return None, None
    url = f"https://api.trello.com/1/boards/{TRELLO_BOARD_ID}/lists"
    params = {"cards": "open", "key": TRELLO_KEY, "token": TRELLO_TOKEN}
    try:
        r = requests.get(url, params=params, timeout=15)
        lists = r.json()
        for lst in lists:
            list_name = lst['name']
            cards = lst.get('cards', [])
            if "庫存" in list_name:
                for c in cards:
                    raw = c['name'].strip()
                    m = re.match(r'^[A-Za-z0-9.]+', raw)
                    sym = m.group() if m else raw
                    name = raw[len(sym):].strip()
                    desc = c.get('desc', '')
                    p = re.search(r"成本[：:]\s*([0-9.]+)", desc)
                    s = re.search(r"股數[：:]\s*([0-9]+)", desc)
                    my_portfolio[sym] = {"name": name if name else sym, "buy_price": float(p.group(1)) if p else 0, "shares": int(s.group(1)) if s else 1000}
            else:
                stock_list = {re.match(r'^[A-Za-z0-9.]+', c['name']).group(): c['name'][len(re.match(r'^[A-Za-z0-9.]+', c['name']).group()):].strip() for c in cards}
                if stock_list: trello_dict[list_name] = stock_list
        return trello_dict, my_portfolio
    except: return None, None

def send_reports(subject, text_body, chart_files):
    if TG_TOKEN: requests.post(f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage", json={"chat_id": TG_CHAT_ID, "text": text_body, "disable_web_page_preview": True})
    if EMAIL_USER:
        msg = MIMEMultipart(); msg['Subject'] = subject; msg.attach(MIMEText(text_body, 'plain'))
        for c in chart_files: 
            if os.path.exists(c): 
                with open(c, 'rb') as f: msg.attach(MIMEImage(f.read(), name=os.path.basename(c)))
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s: s.login(EMAIL_USER, EMAIL_PASS); s.send_message(msg)

# === 3. 分析核心 (✅ 包含 KD/RSI/籌碼/預判) ===
def get_analysis_and_chart(symbol, name):
    try:
        stock = yf.Ticker(symbol); hist = stock.history(period="8mo").dropna()
        if len(hist) < 40: return None
        # MA/ATR/RSI 計算
        hist['20MA'] = hist['Close'].rolling(20).mean(); hist['5MA'] = hist['Close'].rolling(5).mean()
        hist['ATR'] = (hist['High']-hist['Low']).rolling(14).mean()
        delta = hist['Close'].diff(); g = delta.clip(lower=0); l = -delta.clip(upper=0)
        hist['RSI'] = 100 - (100 / (1 + (g.ewm(com=13).mean() / l.ewm(com=13).mean())))
        # KD 計算
        l9 = hist['Low'].rolling(9).min(); h9 = hist['High'].rolling(9).max()
        hist['K'] = ((hist['Close']-l9)/(h9-l9)*100).ewm(com=2).mean()
        hist['D'] = hist['K'].ewm(com=2).mean()
        # 繪圖修復
        cf = f"{symbol}_chart.png"
        mc = mpf.make_marketcolors(up='red', down='green', edge='black', wick='black', volume='gray')
        s = mpf.make_mpf_style(base_mpf_style='yahoo', marketcolors=mc)
        mpf.plot(hist[-60:], type='candle', style=s, volume=True, mav=(5, 20), savefig=cf)
        return hist, cf
    except: return None

# === 4. 主程式 (✅ 回歸 143633.JPG 並顯示所有指標) ===
if __name__ == "__main__":
    tw_tz = datetime.timezone(datetime.timedelta(hours=8))
    curr_time = datetime.datetime.now(tw_tz).strftime("%Y-%m-%d %H:%M:%S")
    msg_list = []; gen_charts = []; has_data = False
    
    TRELLO_DICT, TRELLO_PORTFOLIO = fetch_trello_deployment()
    STOCK_DICT = TRELLO_DICT if TRELLO_DICT else {}
    MY_PORTFOLIO = TRELLO_PORTFOLIO if TRELLO_PORTFOLIO else {}

    # A. 庫存區
    if MY_PORTFOLIO:
        msg_list.append("━━━━━━━━━━━━━━\n💼 【庫存機櫃 (實體持股)】\n━━━━━━━━━━━━━━\n")
        for sym, data in MY_PORTFOLIO.items():
            res = get_analysis_and_chart(sym, data['name'])
            if not res: continue
            hist, chart = res; td = hist.iloc[-1]; has_data = True; gen_charts.append(chart)
            roi = ((td['Close'] - data['buy_price']) / data['buy_price']) * 100
            msg_list.append(f"⚙️ {data['name']} ({sym})\n   損益: {roi:+.2f}% | RSI: {td['RSI']:.1f}\n   👉 指令: {'🔥獲利' if roi>0 else '🩸破網'}\n\n")

    # B. 觀察區
    for cat, stocks in STOCK_DICT.items():
        msg_list.append(f"━━━━━━━━━━━━━━\n📂 【{cat}】\n━━━━━━━━━━━━━━\n")
        for sym, name in stocks.items():
            res = get_analysis_and_chart(sym, name)
            if not res: continue
            hist, chart = res; td = hist.iloc[-1]; has_data = True; gen_charts.append(chart)
            # 完整顯示：籌碼、預判、KD、RSI
            kd_str = f"K:{td['K']:.1f} D:{td['D']:.1f}"
            act = "🚀進場" if td['Close'] > td['5MA'] and td['RSI'] < 70 else "✅續抱"
            msg_list.append(f"⚙️ {name} ({sym})\n   現價: {td['Close']:.2f} | {kd_str}\n   指標: RSI:{td['RSI']:.1f} | 預判:無特殊\n   👉 指令: {act}\n\n")

    if has_data:
        send_reports(f"NOC 戰報", f"📡 NOC v8.5\n📅 {curr_time}\n" + "".join(msg_list), gen_charts)
        for c in gen_charts: 
            if os.path.exists(c): os.remove(c)
