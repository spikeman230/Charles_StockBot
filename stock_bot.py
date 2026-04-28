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

# 🌟 Trello 雲端控制台金鑰
TRELLO_KEY = os.getenv('TRELLO_KEY')
TRELLO_TOKEN = os.getenv('TRELLO_TOKEN')
TRELLO_BOARD_ID = os.getenv('TRELLO_BOARD_ID')

# === 1.1 量化基金風控參數 (v8.4 雲端路由版) ===
TOTAL_CAPITAL = 1000000  # 預設總資金 100 萬台幣
RISK_PER_TRADE = 0.02    # 單筆風險 2%
ATR_MULTIPLIER = 2.0     # 2倍 ATR 動態停損
YOY_EXPLOSION_PCT = 50.0 # 業績大爆發閥值 (50%)
PE_LIMIT = 40.0          # 本益比上限 (超過 40 倍視為過貴)
SILENT_MODE = True       # 開啟靜默模式 (只發送有特殊訊號的股票，減少雜訊)
FRICTION_COST = 0.6      # 🌟 網管摩擦成本 (買賣手續費+稅金約 0.6%)

# === 2. Trello 雲端資料庫讀取引擎 ===
def fetch_trello_deployment():
    trello_dict = {}
    my_portfolio = {}
    
    if not TRELLO_KEY or not TRELLO_TOKEN or not TRELLO_BOARD_ID:
        print("⚠️ 未偵測到 Trello 金鑰，將啟用『機房預設』的靜態兵力部署。")
        return None, None

    print("🌐 正在連線 Trello 戰情看板，讀取最新部署...")
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
                        name = name if name else symbol
                        desc = card.get('desc', '')
                        
                        buy_price = 0.0
                        shares = 1000
                        
                        price_match = re.search(r"成本[：:]\s*([0-9.]+)", desc)
                        shares_match = re.search(r"股數[：:]\s*([0-9]+)", desc)
                        
                        if price_match: buy_price = float(price_match.group(1))
                        if shares_match: shares = int(shares_match.group(1))
                        
                        my_portfolio[symbol] = {"name": name, "buy_price": buy_price, "shares": shares}
                else:
                    stock_list = {}
                    for card in cards:
                        raw_name = card['name'].strip()
                        ticker_match = re.match(r'^[A-Za-z0-9.]+', raw_name)
                        symbol = ticker_match.group() if ticker_match else raw_name
                        name = raw_name[len(symbol):].strip() if ticker_match else raw_name
                        name = name if name else symbol
                        stock_list[symbol] = name
                        
                    if stock_list:
                        trello_dict[list_name] = stock_list
            print("✅ 成功從 Trello 載入最新戰略部署！")
            return trello_dict, my_portfolio
        else:
            print(f"⚠️ Trello API 錯誤: {response.text}")
    except Exception as e:
        print(f"⚠️ Trello 連線失敗: {e}")
        
    return None, None

TRELLO_DICT, TRELLO_PORTFOLIO = fetch_trello_deployment()

# === 2.1 兵力掛載邏輯 ===
if TRELLO_DICT is not None and TRELLO_PORTFOLIO is not None:
    STOCK_DICT = TRELLO_DICT
    MY_PORTFOLIO = TRELLO_PORTFOLIO
else:
    print("⚠️ Trello 斷線，啟動緊急預設名單...")
    STOCK_DICT = {}
    MY_PORTFOLIO = {}

RADAR_FILE = "radar_targets.json"
if os.path.exists(RADAR_FILE):
    try:
        with open(RADAR_FILE, "r", encoding="utf-8") as f:
            radar_stocks = json.load(f)
            if radar_stocks:
                STOCK_DICT["🎯 雷達鎖定 (新進火種區)"] = radar_stocks
    except Exception as e:
        print(f"⚠️ 游擊隊雷達名單讀取失敗: {e}")

LIGHTNING_FILE = "lightning_targets.json"
if os.path.exists(LIGHTNING_FILE):
    try:
        with open(LIGHTNING_FILE, "r", encoding="utf-8") as f:
            lightning_stocks = json.load(f)
            if lightning_stocks:
                STOCK_DICT["⚡ 雷達鎖定 (短線飆股區)"] = lightning_stocks
    except Exception as e:
        print(f"⚠️ 閃電突擊名單讀取失敗: {e}")

# === 3. 實體狀態記憶庫與日誌 ===
STATE_FILE = "noc_state.json"

def load_state():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f: 
                return json.load(f)
        except Exception as e: 
            print(f"⚠️ 讀取記憶體失敗: {e}")
    return {}

def save_state(state_data):
    try:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(state_data, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"⚠️ 寫入記憶體失敗: {e}")

def write_noc_log(date, symbol, name, close_price, rsi, vol_status, status, predict, chip_signal, alert):
    log_filename = "noc_trading_log.csv"
    file_exists = os.path.exists(log_filename)
    try:
        with open(log_filename, mode='a', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow(["日期", "代號", "名稱", "收盤價", "RSI", "量能狀態", "趨勢狀態", "戰場預判", "籌碼訊號", "行動指令"])
            writer.writerow([date, symbol, name, f"{close_price:.2f}", f"{rsi:.2f}", vol_status, status, predict, chip_signal, alert])
    except Exception as e:
        print(f"⚠️ 日誌寫入失敗: {e}")

# === 4. 環境感知：大盤、營收與估值 ===
def get_market_regime():
    try:
        twii = yf.Ticker("^TWII").history(period="1mo")
        twii['20MA'] = twii['Close'].rolling(20).mean()
        if twii['Close'].iloc[-1] > twii['20MA'].iloc[-1]: 
            return True, f"🟢 多頭格局 (站上月線)"
        else: 
            return False, f"🔴 空頭警戒 (跌破月線)"
    except Exception as e: 
        print(f"⚠️ 大盤狀態讀取失敗: {e}")
        return True, "🟡 大盤狀態未知"

def get_revenue_yoy(symbol):
    if not FINMIND_TOKEN: return "N/A"
    match = re.search(r'\d+', symbol)
    if not match: return "N/A"
    fm_symbol = match.group()
    
    try:
        url = "https://api.finmindtrade.com/api/v4/data"
        params = {
            "dataset": "TaiwanStockMonthRevenue", 
            "data_id": fm_symbol, 
            "start_date": (datetime.datetime.now() - datetime.timedelta(days=400)).strftime("%Y-%m-%d"), 
            "token": FINMIND_TOKEN
        }
        r = requests.get(url, params=params, timeout=10)
        data = r.json()
        if data.get("msg") == "success" and len(data.get("data", [])) > 0:
            df = pd.DataFrame(data["data"])
            if 'revenue' in df.columns and 'revenue_month' in df.columns and 'revenue_year' in df.columns:
                latest_record = df.iloc[-1]
                latest_rev = latest_record['revenue']
                target_month = latest_record['revenue_month']
                target_year_last = latest_record['revenue_year'] - 1
                
                last_year_record = df[(df['revenue_year'] == target_year_last) & (df['revenue_month'] == target_month)]
                if not last_year_record.empty:
                    last_year_rev = last_year_record.iloc[-1]['revenue']
                    if last_year_rev > 0: 
                        return float(((latest_rev - last_year_rev) / last_year_rev) * 100)
    except Exception as e: 
        print(f"⚠️ [{symbol}] 營收處理錯誤: {e}")
    return "N/A"

def get_pe_ratio(symbol):
    try:
        info = yf.Ticker(symbol).info
        return info.get('trailingPE', info.get('forwardPE', "N/A"))
    except:
        return "N/A"

# === 5. FinMind 籌碼分析 ===
def get_finmind_chip_data(symbol, start_date_str):
    if not FINMIND_TOKEN: return pd.DataFrame()
    match = re.search(r'\d+', symbol)
    if not match: return pd.DataFrame()
    fm_symbol = match.group()
        
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
    except Exception as e: 
        print(f"⚠️ [{symbol}] 籌碼讀取錯誤: {e}")
    return pd.DataFrame()

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
        choices = ["🤝 土洋齊買", "🏦 投信作帳(連買)", "📈 法人偏多"]
        hist['Chip_Status'] = np.select(conditions, choices, default="➖ 中性/偏空")
    return hist

# === 5.5 特種戰略分析引擎 (AI 提示模組) ===
def get_strategy_tips(symbol, current_price, k_value, ma5, ma20):
    if symbol == "9933.TW":
        if current_price > ma5 and k_value < 30: return "🔥【NOC 訊號】中鼎疑似止跌！符合第一梯隊進場條件(30%)"
        else: return "⏳【NOC 監控】中鼎尚未止跌，請繼續等待 K<20 且站上 5MA。"
    if symbol == "6415.TW":
        if current_price <= 260 and current_price >= 240: return "💎【NOC 訊號】矽力進入支撐區，適合長線波段建倉(破230停損)。"
        else: return "🦅【NOC 監控】矽力距支撐位尚有空間，耐心等待回測。"
    if symbol == "2303.TW":
        if current_price > ma5: return f"🚀【NOC 訊號】聯電強勢站穩 5MA，動能強勁，目標挑戰前高 79.7！"
        else: return f"⚠️【NOC 警訊】聯電短線跌破 5MA，4天閃電戰動能轉弱，注意撤退防守。"
    return ""

# === 6. 核心分析引擎 (🌟 新增 is_etf 參數與季線) ===
def get_analysis_and_chart(symbol, name, is_etf=False):
    try:
        stock = yf.Ticker(symbol)
        hist = stock.history(period="8mo")
        hist = hist.dropna(subset=['Close'])
        if len(hist) < 60: return None  # 為了計算 60MA，長度改為 60
            
        hist['Date_Key'] = hist.index.date
        
        # 🛑 網管路由分流：非 ETF 才抓籌碼
        if not is_etf and FINMIND_TOKEN and (".TW" in symbol or ".TWO" in symbol):
            start_date_str = (datetime.datetime.now() - datetime.timedelta(days=200)).strftime("%Y-%m-%d")
            chip_df = get_finmind_chip_data(symbol, start_date_str)
            if not chip_df.empty:
                hist = hist.merge(chip_df, left_on='Date_Key', right_index=True, how='left')
                hist = hist.fillna({'Foreign_Inv': 0, 'Trust_Inv': 0, 'Dealer_Inv': 0})
            hist = calculate_chip_signals(hist)
        else:
            hist['Chip_Status'] = "ETF無籌碼分析" if is_etf else "無籌碼資料"
        
        hist['5MA'] = hist['Close'].rolling(window=5).mean()
        hist['20MA'] = hist['Close'].rolling(window=20).mean()
        hist['60MA'] = hist['Close'].rolling(window=60).mean()  # 🌟 季線
        hist['5VMA'] = hist['Volume'].rolling(window=5).mean()

        # 🌟 ETF 專用乖離率
        if is_etf:
            hist['Bias_60'] = ((hist['Close'] - hist['60MA']) / hist['60MA']) * 100
        else:
            hist['Bias_60'] = 0

        low_9 = hist['Low'].rolling(window=9).min()
        high_9 = hist['High'].rolling(window=9).max()
        rsv = ((hist['Close'] - low_9) / (high_9 - low_9)) * 100
        hist['K'] = rsv.ewm(com=2, adjust=False).mean()
        hist['D'] = hist['K'].ewm(com=2, adjust=False).mean()
        
        delta = hist['Close'].diff()
        gain = delta.clip(lower=0)
        loss = -1 * delta.clip(upper=0)
        ema_gain = gain.ewm(com=13, adjust=False).mean()
        ema_loss = loss.ewm(com=13, adjust=False).mean()
        hist['RSI'] = 100 - (100 / (1 + (ema_gain / ema_loss)))
        hist['RSI'] = hist['RSI'].fillna(50)
        
        hist['H-L'] = hist['High'] - hist['Low']
        hist['H-PC'] = abs(hist['High'] - hist['Close'].shift(1))
        hist['L-PC'] = abs(hist['Low'] - hist['Close'].shift(1))
        hist['TR'] = hist[['H-L', 'H-PC', 'L-PC']].max(axis=1)
        hist['ATR'] = hist['TR'].rolling(window=14).mean()
        
        hist['MACD'] = hist['Close'].ewm(span=12, adjust=False).mean() - hist['Close'].ewm(span=26, adjust=False).mean()
        hist['Signal'] = hist['MACD'].ewm(span=9, adjust=False).mean()
        hist['MACD_Hist'] = hist['MACD'] - hist['Signal']
        hist['STD20'] = hist['Close'].rolling(window=20).std()
        hist['BB_Width'] = (4 * hist['STD20']) / hist['20MA']
        
        hist['Is_Bottoming'] = (
            (hist['Close'] < hist['5MA']) & 
            (hist['MACD_Hist'].shift(2) < hist['MACD_Hist'].shift(1)) & 
            (hist['MACD_Hist'].shift(1) < hist['MACD_Hist']) & 
            (hist['MACD_Hist'] < 0)
        ).astype(int)
        hist['Recent_Bottoming'] = hist['Is_Bottoming'].rolling(window=3).max().fillna(0).astype(bool)
        hist['Is_Breakout'] = (hist['Close'].shift(1) < hist['5MA'].shift(1)) & (hist['Close'] > hist['5MA']) & (hist['Volume'] > hist['5VMA'] * 1.2)
        hist['Sniper_Signal'] = hist['Recent_Bottoming'] & hist['Is_Breakout']
        hist['Sniper_Memory_5D'] = hist['Sniper_Signal'].rolling(window=5).max().fillna(0)
        
        hist['20_High'] = hist['High'].rolling(window=20).max().shift(1)
        hist['Body_Top'] = hist[['Open', 'Close']].max(axis=1)
        hist['Upper_Shadow'] = hist['High'] - hist['Body_Top']
        hist['K_Length'] = (hist['High'] - hist['Low']).replace(0, 0.001) 
        hist['Shadow_Ratio'] = hist['Upper_Shadow'] / hist['K_Length']
        
        chart_file = f"{symbol}_chart.png"
        try:
            mc = mpf.make_marketcolors(up='red', down='green', edge='black', wick='black', volume='gray')
            tw_style = mpf.make_mpf_style(base_mpf_style='yahoo', marketcolors=mc)
            mpf.plot(hist[-60:], type='candle', style=tw_style, volume=True, mav=(5, 20), title=f"Stock: {symbol}", savefig=chart_file)
        except Exception as e:
            print(f"[{symbol}] 自訂圖表繪製失敗，改用預設風格: {e}")
            mpf.plot(hist[-60:], type='candle', style='yahoo', volume=True, mav=(5, 20), title=f"Stock: {symbol}", savefig=chart_file)
            
        return hist, chart_file
    except Exception as e: 
        print(f"[{symbol}] 核心分析發生錯誤: {e}")
        return None

# === 7. 發送模組 ===
def send_reports(subject, text_body, chart_files):
    if TG_TOKEN and TG_CHAT_ID:
        try:
            max_length = 4000
            message_parts = [text_body[i:i + max_length] for i in range(0, len(text_body), max_length)]
            for part in message_parts:
                resp = requests.post(
                    f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage", 
                    json={"chat_id": TG_CHAT_ID, "text": part, "disable_web_page_preview": True},
                    timeout=10
                )
                if resp.status_code != 200: print(f"⚠️ Telegram API 拒絕發送: {resp.text}")
        except Exception as e: print(f"❌ Telegram 發送失敗: {e}")
    
    if EMAIL_USER and EMAIL_PASS and EMAIL_TO:
        try:
            msg = MIMEMultipart()
            msg['From'] = EMAIL_USER
            msg['To'] = EMAIL_TO
            msg['Subject'] = subject
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
                server.login(EMAIL_USER, EMAIL_PASS)
                server.send_message(msg)
            print("✅ 戰報發送成功！")
        except Exception as e: print(f"❌ Email 發送失敗: {e}")

# === 8. 主程式執行 ===
if __name__ == "__main__":
    tw_tz = datetime.timezone(datetime.timedelta(hours=8))
    curr_date = datetime.datetime.now(tw_tz).date()
    curr_time = datetime.datetime.now(tw_tz).strftime("%Y-%m-%d %H:%M:%S")
    msg_list = []
    generated_charts = []
    has_data = False
    
    print(f"[{curr_time}] NOC 終極融合版 (v8.4 雲端路由版) 啟動...")
    
    is_bull_market, market_msg = get_market_regime()
    noc_state = load_state()
    msg_list.append(f"🌐 【大盤風向】: {market_msg}\n")

    # === 🌟 處理真實持股 (路由分流與摩擦成本) ===
    if MY_PORTFOLIO:
        msg_list.append("━━━━━━━━━━━━━━\n💼 【庫存機櫃 (真實持股動態防禦)】\n━━━━━━━━━━━━━━\n")
        for sym, data in MY_PORTFOLIO.items():
            # 判斷是否為 ETF
            is_etf = sym.startswith("00") or "ETF" in data['name'].upper()
            
            res = get_analysis_and_chart(sym, data['name'], is_etf=is_etf)
            if not res: continue
                
            hist, chart_file = res
            td = hist.iloc[-1]
            has_data = True
            generated_charts.append(chart_file)
            
            curr_price = td['Close']
            buy_price = data['buy_price']
            shares = data['shares']
            # 計算淨損益 (扣除 0.6% 摩擦成本)
            roi_net = ((curr_price - buy_price) / buy_price) * 100 - FRICTION_COST
            
            portfolio_msg = f"🔸 {data['name']} ({sym})\n"
            portfolio_msg += f"   成本: {buy_price:.2f} | 股數: {shares} | 現價: {curr_price:.2f} | 淨損益: {roi_net:+.2f}%\n"

            # 🛤️ 路由 A：ETF 邏輯
            if is_etf:
                bias = td['Bias_60']
                portfolio_msg += f"   📊 季線乖離: {bias:+.2f}%\n"
                
                if bias > 10: pnl_alert = "⚠️【正乖離過大】短線過熱，不宜單筆追高"
                elif bias < -10: pnl_alert = "💰【負乖離超跌】進入甜甜價，適合放大定期定額"
                elif roi_net <= -15: pnl_alert = "🩸【破網警戒】ETF 罕見大跌，檢視大盤系統風險"
                else: pnl_alert = "✅【紀律扣款】常態區間，維持現狀"
                
            # 🛤️ 路由 B：個股高波動邏輯
            else:
                atr = td['ATR']
                stop_distance = atr * ATR_MULTIPLIER
                sym_state = noc_state.get(sym, {"status": "NONE"})
                
                if sym_state["status"] != "REAL_HOLD":
                    noc_state[sym] = {"status": "REAL_HOLD", "entry": buy_price, "trailing_stop": curr_price - stop_distance}
                    sym_state = noc_state[sym]
                    
                final_stop = max(sym_state["trailing_stop"], curr_price - stop_distance)
                
                portfolio_msg += f"   💰 籌碼: {td['Chip_Status']}\n"
                if roi_net >= 20: pnl_alert = "💰【達標警戒】淨獲利突破 20%，建議分批獲利入袋！"
                elif curr_price < final_stop: pnl_alert = f"🩸【拔線警戒】跌破防守線 {final_stop:.1f}，強制離場！"
                else:
                    noc_state[sym]["trailing_stop"] = final_stop 
                    if roi_net > 0: pnl_alert = f"🔥 獲利巡航 | 📍 防線: {final_stop:.1f}"
                    else: pnl_alert = f"🟡 浮虧中 | 📍 死守: {final_stop:.1f}"

            portfolio_msg += f"   👉 指令: {pnl_alert}\n\n"
            msg_list.append(portfolio_msg)

    # === 🌟 處理觀察網域 (API 資源優化) ===
    for cat, stocks in STOCK_DICT.items():
        if not stocks: continue 
        cat_printed = False 
        
        for sym, name in stocks.items():
            # 觀察網域也進行 ETF 判定
            is_etf = sym.startswith("00") or "ETF" in name.upper()
            
            res = get_analysis_and_chart(sym, name, is_etf=is_etf)
            if not res: continue
                
            hist, chart_file = res
            td = hist.iloc[-1]
            has_data = True
                
            close = td['Close']
            atr = td['ATR']
            rsi = td['RSI']
            vma5 = td['5VMA']
            ma5 = td['5MA']
            ma20 = td['20MA']
            k = td['K']
            d = td['D']
            
            # 🛑 節省 API：ETF 直接 Bypass Yahoo 財報連線
            if is_etf:
                pe, yoy = "N/A", "N/A"
            else:
                pe = get_pe_ratio(sym)
                yoy = get_revenue_yoy(sym)
                
            vol_status = "📈 出量" if td['Volume'] > vma5 * 1.2 else "📉 量縮" if td['Volume'] < vma5 * 0.8 else "➖ 量平"
            trend_status = "🔥 多頭" if close > td['5MA'] > td['20MA'] else "🧊 空頭" if close < td['5MA'] < td['20MA'] else "🔄 盤整"
            
            yoy_label = f"{yoy:.2f}%" if isinstance(yoy, float) else yoy
            is_yoy_explosion = isinstance(yoy, float) and yoy >= YOY_EXPLOSION_PCT
            if is_yoy_explosion: yoy_label += " (🌟 業績大爆發)"

            kd_str = f"K:{k:.1f} D:{d:.1f}"
            if k < 30 and k > d and hist['K'].iloc[-2] <= hist['D'].iloc[-2]: kd_str += " (🌟 KD金叉)"
            elif k > 80: kd_str += " (⚠️ 短線過熱)"

            pe_str = f"{pe:.1f}" if isinstance(pe, float) else pe
            is_overvalued = isinstance(pe, float) and pe > PE_LIMIT

            predict_msg = "無特殊徵兆"
            if td['Volume'] > vma5 * 2 and rsi > 70: predict_msg = "💀【動能竭盡】高檔爆量轉折！"
            elif td['Shadow_Ratio'] > 0.5 and td['Volume'] > vma5 * 1.5: predict_msg = "⚠️【避雷針陷阱】高檔長上影線！"
            elif close > td['20_High'] and td['Volume'] > vma5 * 1.2: predict_msg = "🚀【無壓巡航】突破 20 日高！"
            elif td['BB_Width'] < 0.08: predict_msg = "⚠️【大變盤預警】通道極度壓縮！"
            
            stop_distance = atr * ATR_MULTIPLIER
            safe_stop_distance = stop_distance if not pd.isna(stop_distance) and stop_distance > 0 else 999999
            safe_close = close if not pd.isna(close) and close > 0 else 1.0
            
            suggested_shares = min(math.floor((TOTAL_CAPITAL * RISK_PER_TRADE) / safe_stop_distance), math.floor(TOTAL_CAPITAL / safe_close))
            
            sym_state = noc_state.get(sym, {"status": "NONE"})
            alert = "✅ 持股觀望"
            
            if sym_state["status"] == "REAL_HOLD": 
                alert = f"💼 已列入持股防禦區 | 📍 防線: {sym_state['trailing_stop']:.1f}"
            elif sym_state["status"] == "NONE":
                if td['Sniper_Signal']: 
                    if not is_bull_market: alert = "🛡️【大盤攔截】大盤偏空，放棄狙擊。"
                    elif isinstance(yoy, float) and yoy < 0: alert = f"🛡️【基本面攔截】營收衰退，避開地雷。"
                    elif is_overvalued: alert = f"🛡️【估值攔截】PE {pe_str} 過高，風險極大。"
                    else:
                        stop_price = close - stop_distance
                        noc_state[sym] = {"status": "HOLD", "entry": close, "trailing_stop": stop_price}
                        alert_prefix = "⚔️【雙劍合璧：終極狙擊】" if is_yoy_explosion else "🚀【啟動狙擊】"
                        alert = f"{alert_prefix}建議買入 {suggested_shares/1000:.1f} 張，停損設 {stop_price:.1f}"
                elif td['Sniper_Memory_5D'] == 1: 
                    alert = "🔥【狙擊延續】站穩5日線！" if close > td['5MA'] else "⚠️【狙擊失效】跌破5日線！"
            elif sym_state["status"] == "HOLD":
                new_stop = max(sym_state["trailing_stop"], close - stop_distance)
                if close < new_stop: 
                    alert = f"🩸【拔線離場】跌破防守線 {new_stop:.1f}！"
                    noc_state[sym] = {"status": "NONE"}
                else: 
                    noc_state[sym]["trailing_stop"] = new_stop
                    alert = f"🔥【波段抱牢】損益: {((close - sym_state['entry']) / sym_state['entry']) * 100:+.2f}% | 防守線: {new_stop:.1f}"

            write_noc_log(curr_date, sym, name, close, rsi, vol_status, trend_status, predict_msg, td['Chip_Status'], alert)
            tips = get_strategy_tips(symbol=sym, current_price=close, k_value=k, ma5=ma5, ma20=ma20)
            
            is_notable_signal = False
            if alert != "✅ 持股觀望": is_notable_signal = True
            if predict_msg != "無特殊徵兆": is_notable_signal = True
            if tips != "": is_notable_signal = True
            if td['Chip_Status'] in ["🤝 土洋齊買", "🏦 投信作帳(連買)"]: is_notable_signal = True
            
            if (not SILENT_MODE) or is_notable_signal:
                if not cat_printed: 
                    msg_list.append(f"━━━━━━━━━━━━━━\n📂 【{cat}】\n━━━━━━━━━━━━━━\n")
                    cat_printed = True
                
                if chart_file not in generated_charts: generated_charts.append(chart_file)
                    
                stock_msg = f"🔸 {name} ({sym})\n"
                stock_msg += f"   現價: {close:.2f} | PE: {pe_str} | 營收YoY: {yoy_label}\n"
                stock_msg += f"   指標: {kd_str} | RSI: {rsi:.1f}\n"
                stock_msg += f"   狀態: {trend_status} | {vol_status}\n"
                stock_msg += f"   💰 籌碼: {td['Chip_Status']}\n"
                stock_msg += f"   🔮 預判: {predict_msg}\n"
                stock_msg += f"   👉 指令: {alert}\n"
                if tips: stock_msg += f"   <i>{tips}</i>\n"
                stock_msg += "\n"
                msg_list.append(stock_msg)
            else:
                if os.path.exists(chart_file): os.remove(chart_file)

    if has_data or len(msg_list) > 0:
        save_state(noc_state) 
        if len(msg_list) == 1 and "大盤風向" in msg_list[0]:
            msg_list.append("\n🔕 【靜默模式】目前觀察網域無特殊狙擊訊號或危機，已省略一般觀望標的，詳細數據請參閱 CSV 日誌。")
            
        final_text = f"📡 【NOC 終極戰情室 v8.4 (雲端路由版)】\n📅 時間：{curr_time}\n━━━━━━━━━━━━━━\n" + "".join(msg_list)
        send_reports(f"NOC 戰情報告 {curr_date}", final_text, generated_charts)
        for chart in generated_charts:
            if os.path.exists(chart): os.remove(chart)
    else:
        print("休市或資料讀取失敗，伺服器待命。")
