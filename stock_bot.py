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

# === 1.1 量化基金風控參數 ===
TOTAL_CAPITAL = 1000000
RISK_PER_TRADE = 0.02
ATR_MULTIPLIER = 2.0
YOY_EXPLOSION_PCT = 50.0
PE_LIMIT = 40.0
SILENT_MODE = True       

# === 1.2 🌟 ETF 專屬判定引擎 (擴充更新版) ===
def get_etf_strategy(symbol, name):
    # 高股息與優息家族
    div_keys = ["高股息", "優息", "0056", "00878", "00919", "00929", "00915", "00713", "00939", "00940", "00936"]
    
    # 擴充：加入 00881、科技、半導體等「主題與市值型」關鍵字
    mkt_keys = ["0050", "006208", "市值", "AAPL", "NVDA", "TSM", "00881", "科技", "半導體", "5G", "00891", "00892"]
    
    if any(k in name or k in symbol for k in div_keys):
        return "💰高股息", 5.0, "控管殖利率 (5%乖離預警)"
    elif any(k in name or k in symbol for k in mkt_keys):
        return "🚀市值/主題型", 10.0, "成長動能區 (10%乖離預警)"
        
    return "🔸一般型", 8.0, "趨勢防禦區 (8%乖離預警)"

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
                
                if "庫存" in list_name or "庫藏" in list_name:
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
                        
                        if price_match: 
                            buy_price = float(price_match.group(1))
                        if shares_match: 
                            shares = int(shares_match.group(1))
                        
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

if TRELLO_DICT is not None and TRELLO_PORTFOLIO is not None:
    STOCK_DICT = TRELLO_DICT
    MY_PORTFOLIO = TRELLO_PORTFOLIO
else:
    print("⚠️ Trello 斷線，啟動緊急預設名單...")
    STOCK_DICT = {}
    MY_PORTFOLIO = {}

# === 2.2 🚀 動態掛載：讀取兩大雷達的傳令兵名單 ===
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
    try:
        url = "https://api.finmindtrade.com/api/v4/data"
        params = {"dataset": "TaiwanStockMonthRevenue", "data_id": match.group(), "start_date": (datetime.datetime.now() - datetime.timedelta(days=400)).strftime("%Y-%m-%d"), "token": FINMIND_TOKEN}
        r = requests.get(url, params=params, timeout=10).json()
        if r.get("msg") == "success" and r.get("data"):
            df = pd.DataFrame(r["data"])
            latest = df.iloc[-1]
            prev = df[(df['revenue_year'] == latest['revenue_year']-1) & (df['revenue_month'] == latest['revenue_month'])]
            if not prev.empty and prev.iloc[-1]['revenue'] > 0: 
                return float(((latest['revenue'] - prev.iloc[-1]['revenue']) / prev.iloc[-1]['revenue']) * 100)
    except Exception as e: 
        print(f"⚠️ [{symbol}] 營收處理錯誤: {e}")
    return "N/A"

def get_pe_ratio(symbol):
    try: 
        info = yf.Ticker(symbol).info
        return info.get('trailingPE', info.get('forwardPE', "N/A"))
    except: 
        return "N/A"

# === 5. FinMind 籌碼分析 (🌟 增強：投信連買賣天數) ===
def get_finmind_chip_data(symbol, start_date_str):
    if not FINMIND_TOKEN: return pd.DataFrame()
    match = re.search(r'\d+', symbol)
    if not match: return pd.DataFrame()
    try:
        url = "https://api.finmindtrade.com/api/v4/data"
        params = {"dataset": "TaiwanStockInstitutionalInvestorsBuySell", "data_id": match.group(), "start_date": start_date_str, "token": FINMIND_TOKEN}
        r = requests.get(url, params=params, timeout=10).json()
        if r.get("msg") == "success" and r.get("data"):
            df = pd.DataFrame(r["data"])
            df['net_buy'] = df['buy'] - df['sell']
            df['type'] = 'Other'
            df.loc[df['name'].str.contains('外資'), 'type'] = 'Foreign_Inv'
            df.loc[df['name'].str.contains('投信'), 'type'] = 'Trust_Inv'
            df.loc[df['name'].str.contains('自營商'), 'type'] = 'Dealer_Inv'
            pivot_df = df.groupby(['date', 'type'])['net_buy'].sum().unstack(fill_value=0).reset_index()
            for col in ['Foreign_Inv', 'Trust_Inv', 'Dealer_Inv']:
                if col not in pivot_df.columns: 
                    pivot_df[col] = 0
            pivot_df['Date'] = pd.to_datetime(pivot_df['date']).dt.date
            pivot_df.set_index('Date', inplace=True)
            return pivot_df[['Foreign_Inv', 'Trust_Inv', 'Dealer_Inv']]
    except Exception as e: 
        print(f"⚠️ [{symbol}] 籌碼讀取錯誤: {e}")
    return pd.DataFrame()

def calculate_chip_signals(hist: pd.DataFrame) -> pd.DataFrame:
    hist['Chip_Status'] = "無資料"
    hist['Trust_Streak'] = 0
    if all(col in hist.columns for col in ['Foreign_Inv', 'Trust_Inv', 'Dealer_Inv']):
        hist['Total_Institutional'] = hist['Foreign_Inv'] + hist['Trust_Inv'] + hist['Dealer_Inv']
        hist['Signal_CoBuy'] = (hist['Foreign_Inv'] > 0) & (hist['Trust_Inv'] > 0)
        hist['Signal_Trust_Trend'] = ((hist['Trust_Inv'] > 0).astype(int).rolling(5).sum() >= 4) & (hist['Trust_Inv'] > 0)
        
        # 🌟 精準計算投信連續買賣天數
        trust_dir = np.sign(hist['Trust_Inv'])
        hist['Trust_Streak'] = trust_dir.groupby((trust_dir != trust_dir.shift()).cumsum()).cumsum()
        
        conds = [(hist['Signal_CoBuy'] == True), (hist['Signal_Trust_Trend'] == True), (hist['Total_Institutional'] > 0)]
        hist['Chip_Status'] = np.select(conds, ["🤝 土洋齊買", "🏦 投信作帳", "📈 法人偏多"], default="➖ 中性/偏空")
    return hist

# === 5.5 特種戰略分析引擎 ===
def get_strategy_tips(symbol, current_price, k_value, ma5, ma20):
    if symbol == "9933.TW": 
        if current_price > ma5 and k_value < 30:
            return "🔥【NOC 訊號】中鼎疑似止跌！符合進場條件"
        else:
            return "⏳【NOC 監控】尚未止跌，繼續等待。"
    if symbol == "6415.TW": 
        if 240 <= current_price <= 260:
            return "💎【NOC 訊號】矽力進入支撐區" 
        else:
            return "🦅【NOC 監控】等待回測。"
    if symbol == "2303.TW": 
        if current_price > ma5:
            return "🚀【NOC 訊號】聯電強勢站穩 5MA！" 
        else:
            return "⚠️【NOC 警訊】聯電轉弱。"
    return ""

# === 6. 核心分析引擎 ===
def get_analysis_and_chart(symbol, name):
    try:
        stock = yf.Ticker(symbol)
        hist = stock.history(period="8mo")
        hist = hist.dropna(subset=['Close'])
        
        if len(hist) < 40: 
            return None
            
        hist['Date_Key'] = hist.index.date
        if FINMIND_TOKEN and (".TW" in symbol or ".TWO" in symbol):
            chip_df = get_finmind_chip_data(symbol, (datetime.datetime.now() - datetime.timedelta(days=200)).strftime("%Y-%m-%d"))
            if not chip_df.empty: 
                hist = hist.merge(chip_df, left_on='Date_Key', right_index=True, how='left')
                hist = hist.fillna(0)
                
        hist = calculate_chip_signals(hist)
        
        hist['5MA'] = hist['Close'].rolling(5).mean()
        hist['20MA'] = hist['Close'].rolling(20).mean()
        hist['5VMA'] = hist['Volume'].rolling(5).mean()

        # 🌟 增強：計算 60 日相對位階 (Price_Position)
        hist['High_60'] = hist['High'].rolling(window=60, min_periods=20).max()
        hist['Low_60'] = hist['Low'].rolling(window=60, min_periods=20).min()
        hist['Price_Position'] = (hist['Close'] - hist['Low_60']) / (hist['High_60'] - hist['Low_60'])

        l9 = hist['Low'].rolling(9).min()
        h9 = hist['High'].rolling(9).max()
        hist['K'] = ((hist['Close'] - l9) / (h9 - l9) * 100).ewm(com=2, adjust=False).mean()
        hist['D'] = hist['K'].ewm(com=2, adjust=False).mean()
        
        delta = hist['Close'].diff()
        gain = delta.clip(lower=0)
        loss = -delta.clip(upper=0)
        hist['RSI'] = 100 - (100 / (1 + (gain.ewm(com=13, adjust=False).mean() / loss.ewm(com=13, adjust=False).mean())))
        hist['RSI'] = hist['RSI'].fillna(50)
        
        hist['ATR'] = pd.concat([hist['High'] - hist['Low'], abs(hist['High'] - hist['Close'].shift(1)), abs(hist['Low'] - hist['Close'].shift(1))], axis=1).max(axis=1).rolling(14).mean()
        
        hist['MACD'] = hist['Close'].ewm(span=12, adjust=False).mean() - hist['Close'].ewm(span=26, adjust=False).mean()
        hist['MACD_Hist'] = hist['MACD'] - hist['MACD'].ewm(span=9, adjust=False).mean()
        hist['STD20'] = hist['Close'].rolling(20).std()
        hist['BB_Width'] = (4 * hist['STD20']) / hist['20MA']
        
        hist['Is_Bottoming'] = ((hist['Close'] < hist['5MA']) & (hist['MACD_Hist'].shift(2) < hist['MACD_Hist'].shift(1)) & (hist['MACD_Hist'].shift(1) < hist['MACD_Hist']) & (hist['MACD_Hist'] < 0)).astype(int)
        hist['Is_Breakout'] = (hist['Close'].shift(1) < hist['5MA'].shift(1)) & (hist['Close'] > hist['5MA']) & (hist['Volume'] > hist['5VMA'] * 1.2)
        hist['Sniper_Signal'] = hist['Is_Bottoming'].rolling(3).max().fillna(0).astype(bool) & hist['Is_Breakout']
        hist['Sniper_Memory_5D'] = hist['Sniper_Signal'].rolling(5).max().fillna(0)
        
        hist['20_High'] = hist['High'].rolling(20).max().shift(1)
        hist['Shadow_Ratio'] = (hist['High'] - hist[['Open', 'Close']].max(axis=1)) / (hist['High'] - hist['Low']).replace(0, 0.001)
        
        chart_file = f"{symbol}_chart.png"
        try:
            mc = mpf.make_marketcolors(up='red', down='green', edge='black', wick='black', volume='gray')
            mpf.plot(hist[-60:], type='candle', style=mpf.make_mpf_style(base_mpf_style='yahoo', marketcolors=mc), volume=True, mav=(5, 20), title=f"Stock: {symbol}", savefig=chart_file)
        except:
            mpf.plot(hist[-60:], type='candle', style='yahoo', volume=True, mav=(5, 20), title=f"Stock: {symbol}", savefig=chart_file)
            
        return hist, chart_file
    except Exception as e: 
        print(f"[{symbol}] 核心分析發生錯誤: {e}")
        return None

# === 7. 發送模組 ===
def send_reports(subject, text_body, chart_files):
    if TG_TOKEN and TG_CHAT_ID:
        try:
            for part in [text_body[i:i + 4000] for i in range(0, len(text_body), 4000)]:
                requests.post(f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage", json={"chat_id": TG_CHAT_ID, "text": part, "disable_web_page_preview": True}, timeout=10)
        except Exception as e: 
            print(f"❌ Telegram 發送失敗: {e}")
            
    if EMAIL_USER and EMAIL_PASS and EMAIL_TO:
        try:
            msg = MIMEMultipart()
            msg['From'] = EMAIL_USER
            msg['To'] = EMAIL_TO
            msg['Subject'] = subject
            msg.attach(MIMEText(text_body, 'plain'))
            
            for chart in chart_files:
                if os.path.exists(chart): 
                    msg.attach(MIMEImage(open(chart, 'rb').read(), name=os.path.basename(chart)))
                    
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
                server.login(EMAIL_USER, EMAIL_PASS)
                server.send_message(msg)
        except Exception as e: 
            print(f"❌ Email 發送失敗: {e}")

# === 8. 主程式執行 (五大戰區顯示分流邏輯) ===
if __name__ == "__main__":
    tw_tz = datetime.timezone(datetime.timedelta(hours=8))
    curr_date = datetime.datetime.now(tw_tz).date()
    curr_time = datetime.datetime.now(tw_tz).strftime("%Y-%m-%d %H:%M:%S")
    msg_list = []
    generated_charts = []
    has_data = False
    
    print(f"[{curr_time}] NOC 終極戰情室 v8.12 (量價位階與籌碼強化版) 啟動...")
    
    is_bull_market, market_msg = get_market_regime()
    noc_state = load_state()
    msg_list.append(f"🌐 【大盤風向】: {market_msg}\n")

    # === 戰區 1：庫藏股 ===
    if MY_PORTFOLIO:
        msg_list.append("━━━━━━━━━━━━━━\n💼 【庫藏股 (實體持股動態防禦)】\n━━━━━━━━━━━━━━\n")
        for sym, data in MY_PORTFOLIO.items():
            res = get_analysis_and_chart(sym, data['name'])
            if not res: 
                continue
                
            hist, chart_file = res
            td = hist.iloc[-1]
            has_data = True
            generated_charts.append(chart_file)
            
            curr_price = td['Close']
            atr = td['ATR']
            buy_price = data['buy_price']
            roi_pct = ((curr_price - buy_price) / buy_price) * 100
            
            # 🌟 1. 提早呼叫 ETF 判定引擎，確認標的屬性
            etf_icon, bias_limit, etf_desc = get_etf_strategy(sym, data['name'])
            is_etf = "一般型" not in etf_icon  # 如果不是一般型，就是大盤或高股息 ETF
            
            sym_state = noc_state.get(sym, {"status": "NONE"})
            
            # 🌟 2. 啟動分流防禦機制
            if is_etf:
                # 🛡️ 【ETF 紀律模式】：關閉 ATR 停損，啟動越跌越買邏輯
                if sym_state["status"] != "REAL_HOLD_ETF":
                    noc_state[sym] = {"status": "REAL_HOLD_ETF", "entry": buy_price}
                
                # ETF 的操作指令：不看防守線，看回檔深度
                if roi_pct <= -10.0:
                    pnl_alert = f"💎【黃金坑加碼】帳面回檔 {roi_pct:.2f}%，啟動大額建倉！"
                elif roi_pct <= -5.0:
                    pnl_alert = f"📉【紀律扣款】帳面回檔 {roi_pct:.2f}%，維持定期定額。"
                else:
                    pnl_alert = f"🧘‍♂️【長線鎖籌】無懼波動，靜待資產翻倍。"
                    
            else:
                # ⚔️ 【個股波段模式】：維持原有的 2倍 ATR 動態停損
                stop_distance = atr * ATR_MULTIPLIER
                
                if sym_state["status"] != "REAL_HOLD":
                    # 重新進場或初次抓取，設定初始防線
                    noc_state[sym] = {"status": "REAL_HOLD", "entry": buy_price, "trailing_stop": curr_price - stop_distance}
                    sym_state = noc_state[sym]
                    
                # 計算當前應該墊高到的防守線
                final_stop = max(sym_state["trailing_stop"], curr_price - stop_distance)
                
                if curr_price < final_stop: 
                    pnl_alert = f"🩸【拔線警戒】跌破防守線 {final_stop:.1f}，請嚴格執行離場！"
                else:
                    noc_state[sym]["trailing_stop"] = final_stop 
                    if roi_pct > 0: 
                        pnl_alert = f"🔥 獲利巡航 | 📍 防線墊高至: {final_stop:.1f}"
                    else: 
                        pnl_alert = f"🟡 浮虧防禦 | 📍 死守底線: {final_stop:.1f}"
            
            # 🌟 3. 輸出戰報訊息
            portfolio_msg = f"{etf_icon} {data['name']} ({sym})\n"
            portfolio_msg += f"   成本: {buy_price:.2f} | 股數: {data['shares']} | 現價: {curr_price:.2f}\n"
            portfolio_msg += f"   損益: {roi_pct:+.2f}% | 👉 指令: {pnl_alert}\n\n"
            msg_list.append(portfolio_msg)

    # === Trello & 本地雷達 觀察網域分流 ===
    for cat, stocks in STOCK_DICT.items():
        if not stocks: 
            continue 
        
        # 🌟 判定五大戰區類別
        is_etf_zone = "ETF" in cat.upper()
        is_radar_zone = "雷達" in cat
        is_key_obs = "重點觀測" in cat
        is_normal_obs = "觀察" in cat and not is_key_obs and not is_radar_zone
        
        cat_msg_list = [] 
        
        for sym, name in stocks.items():
            res = get_analysis_and_chart(sym, name)
            if not res: 
                continue
                
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
            pe = get_pe_ratio(sym)
            
            # 取得新計算的位階與投信連買賣資料
            pos = td['Price_Position'] if 'Price_Position' in td and not pd.isna(td['Price_Position']) else 0.5
            trust_streak = int(td['Trust_Streak']) if 'Trust_Streak' in td and not pd.isna(td['Trust_Streak']) else 0
            
            bias = ((close - ma20) / ma20) * 100 if ma20 else 0
            
            if td['Volume'] > vma5 * 1.2:
                vol_status = "📈 出量"
            elif td['Volume'] < vma5 * 0.8:
                vol_status = "📉 量縮"
            else:
                vol_status = "➖ 量平"
                
            if close > ma5 > ma20:
                trend_status = "🔥 多頭"
            elif close < ma5 < ma20:
                trend_status = "🧊 空頭"
            else:
                trend_status = "🔄 盤整"
            
            yoy = get_revenue_yoy(sym)
            yoy_label = f"{yoy:.2f}%" if isinstance(yoy, float) else yoy
            if isinstance(yoy, float) and yoy >= YOY_EXPLOSION_PCT: 
                yoy_label += " (🌟 業績大爆發)"

            kd_str = f"K:{k:.1f} D:{d:.1f}"
            if k < 30 and k > d and hist['K'].iloc[-2] <= hist['D'].iloc[-2]: 
                kd_str += " (🌟 KD金叉)"
            elif k > 80: 
                kd_str += " (⚠️ 短線過熱)"

            pe_str = f"{pe:.1f}" if isinstance(pe, float) else pe
            is_overvalued = isinstance(pe, float) and pe > PE_LIMIT

            # 🌟 籌碼增強：將投信天數整合進狀態顯示
            chip_msg = td['Chip_Status']
            if trust_streak > 0:
                chip_msg += f" (連買 {trust_streak} 天)"
            elif trust_streak < 0:
                chip_msg += f" (連賣 {abs(trust_streak)} 天)"

            # 🌟 預判增強：納入 60 日股價位階 (高於 0.7 屬高檔，低於 0.3 屬低檔)
            predict_msg = "無特殊徵兆"
            if td['Volume'] > vma5 * 2: 
                if pos > 0.7:
                    predict_msg = "💀【動能竭盡】高檔爆量轉折！"
                elif pos < 0.3:
                    predict_msg = "🔥【底部換手】低檔爆量，醞釀反彈！"
                else:
                    predict_msg = "⚠️【中繼爆量】留意方向表態！"
            elif td['Shadow_Ratio'] > 0.5 and td['Volume'] > vma5 * 1.5: 
                if pos > 0.7:
                    predict_msg = "⚠️【避雷針陷阱】高檔長上影線！"
                elif pos < 0.3:
                    predict_msg = "🌟【仙人指路】低檔長上影線試盤！"
            elif close > td['20_High'] and td['Volume'] > vma5 * 1.2: 
                predict_msg = "🚀【無壓巡航】突破 20 日高！"
            elif td['BB_Width'] < 0.08: 
                predict_msg = "⚠️【大變盤預警】通道極度壓縮！"
            
            stop_distance = atr * ATR_MULTIPLIER
            safe_stop_distance = stop_distance if stop_distance > 0 else 999999
            suggested_shares = min(math.floor((TOTAL_CAPITAL * RISK_PER_TRADE) / safe_stop_distance), math.floor(TOTAL_CAPITAL / (close if close > 0 else 1.0)))
            
            sym_state = noc_state.get(sym, {"status": "NONE"})
            alert = "✅ 持股觀望"
            
            if sym_state["status"] == "REAL_HOLD": 
                alert = f"💼 已列入持股防禦區 | 📍 防線: {sym_state['trailing_stop']:.1f}"
            elif sym_state["status"] == "NONE":
                if td['Sniper_Signal']: 
                    if not is_bull_market: 
                        alert = "🛡️【大盤攔截】大盤偏空，放棄狙擊。"
                    elif isinstance(yoy, float) and yoy < 0: 
                        alert = "🛡️【基本面攔截】營收衰退，避開地雷。"
                    elif is_overvalued: 
                        alert = f"🛡️【估值攔截】PE {pe_str} 過高，風險極大。"
                    else:
                        stop_price = close - stop_distance
                        noc_state[sym] = {"status": "HOLD", "entry": close, "trailing_stop": stop_price}
                        alert = f"{'⚔️【雙劍合璧】' if isinstance(yoy, float) and yoy >= YOY_EXPLOSION_PCT else '🚀【啟動狙擊】'}買入 {suggested_shares/1000:.1f} 張，停損 {stop_price:.1f}"
                elif td['Sniper_Memory_5D'] == 1: 
                    if close > ma5:
                        alert = "🔥【狙擊延續】站穩5日線！" 
                    else:
                        alert = "⚠️【狙擊失效】跌破5日線！"
            elif sym_state["status"] == "HOLD":
                new_stop = max(sym_state["trailing_stop"], close - stop_distance)
                if close < new_stop: 
                    alert = f"🩸【拔線離場】跌破防守線 {new_stop:.1f}！"
                    noc_state[sym] = {"status": "NONE"}
                else: 
                    noc_state[sym]["trailing_stop"] = new_stop
                    alert = f"🔥【波段抱牢】防守線: {new_stop:.1f}"

            write_noc_log(curr_date, sym, name, close, rsi, vol_status, trend_status, predict_msg, chip_msg, alert)
            tips = get_strategy_tips(sym, close, k, ma5, ma20)

            # --- 戰區 2：ETF 專區 ---
            if is_etf_zone:
                etf_type, bias_limit, etf_desc = get_etf_strategy(sym, name)
                if bias > bias_limit: 
                    etf_cmd = "⚠️ 乖離過熱，建議分批獲利了結"
                elif k < 30 and k > d: 
                    etf_cmd = "🔥 KD低檔金叉，建議佈局買進"
                elif close > ma5: 
                    etf_cmd = "✅ 趨勢向上，續抱"
                else: 
                    etf_cmd = "⏳ 趨勢偏弱，觀望"
                
                stock_msg = f"{etf_type} {name} ({sym})\n"
                stock_msg += f"   現價: {close:.2f} | 乖離: {bias:+.1f}% ({'🚨過熱' if bias > bias_limit else '✅穩定'})\n"
                stock_msg += f"   屬性: {etf_desc}\n"
                stock_msg += f"   👉 指令: {etf_cmd}\n\n"
                cat_msg_list.append(stock_msg)
                
                if chart_file not in generated_charts: 
                    generated_charts.append(chart_file)

            # --- 戰區 5：🎯 雷達鎖定區 (短線狙擊專用) ---
            elif is_radar_zone:
                stock_msg = f"🎯 {name} ({sym})\n"
                stock_msg += f"   現價: {close:.2f} | 狀態: {trend_status} | {vol_status}\n"
                stock_msg += f"   指標: {kd_str} | RSI: {rsi:.1f}\n"
                stock_msg += f"   💰 籌碼: {chip_msg}\n"
                stock_msg += f"   👉 指令: {alert}\n"
                if tips: 
                    stock_msg += f"   💡 戰略提示: {tips}\n"
                stock_msg += "\n"
                cat_msg_list.append(stock_msg)
                
                if chart_file not in generated_charts: 
                    generated_charts.append(chart_file)

            # --- 戰區 3：重點觀測區 ---
            elif is_key_obs:
                etf_icon, _, etf_desc = get_etf_strategy(sym, name)
                stock_msg = f"{etf_icon} {name} ({sym})\n"
                stock_msg += f"   現價: {close:.2f} | 乖離: {bias:+.1f}% | PE: {pe_str}\n"
                stock_msg += f"   指標: {kd_str} | RSI: {rsi:.1f} | 類型: {etf_desc}\n"
                stock_msg += f"   狀態: {trend_status} | YoY: {yoy_label}\n"
                stock_msg += f"   💰 籌碼: {chip_msg}\n"
                stock_msg += f"   🔮 預判: {predict_msg}\n"
                stock_msg += f"   👉 指令: {alert}\n"
                if tips: 
                    stock_msg += f"   💡 NOC訊號: {tips}\n"
                stock_msg += "\n"
                cat_msg_list.append(stock_msg)
                
                if chart_file not in generated_charts: 
                    generated_charts.append(chart_file)

            # --- 戰區 4：觀察區 (條件觸發才顯示) ---
            elif is_normal_obs:
                # 🌟 更新陷阱與復甦的觸發條件，納入新版位階預判
                is_trap = predict_msg in ["💀【動能竭盡】高檔爆量轉折！", "⚠️【避雷針陷阱】高檔長上影線！", "⚠️【大變盤預警】通道極度壓縮！"]
                is_recovery = td['Sniper_Signal'] or (k < 30 and k > d) or ("止跌" in tips or "支撐" in tips) or predict_msg in ["🔥【底部換手】低檔爆量，醞釀反彈！", "🌟【仙人指路】低檔長上影線試盤！"]
                
                if is_trap or is_recovery:
                    stock_msg = f"👀 {name} ({sym})\n"
                    stock_msg += f"   現價: {close:.2f} | RSI: {rsi:.1f} | 乖離: {bias:+.1f}%\n"
                    stock_msg += f"   💰 籌碼: {chip_msg}\n"
                    stock_msg += f"   🎯 條件觸發: {'🚨 陷阱預警' if is_trap else '🔥 復甦/狙擊訊號'}\n"
                    stock_msg += f"   👉 預判/指令: {predict_msg if is_trap else alert}\n"
                    if tips: 
                        stock_msg += f"   💡 NOC訊號: {tips}\n"
                    stock_msg += "\n"
                    cat_msg_list.append(stock_msg)
                    
                    if chart_file not in generated_charts: 
                        generated_charts.append(chart_file)
                else:
                    if os.path.exists(chart_file): 
                        os.remove(chart_file)
                    continue
                    
            # --- 其他未分類戰區 ---
            else:
                stock_msg = f"🔸 {name} ({sym})\n"
                stock_msg += f"   現價: {close:.2f} | 狀態: {trend_status}\n"
                stock_msg += f"   👉 指令: {alert}\n\n"
                cat_msg_list.append(stock_msg)
                
                if chart_file not in generated_charts: 
                    generated_charts.append(chart_file)

        if cat_msg_list:
            msg_list.append(f"━━━━━━━━━━━━━━\n📂 【{cat}】\n━━━━━━━━━━━━━━\n")
            msg_list.extend(cat_msg_list)
    # =========================================================
    # === 9. 🏆 ETF 雙引擎績效競技場 (自動汰弱留強模組) ===
    # =========================================================
    etf_arena = {"💰高股息防禦組": [], "🚀市值與主題成長組": []}
    current_year = curr_date.year

    # 收集所有出現在雷達與庫藏中的唯一 ETF 標的
    all_etfs = {}
    if MY_PORTFOLIO:
        for sym, data in MY_PORTFOLIO.items():
            all_etfs[sym] = data['name']
    for cat, stocks in STOCK_DICT.items():
        if stocks:
            for sym, name in stocks.items():
                all_etfs[sym] = name

    # 進入競技場後台運算
    for sym, name in all_etfs.items():
        etf_icon, _, _ = get_etf_strategy(sym, name)
        is_etf = "一般型" not in etf_icon
        
        if is_etf:
            # 取得歷史資料 (由於前面戰區可能已抓過，若有快取機制更好，這裡為求穩定直接調用)
            res = get_analysis_and_chart(sym, name)
            if not res: 
                continue
            hist, _ = res
            
            if len(hist) < 10: 
                continue
            
            close_price = hist['Close'].iloc[-1]
            
            # 運算 1：近一季 (60個交易日) 動能
            qtr_days = min(60, len(hist)-1)
            qtr_price = hist['Close'].iloc[-(qtr_days+1)]
            qtr_roi = ((close_price - qtr_price) / qtr_price) * 100
            
            # 運算 2：今年以來 (YTD) 績效
            hist_ytd = hist[hist.index.year == current_year]
            if not hist_ytd.empty:
                ytd_start_price = hist_ytd['Close'].iloc[0]
                ytd_roi = ((close_price - ytd_start_price) / ytd_start_price) * 100
            else:
                ytd_roi = qtr_roi # 若無今年初資料防呆機制
                
            group_key = "💰高股息防禦組" if "高股息" in etf_icon else "🚀市值與主題成長組"
            etf_arena[group_key].append({
                "name": name, 
                "sym": sym, 
                "qtr_roi": qtr_roi, 
                "ytd_roi": ytd_roi
            })

    # 產出競技場戰報文字
    arena_msg = []
    if etf_arena["💰高股息防禦組"] or etf_arena["🚀市值與主題成長組"]:
        arena_msg.append("━━━━━━━━━━━━━━\n🏆 【ETF 雙引擎績效競技場 (自動汰弱留強)】\n━━━━━━━━━━━━━━\n")
        
        for group_name, group_data in etf_arena.items():
            if not group_data: 
                continue
            arena_msg.append(f"**{group_name}**\n")
            
            # 依據季動能 (qtr_roi) 降冪排序，動能強的排前面
            sorted_etfs = sorted(group_data, key=lambda x: x['qtr_roi'], reverse=True)
            
            medals = ["🥇", "🥈", "🥉"]
            for idx, etf in enumerate(sorted_etfs):
                medal = medals[idx] if idx < 3 else "🔸"
                q_str = f"{etf['qtr_roi']:+.1f}%"
                y_str = f"{etf['ytd_roi']:+.1f}%"
                
                # AI 智能評語邏輯
                if etf['qtr_roi'] > 5.0 and etf['ytd_roi'] > 10.0:
                    status = "🔥 雙料強勢"
                elif etf['qtr_roi'] < 0 and etf['ytd_roi'] > 0:
                    status = "⏳ 短線洗盤，長線穩健"
                elif etf['qtr_roi'] < -2.0 and etf['ytd_roi'] < 0:
                    status = "⚠️ 嚴重落後，請檢視佔比"
                else:
                    status = "✅ 穩定跟隨"
                    
                arena_msg.append(f"{medal} {etf['name']} ({etf['sym']})\n   季動能 {q_str} ｜ 本年累計 {y_str} ({status})\n")
            arena_msg.append("\n")
            
        msg_list.extend(arena_msg)
    # =========================================================
    
    if has_data or len(msg_list) > 0:
        save_state(noc_state) 
        if len(msg_list) == 1 and "大盤風向" in msg_list[0]: 
            msg_list.append("\n🔕 【靜默模式】無觸發條件。")
            
        final_text = f"📡 【NOC 終極戰情室 v8.12 (量價位階強化版)】\n📅 時間：{curr_time}\n━━━━━━━━━━━━━━\n" + "".join(msg_list)
        send_reports(f"NOC 戰情報告 {curr_date}", final_text, generated_charts)
        
        for chart in generated_charts:
            if os.path.exists(chart): 
                os.remove(chart)
    else:
        print("休市或資料讀取失敗，伺服器待命。")
