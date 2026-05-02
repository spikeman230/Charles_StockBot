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
import sys
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

# ==========================================
# 🌟 系統核心架構升級：全域記憶體快取
# ==========================================
GLOBAL_DATA_CACHE = {}

# === 1.2 🌟 ETF 專屬判定引擎 ===
def get_etf_strategy(symbol, name):
    div_keys = ["高股息", "優息", "0056", "00878", "00919", "00929", "00915", "00713", "00939", "00940", "00936"]
    mkt_keys = ["0050", "006208", "市值", "AAPL", "NVDA", "TSM", "00881", "科技", "半導體", "5G", "00891", "00892", "009816"]
    
    if any(k in name or k in symbol for k in div_keys): return "💰高股息", 5.0, "控管殖利率 (5%乖離預警)"
    elif any(k in name or k in symbol for k in mkt_keys): return "🚀市值/主題型", 10.0, "成長動能區 (10%乖離預警)"
    return "🔸一般型", 8.0, "趨勢防禦區 (8%乖離預警)"

# === 1.3 🌟 開市絕對攔截機制 ===
def is_trading_day(curr_date):
    try:
        tsm = yf.Ticker("2330.TW").history(period="1d")
        if tsm.empty: return False
        return tsm.index[-1].date() == curr_date
    except Exception as e:
        print(f"⚠️ 交易日判斷 API 異常: {e}")
        return curr_date.weekday() < 5

# === 2. 🌟 升級版 Trello 雲端資料庫 (智慧戰略解析) ===
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
                        name_part = raw_name[len(symbol):].strip() if ticker_match else raw_name
                        name = re.sub(r'\(.*?\)', '', name_part).strip() if name_part else symbol
                        
                        desc = card.get('desc', '')
                        buy_price = 0.0
                        shares = 1000
                        
                        price_match = re.search(r"成本[：:]\s*([0-9.]+)", desc)
                        shares_match = re.search(r"股數[：:]\s*([0-9]+)", desc)
                        
                        if price_match: buy_price = float(price_match.group(1))
                        if shares_match: shares = int(shares_match.group(1))
                        
                        my_portfolio[symbol] = {"name": name, "buy_price": buy_price, "shares": shares, "trello_tip": desc}
                else:
                    stock_list = {}
                    for card in cards:
                        raw_name = card['name'].strip()
                        ticker_match = re.match(r'^[A-Za-z0-9.]+', raw_name)
                        symbol = ticker_match.group() if ticker_match else raw_name
                        name_part = raw_name[len(symbol):].strip() if ticker_match else raw_name
                        
                        # 清理括號，取得乾淨名稱
                        name = re.sub(r'\(.*?\)', '', name_part).strip() if name_part else symbol
                        
                        # 💡 智慧戰略解耦：抓取括號內的文字，若無則抓取描述
                        title_tip_match = re.search(r'\((.*?)\)', name_part)
                        trello_tip = title_tip_match.group(1) if title_tip_match else card.get('desc', '').strip()
                        
                        stock_list[symbol] = {"name": name, "trello_tip": trello_tip}
                        
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

# === 2.2 🚀 動態掛載：讀取兩大雷達名單 (由 Auto-commit 與分離程式產生) ===
RADAR_FILE = "radar_targets.json"
if os.path.exists(RADAR_FILE):
    try:
        with open(RADAR_FILE, "r", encoding="utf-8") as f:
            radar_stocks = json.load(f)
            if radar_stocks: STOCK_DICT["🎯 雷達鎖定 (新進火種區)"] = radar_stocks
    except Exception as e: print(f"⚠️ 游擊隊雷達名單讀取失敗: {e}")

LIGHTNING_FILE = "lightning_targets.json"
if os.path.exists(LIGHTNING_FILE):
    try:
        with open(LIGHTNING_FILE, "r", encoding="utf-8") as f:
            lightning_stocks = json.load(f)
            if lightning_stocks: STOCK_DICT["⚡ 雷達鎖定 (短線飆股區)"] = lightning_stocks
    except Exception as e: print(f"⚠️ 閃電突擊名單讀取失敗: {e}")

# === 3. 實體狀態記憶庫與日誌 ===
STATE_FILE = "noc_state.json"

def load_state():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f: return json.load(f)
        except: pass
    return {}

def save_state(state_data):
    try:
        with open(STATE_FILE, "w", encoding="utf-8") as f: json.dump(state_data, f, ensure_ascii=False, indent=4)
    except: pass

def write_noc_log(date, symbol, name, close_price, rsi, vol_status, status, predict, chip_signal, alert):
    log_filename = "noc_trading_log.csv"
    file_exists = os.path.exists(log_filename)
    try:
        with open(log_filename, mode='a', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            if not file_exists: writer.writerow(["日期", "代號", "名稱", "收盤價", "RSI", "量能狀態", "趨勢狀態", "戰場預判", "籌碼訊號", "行動指令"])
            writer.writerow([date, symbol, name, f"{close_price:.2f}", f"{rsi:.2f}", vol_status, status, predict, chip_signal, alert])
    except: pass

# === 4. 環境感知：大盤、營收與估值 ===
def get_market_regime():
    try:
        twii = yf.Ticker("^TWII").history(period="1mo")
        twii['20MA'] = twii['Close'].rolling(20).mean()
        if twii['Close'].iloc[-1] > twii['20MA'].iloc[-1]: return True, f"🟢 多頭格局 (站上月線)"
        else: return False, f"🔴 空頭警戒 (跌破月線)"
    except: return True, "🟡 大盤狀態未知"

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
    except: pass
    return "N/A"

def get_pe_ratio(symbol):
    try: 
        info = yf.Ticker(symbol).info
        return info.get('trailingPE', info.get('forwardPE', "N/A"))
    except: return "N/A"

# === 5. FinMind 籌碼分析 ===
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
                if col not in pivot_df.columns: pivot_df[col] = 0
            pivot_df['Date'] = pd.to_datetime(pivot_df['date']).dt.date
            pivot_df.set_index('Date', inplace=True)
            return pivot_df[['Foreign_Inv', 'Trust_Inv', 'Dealer_Inv']]
    except: pass
    return pd.DataFrame()

def calculate_chip_signals(hist: pd.DataFrame) -> pd.DataFrame:
    hist['Chip_Status'] = "無資料"
    hist['Trust_Streak'] = 0
    if all(col in hist.columns for col in ['Foreign_Inv', 'Trust_Inv', 'Dealer_Inv']):
        hist['Total_Institutional'] = hist['Foreign_Inv'] + hist['Trust_Inv'] + hist['Dealer_Inv']
        hist['Signal_CoBuy'] = (hist['Foreign_Inv'] > 0) & (hist['Trust_Inv'] > 0)
        hist['Signal_Trust_Trend'] = ((hist['Trust_Inv'] > 0).astype(int).rolling(5).sum() >= 4) & (hist['Trust_Inv'] > 0)
        
        trust_dir = np.sign(hist['Trust_Inv'])
        hist['Trust_Streak'] = trust_dir.groupby((trust_dir != trust_dir.shift()).cumsum()).cumsum()
        
        conds = [(hist['Signal_CoBuy'] == True), (hist['Signal_Trust_Trend'] == True), (hist['Total_Institutional'] > 0)]
        hist['Chip_Status'] = np.select(conds, ["🤝 土洋齊買", "🏦 投信作帳", "📈 法人偏多"], default="➖ 中性/偏空")
    return hist

# === 6. 🌟 極速資料獲取與指標引擎 (快取版) ===
def get_stock_data(symbol, name):
    # ⚡ 命中快取，直接回傳！
    if symbol in GLOBAL_DATA_CACHE:
        return GLOBAL_DATA_CACHE[symbol]
        
    try:
        stock = yf.Ticker(symbol)
        hist = stock.history(period="8mo")
        hist = hist.dropna(subset=['Close'])
        
        if len(hist) < 40: return None
            
        hist['Date_Key'] = hist.index.date
        if FINMIND_TOKEN and (".TW" in symbol or ".TWO" in symbol):
            chip_df = get_finmind_chip_data(symbol, (datetime.datetime.now() - datetime.timedelta(days=200)).strftime("%Y-%m-%d"))
            if not chip_df.empty: 
                hist = hist.merge(chip_df, left_on='Date_Key', right_index=True, how='left')
                hist = hist.fillna(0)
                
        hist = calculate_chip_signals(hist)
        hist['5MA'] = hist['Close'].rolling(5).mean()
        hist['20MA'] = hist['Close'].rolling(20).mean()
        
        curr_hour = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).hour
        vol_multiplier = 1.0
        if curr_hour == 10: vol_multiplier = 4.5
        elif curr_hour == 12: vol_multiplier = 1.5
        elif curr_hour == 13: vol_multiplier = 1.1 

        hist['Est_Volume'] = hist['Volume'].copy()
        if len(hist) > 0:
            hist.iloc[-1, hist.columns.get_loc('Est_Volume')] = hist['Volume'].iloc[-1] * vol_multiplier
            
        hist['5VMA'] = hist['Est_Volume'].rolling(5).mean()

        hist['25MA'] = hist['Close'].rolling(25).mean()
        hist['60VMA'] = hist['Volume'].rolling(60).mean()
        
        cond_trend = hist['25MA'] > hist['25MA'].shift(3)
        cond_vol_mom = hist['5VMA'] > hist['60VMA']
        cond_pullback = (hist['Low'] <= hist['25MA'] * 1.015) & (hist['Close'] >= hist['25MA'] * 0.985)
        cond_shrink = hist['Est_Volume'] < hist['5VMA'] 
        
        hist['Signal_2560'] = cond_trend & cond_vol_mom & cond_pullback & cond_shrink

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
        hist['Is_Breakout'] = (hist['Close'].shift(1) < hist['5MA'].shift(1)) & (hist['Close'] > hist['5MA']) & (hist['Est_Volume'] > hist['5VMA'] * 1.2)
        hist['Sniper_Signal'] = hist['Is_Bottoming'].rolling(3).max().fillna(0).astype(bool) & hist['Is_Breakout']
        hist['Sniper_Memory_5D'] = hist['Sniper_Signal'].rolling(5).max().fillna(0)
        
        hist['20_High'] = hist['High'].rolling(20).max().shift(1)
        hist['Shadow_Ratio'] = (hist['High'] - hist[['Open', 'Close']].max(axis=1)) / (hist['High'] - hist['Low']).replace(0, 0.001)
        
        # 存入全域快取
        GLOBAL_DATA_CACHE[symbol] = hist
        return hist
    except Exception as e: 
        print(f"[{symbol}] 核心分析發生錯誤: {e}")
        return None

# === 6.5 🌟 延遲渲染引擎 (省算力) ===
def draw_chart_if_needed(hist, symbol):
    chart_file = f"{symbol}_chart.png"
    try:
        mc = mpf.make_marketcolors(up='red', down='green', edge='black', wick='black', volume='gray')
        mpf.plot(hist[-60:], type='candle', style=mpf.make_mpf_style(base_mpf_style='yahoo', marketcolors=mc), volume=True, mav=(5, 20), title=f"Stock: {symbol}", savefig=chart_file)
    except:
        mpf.plot(hist[-60:], type='candle', style='yahoo', volume=True, mav=(5, 20), title=f"Stock: {symbol}", savefig=chart_file)
    return chart_file

# === 7. 發送模組 ===
def send_reports(subject, text_body, chart_files):
    if TG_TOKEN and TG_CHAT_ID:
        try:
            for part in [text_body[i:i + 4000] for i in range(0, len(text_body), 4000)]:
                requests.post(f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage", json={"chat_id": TG_CHAT_ID, "text": part, "disable_web_page_preview": True}, timeout=10)
        except Exception as e: print(f"❌ Telegram 發送失敗: {e}")
            
    if EMAIL_USER and EMAIL_PASS and EMAIL_TO:
        try:
            msg = MIMEMultipart()
            msg['From'] = EMAIL_USER
            msg['To'] = EMAIL_TO
            msg['Subject'] = subject
            msg.attach(MIMEText(text_body, 'plain'))
            
            for chart in chart_files:
                if os.path.exists(chart): msg.attach(MIMEImage(open(chart, 'rb').read(), name=os.path.basename(chart)))
                    
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
                server.login(EMAIL_USER, EMAIL_PASS)
                server.send_message(msg)
        except Exception as e: print(f"❌ Email 發送失敗: {e}")

# === 8. 主程式執行 (五大戰區顯示分流邏輯) ===
if __name__ == "__main__":
    tw_tz = datetime.timezone(datetime.timedelta(hours=8))
    curr_date = datetime.datetime.now(tw_tz).date()
    curr_time = datetime.datetime.now(tw_tz).strftime("%Y-%m-%d %H:%M:%S")

    print(f"[{curr_time}] NOC 終極戰情室啟動，檢查開市狀態...")
    if not is_trading_day(curr_date):
        print("📅 今日為週末或國定假日休市，戰情室暫停推播，伺服器進入休眠。")
        sys.exit()

    msg_list = []
    generated_charts = []
    has_data = False
    
    print(f"[{curr_time}] NOC 終極戰情室 v10.0 (全模組整合企業版) 執行中...")
    
    is_bull_market, market_msg = get_market_regime()
    noc_state = load_state()
    msg_list.append(f"🌐 【大盤風向】: {market_msg}\n")

    # === 戰區 1：庫藏股 ===
    if MY_PORTFOLIO:
        msg_list.append("━━━━━━━━━━━━━━\n💼 【庫藏股 (實體持股動態防禦)】\n━━━━━━━━━━━━━━\n")
        for sym, data in MY_PORTFOLIO.items():
            hist = get_stock_data(sym, data['name'])
            if hist is None: continue
                
            td = hist.iloc[-1]
            has_data = True
            
            # 庫藏股固定畫圖
            chart_file = draw_chart_if_needed(hist, sym)
            generated_charts.append(chart_file)
            
            curr_price = td['Close']
            atr = td['ATR']
            buy_price = data['buy_price']
            roi_pct = ((curr_price - buy_price) / buy_price) * 100
            
            etf_icon, bias_limit, etf_desc = get_etf_strategy(sym, data['name'])
            is_etf = "一般型" not in etf_icon  
            
            sym_state = noc_state.get(sym, {"status": "NONE"})
            
            if is_etf:
                if sym_state["status"] != "REAL_HOLD_ETF":
                    noc_state[sym] = {"status": "REAL_HOLD_ETF", "entry": buy_price}
                
                if roi_pct <= -10.0: pnl_alert = f"💎【黃金坑加碼】帳面回檔 {roi_pct:.2f}%，啟動大額建倉！"
                elif roi_pct <= -5.0: pnl_alert = f"📉【紀律扣款】帳面回檔 {roi_pct:.2f}%，維持定期定額。"
                else: pnl_alert = f"🧘‍♂️【長線鎖籌】無懼波動，靜待資產翻倍。"
            else:
                stop_distance = atr * ATR_MULTIPLIER
                
                if sym_state["status"] != "REAL_HOLD":
                    noc_state[sym] = {"status": "REAL_HOLD", "entry": buy_price, "trailing_stop": curr_price - stop_distance}
                    sym_state = noc_state[sym]
                    
                final_stop = max(sym_state["trailing_stop"], curr_price - stop_distance)
                
                if curr_price < final_stop: pnl_alert = f"🩸【拔線警戒】跌破防守線 {final_stop:.1f}，請嚴格執行離場！"
                else:
                    noc_state[sym]["trailing_stop"] = final_stop 
                    if roi_pct > 0: pnl_alert = f"🔥 獲利巡航 | 📍 防線墊高至: {final_stop:.1f}"
                    else: pnl_alert = f"🟡 浮虧防禦 | 📍 死守底線: {final_stop:.1f}"
            
            portfolio_msg = f"{etf_icon} {data['name']} ({sym})\n"
            portfolio_msg += f"   成本: {buy_price:.2f} | 股數: {data['shares']} | 現價: {curr_price:.2f}\n"
            portfolio_msg += f"   損益: {roi_pct:+.2f}% | 👉 指令: {pnl_alert}\n\n"
            msg_list.append(portfolio_msg)

    # === Trello & 本地雷達 觀察網域分流 ===
    for cat, stocks in STOCK_DICT.items():
        if not stocks: continue 
        
        is_etf_zone = "ETF" in cat.upper()
        is_radar_zone = "雷達" in cat
        is_key_obs = "重點觀測" in cat
        is_normal_obs = "觀察" in cat and not is_key_obs and not is_radar_zone
        
        cat_msg_list = [] 
        
        for sym, item in stocks.items():
            # 解耦處理：判斷來源是 Trello 字典還是 JSON 字串
            if isinstance(item, dict):
                name = item.get("name", sym)
                tips = item.get("trello_tip", "")
            else:
                name = item
                tips = ""

            hist = get_stock_data(sym, name)
            if hist is None: continue
                
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
            
            est_vol = td.get('Est_Volume', td['Volume'])
            pos = td['Price_Position'] if 'Price_Position' in td and not pd.isna(td['Price_Position']) else 0.5
            trust_streak = int(td['Trust_Streak']) if 'Trust_Streak' in td and not pd.isna(td['Trust_Streak']) else 0
            bias = ((close - ma20) / ma20) * 100 if ma20 else 0
            
            if est_vol > vma5 * 1.2: vol_status = "📈 出量"
            elif est_vol < vma5 * 0.8: vol_status = "📉 量縮"
            else: vol_status = "➖ 量平"
                
            if close > ma5 > ma20: trend_status = "🔥 多頭"
            elif close < ma5 < ma20: trend_status = "🧊 空頭"
            else: trend_status = "🔄 盤整"
            
            yoy = get_revenue_yoy(sym)
            yoy_label = f"{yoy:.2f}%" if isinstance(yoy, float) else yoy
            if isinstance(yoy, float) and yoy >= YOY_EXPLOSION_PCT: yoy_label += " (🌟 業績大爆發)"

            kd_str = f"K:{k:.1f} D:{d:.1f}"
            if k < 30 and k > d and hist['K'].iloc[-2] <= hist['D'].iloc[-2]: kd_str += " (🌟 KD金叉)"
            elif k > 80: kd_str += " (⚠️ 短線過熱)"

            pe_str = f"{pe:.1f}" if isinstance(pe, float) else pe
            is_overvalued = isinstance(pe, float) and pe > PE_LIMIT

            chip_msg = td['Chip_Status']
            if trust_streak > 0: chip_msg += f" (連買 {trust_streak} 天)"
            elif trust_streak < 0: chip_msg += f" (連賣 {abs(trust_streak)} 天)"

            predict_msg = "無特殊徵兆"
            if est_vol > vma5 * 2: 
                if pos > 0.7: predict_msg = "💀【動能竭盡】高檔爆量轉折！"
                elif pos < 0.3: predict_msg = "🔥【底部換手】低檔爆量，醞釀反彈！"
                else: predict_msg = "⚠️【中繼爆量】留意方向表態！"
            elif td['Shadow_Ratio'] > 0.5 and est_vol > vma5 * 1.5: 
                if pos > 0.7: predict_msg = "⚠️【避雷針陷阱】高檔長上影線！"
                elif pos < 0.3: predict_msg = "🌟【仙人指路】低檔長上影線試盤！"
            elif close > td['20_High'] and est_vol > vma5 * 1.2: predict_msg = "🚀【無壓巡航】突破 20 日高！"
            elif td['BB_Width'] < 0.08: predict_msg = "⚠️【大變盤預警】通道極度壓縮！"
            
            stop_distance = atr * ATR_MULTIPLIER
            safe_stop_distance = stop_distance if stop_distance > 0 else 999999
            suggested_shares = min(math.floor((TOTAL_CAPITAL * RISK_PER_TRADE) / safe_stop_distance), math.floor(TOTAL_CAPITAL / (close if close > 0 else 1.0)))
            
            sym_state = noc_state.get(sym, {"status": "NONE"})
            alert = "✅ 持股觀望"
            
            if sym_state["status"] == "REAL_HOLD": alert = f"💼 已列入持股防禦區 | 📍 防線: {sym_state['trailing_stop']:.1f}"
            elif sym_state["status"] == "NONE":
                if td['Sniper_Signal']: 
                    if not is_bull_market: alert = "🛡️【大盤攔截】大盤偏空，放棄狙擊。"
                    elif isinstance(yoy, float) and yoy < 0: alert = "🛡️【基本面攔截】營收衰退，避開地雷。"
                    elif is_overvalued: alert = f"🛡️【估值攔截】PE {pe_str} 過高，風險極大。"
                    else:
                        stop_price = close - stop_distance
                        noc_state[sym] = {"status": "HOLD", "entry": close, "trailing_stop": stop_price}
                        alert = f"{'⚔️【雙劍合璧】' if isinstance(yoy, float) and yoy >= YOY_EXPLOSION_PCT else '🚀【啟動狙擊】'}買入 {suggested_shares/1000:.1f} 張，停損 {stop_price:.1f}"
                elif td['Sniper_Memory_5D'] == 1: 
                    if close > ma5: alert = "🔥【狙擊延續】站穩5日線！" 
                    else: alert = "⚠️【狙擊失效】跌破5日線！"
            elif sym_state["status"] == "HOLD":
                new_stop = max(sym_state["trailing_stop"], close - stop_distance)
                if close < new_stop: 
                    alert = f"🩸【拔線離場】跌破防守線 {new_stop:.1f}！"
                    noc_state[sym] = {"status": "NONE"}
                else: 
                    noc_state[sym]["trailing_stop"] = new_stop
                    alert = f"🔥【波段抱牢】防守線: {new_stop:.1f}"

            write_noc_log(curr_date, sym, name, close, rsi, vol_status, trend_status, predict_msg, chip_msg, alert)

            # --- 戰區 2：ETF 專區 ---
            if is_etf_zone:
                etf_type, bias_limit, etf_desc = get_etf_strategy(sym, name)
                if bias > bias_limit: etf_cmd = "⚠️ 乖離過熱，建議分批獲利了結"
                elif k < 30 and k > d: etf_cmd = "🔥 KD低檔金叉，建議佈局買進"
                elif close > ma5: etf_cmd = "✅ 趨勢向上，續抱"
                else: etf_cmd = "⏳ 趨勢偏弱，觀望"
                
                stock_msg = f"{etf_type} {name} ({sym})\n"
                stock_msg += f"   現價: {close:.2f} | 乖離: {bias:+.1f}% ({'🚨過熱' if bias > bias_limit else '✅穩定'})\n"
                stock_msg += f"   屬性: {etf_desc}\n"
                stock_msg += f"   👉 指令: {etf_cmd}\n"
                if tips: stock_msg += f"   💡 戰略提示: {tips}\n"
                stock_msg += "\n"
                cat_msg_list.append(stock_msg)
                
                chart_file = draw_chart_if_needed(hist, sym)
                if chart_file not in generated_charts: generated_charts.append(chart_file)

            # --- 戰區 5：🎯 雷達鎖定區 ---
            elif is_radar_zone:
                stock_msg = f"🎯 {name} ({sym})\n"
                stock_msg += f"   現價: {close:.2f} | 狀態: {trend_status} | {vol_status}\n"
                stock_msg += f"   指標: {kd_str} | RSI: {rsi:.1f}\n"
                stock_msg += f"   💰 籌碼: {chip_msg}\n"
                stock_msg += f"   👉 指令: {alert}\n"
                if tips: stock_msg += f"   💡 戰略提示: {tips}\n"
                stock_msg += "\n"
                cat_msg_list.append(stock_msg)
                
                chart_file = draw_chart_if_needed(hist, sym)
                if chart_file not in generated_charts: generated_charts.append(chart_file)

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
                if tips: stock_msg += f"   💡 戰略提示: {tips}\n"
                stock_msg += "\n"
                cat_msg_list.append(stock_msg)
                
                chart_file = draw_chart_if_needed(hist, sym)
                if chart_file not in generated_charts: generated_charts.append(chart_file)

            # --- 戰區 4：觀察區 (條件觸發才渲染與顯示) ---
            elif is_normal_obs:
                is_2560 = td.get('Signal_2560', False)
                is_trap = predict_msg in ["💀【動能竭盡】高檔爆量轉折！", "⚠️【避雷針陷阱】高檔長上影線！", "⚠️【大變盤預警】通道極度壓縮！"]
                is_recovery = td['Sniper_Signal'] or (k < 30 and k > d) or ("止跌" in tips or "支撐" in tips) or predict_msg in ["🔥【底部換手】低檔爆量，醞釀反彈！", "🌟【仙人指路】低檔長上影線試盤！"] or is_2560
                
                if is_2560:
                    predict_msg = "🎯【2560戰法】量縮回踩 25MA，絕佳左側佈局點！"
                    alert = "✅ 準備進場 (請留意停損設 25MA 下方 3%)"

                if is_trap or is_recovery:
                    # 🌟 觸發警報才花費算力繪圖
                    chart_file = draw_chart_if_needed(hist, sym)
                    if chart_file not in generated_charts: generated_charts.append(chart_file)
                    
                    stock_msg = f"👀 {name} ({sym})\n"
                    stock_msg += f"   現價: {close:.2f} | RSI: {rsi:.1f} | 乖離: {bias:+.1f}%\n"
                    stock_msg += f"   💰 籌碼: {chip_msg}\n"
                    
                    if is_2560: stock_msg += f"   🎯 條件觸發: 🌟 高勝率回踩狙擊\n"
                    else: stock_msg += f"   🎯 條件觸發: {'🚨 陷阱預警' if is_trap else '🔥 復甦/狙擊訊號'}\n"
                        
                    stock_msg += f"   👉 預判/指令: {predict_msg if is_trap else alert}\n"
                    if tips: stock_msg += f"   💡 戰略提示: {tips}\n"
                    stock_msg += "\n"
                    cat_msg_list.append(stock_msg)
                else:
                    continue
                    
            # --- 其他未分類戰區 ---
            else:
                stock_msg = f"🔸 {name} ({sym})\n"
                stock_msg += f"   現價: {close:.2f} | 狀態: {trend_status}\n"
                stock_msg += f"   👉 指令: {alert}\n"
                if tips: stock_msg += f"   💡 戰略提示: {tips}\n"
                stock_msg += "\n"
                cat_msg_list.append(stock_msg)
                
                chart_file = draw_chart_if_needed(hist, sym)
                if chart_file not in generated_charts: generated_charts.append(chart_file)

        if cat_msg_list:
            msg_list.append(f"━━━━━━━━━━━━━━\n📂 【{cat}】\n━━━━━━━━━━━━━━\n")
            msg_list.extend(cat_msg_list)

    # =========================================================
    # === 9. 🏆 ETF 雙引擎績效競技場 (高速快取版) ===
    # =========================================================
    etf_arena = {"💰高股息防禦組": [], "🚀市值與主題成長組": []}
    current_year = curr_date.year

    all_etfs = {}
    if MY_PORTFOLIO:
        for sym, data in MY_PORTFOLIO.items():
            all_etfs[sym] = data['name']
    for cat, stocks in STOCK_DICT.items():
        if stocks:
            for sym, item in stocks.items():
                all_etfs[sym] = item.get("name", sym) if isinstance(item, dict) else item

    for sym, name in all_etfs.items():
        etf_icon, _, _ = get_etf_strategy(sym, name)
        is_etf = "一般型" not in etf_icon
        
        if is_etf:
            # ⚡ 秒速從記憶體提取資料
            hist = get_stock_data(sym, name)
            if hist is None or len(hist) < 10: continue
            
            close_price = hist['Close'].iloc[-1]
            qtr_days = min(60, len(hist)-1)
            qtr_price = hist['Close'].iloc[-(qtr_days+1)]
            qtr_roi = ((close_price - qtr_price) / qtr_price) * 100
            
            hist_ytd = hist[hist.index.year == current_year]
            if not hist_ytd.empty:
                ytd_start_price = hist_ytd['Close'].iloc[0]
                ytd_roi = ((close_price - ytd_start_price) / ytd_start_price) * 100
            else:
                ytd_roi = qtr_roi 
                
            group_key = "💰高股息防禦組" if "高股息" in etf_icon else "🚀市值與主題成長組"
            etf_arena[group_key].append({"name": name, "sym": sym, "qtr_roi": qtr_roi, "ytd_roi": ytd_roi})

    arena_msg = []
    if etf_arena["💰高股息防禦組"] or etf_arena["🚀市值與主題成長組"]:
        arena_msg.append("━━━━━━━━━━━━━━\n🏆 【ETF 雙引擎績效競技場 (自動汰弱留強)】\n━━━━━━━━━━━━━━\n")
        
        for group_name, group_data in etf_arena.items():
            if not group_data: continue
            arena_msg.append(f"**{group_name}**\n")
            
            sorted_etfs = sorted(group_data, key=lambda x: x['qtr_roi'], reverse=True)
            medals = ["🥇", "🥈", "🥉"]
            for idx, etf in enumerate(sorted_etfs):
                medal = medals[idx] if idx < 3 else "🔸"
                q_str = f"{etf['qtr_roi']:+.1f}%"
                y_str = f"{etf['ytd_roi']:+.1f}%"
                
                if etf['qtr_roi'] > 5.0 and etf['ytd_roi'] > 10.0: status = "🔥 雙料強勢"
                elif etf['qtr_roi'] < 0 and etf['ytd_roi'] > 0: status = "⏳ 短線洗盤，長線穩健"
                elif etf['qtr_roi'] < -2.0 and etf['ytd_roi'] < 0: status = "⚠️ 嚴重落後，請檢視佔比"
                else: status = "✅ 穩定跟隨"
                    
                arena_msg.append(f"{medal} {etf['name']} ({etf['sym']})\n   季動能 {q_str} ｜ 本年累計 {y_str} ({status})\n")
            arena_msg.append("\n")
            
        msg_list.extend(arena_msg)
    
    # === 最終收尾與發送 ===
    if has_data or len(msg_list) > 0:
        save_state(noc_state) 
        if len(msg_list) == 1 and "大盤風向" in msg_list[0]: 
            msg_list.append("\n🔕 【靜默模式】無觸發條件。")
            
        final_text = f"📡 【NOC 終極戰情室 v10.0 (全模組整合企業版)】\n📅 時間：{curr_time}\n━━━━━━━━━━━━━━\n" + "".join(msg_list)
        send_reports(f"NOC 戰情報告 {curr_date}", final_text, generated_charts)
        
        for chart in generated_charts:
            if os.path.exists(chart): os.remove(chart)
    else:
        print("休市或資料讀取失敗，伺服器待命。")
