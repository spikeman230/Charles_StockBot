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

# === 1.2 🌟 ETF 自動判定模組 ===
def get_etf_config(symbol, name):
    div_keys = ["高股息", "優息", "0056", "00878", "00919", "00929"]
    mkt_keys = ["0050", "006208", "市值", "AAPL", "NVDA", "TSM"]
    if any(k in name or k in symbol for k in div_keys):
        return "💰", 5.0, "高股息控管 (5%預警)"
    elif any(k in name or k in symbol for k in mkt_keys):
        return "🚀", 10.0, "成長型動能 (10%預警)"
    return "🔸", 8.0, "一般防禦區 (8%預警)"

# === 2. Trello 雲端資料庫讀取引擎 (原封不動) ===
def fetch_trello_deployment():
    trello_dict = {}
    my_portfolio = {}
    
    if not TRELLO_KEY or not TRELLO_TOKEN or not TRELLO_BOARD_ID:
        print("⚠️ 未偵測到 Trello 金鑰，將啟用『機房預設』的靜態兵力部署。")
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
                        name = name if name else symbol
                        desc = card.get('desc', '')
                        
                        buy_price = 0.0; shares = 1000
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
                        stock_list[symbol] = name if name else symbol
                    if stock_list: trello_dict[list_name] = stock_list
            return trello_dict, my_portfolio
    except Exception as e:
        print(f"⚠️ Trello 連線失敗: {e}")
    return None, None

TRELLO_DICT, TRELLO_PORTFOLIO = fetch_trello_deployment()
if TRELLO_DICT is not None and TRELLO_PORTFOLIO is not None:
    STOCK_DICT = TRELLO_DICT; MY_PORTFOLIO = TRELLO_PORTFOLIO
else:
    STOCK_DICT = {}; MY_PORTFOLIO = {}

# 雷達名單讀取 (原封不動)
for file, label in [("radar_targets.json", "🎯 雷達鎖定 (新進火種區)"), ("lightning_targets.json", "⚡ 雷達鎖定 (短線飆股區)")]:
    if os.path.exists(file):
        try:
            with open(file, "r", encoding="utf-8") as f:
                data = json.load(f)
                if data: STOCK_DICT[label] = data
        except: pass

# === 3. 狀態記憶與日誌 (原封不動) ===
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
    try:
        with open(log_filename, mode='a', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            if not os.path.exists(log_filename): writer.writerow(["日期", "代號", "名稱", "收盤價", "RSI", "量能狀態", "趨勢狀態", "戰場預判", "籌碼訊號", "行動指令"])
            writer.writerow([date, symbol, name, f"{close_price:.2f}", f"{rsi:.2f}", vol_status, status, predict, chip_signal, alert])
    except: pass

# === 4. 環境感知 (原封不動) ===
def get_market_regime():
    try:
        twii = yf.Ticker("^TWII").history(period="1mo")
        twii['20MA'] = twii['Close'].rolling(20).mean()
        if twii['Close'].iloc[-1] > twii['20MA'].iloc[-1]: return True, "🟢 多頭格局 (站上月線)"
        else: return False, "🔴 空頭警戒 (跌破月線)"
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
    try: return yf.Ticker(symbol).info.get('trailingPE', yf.Ticker(symbol).info.get('forwardPE', "N/A"))
    except: return "N/A"

# === 5. FinMind 籌碼分析 (原封不動) ===
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
    if all(col in hist.columns for col in ['Foreign_Inv', 'Trust_Inv', 'Dealer_Inv']):
        hist['Total_Institutional'] = hist['Foreign_Inv'] + hist['Trust_Inv'] + hist['Dealer_Inv']
        hist['Signal_CoBuy'] = (hist['Foreign_Inv'] > 0) & (hist['Trust_Inv'] > 0)
        hist['Signal_Trust_Trend'] = ((hist['Trust_Inv'] > 0).astype(int).rolling(5).sum() >= 4) & (hist['Trust_Inv'] > 0)
        conds = [(hist['Signal_CoBuy'] == True), (hist['Signal_Trust_Trend'] == True), (hist['Total_Institutional'] > 0)]
        hist['Chip_Status'] = np.select(conds, ["🤝 土洋齊買", "🏦 投信作帳(連買)", "📈 法人偏多"], default="➖ 中性/偏空")
    return hist

def get_strategy_tips(symbol, current_price, k_value, ma5, ma20):
    if symbol == "9933.TW": return "🔥【NOC 訊號】中鼎疑似止跌！" if current_price > ma5 and k_value < 30 else "⏳【NOC 監控】中鼎尚未止跌。"
    if symbol == "6415.TW": return "💎【NOC 訊號】矽力進入支撐區" if 240 <= current_price <= 260 else "🦅【NOC 監控】矽力耐心等待回測。"
    if symbol == "2303.TW": return "🚀【NOC 訊號】聯電強勢！" if current_price > ma5 else "⚠️【NOC 警訊】聯電轉弱。"
    return ""

# === 6. 核心分析引擎 (🌟 重點：指標計算大一統) ===
def get_analysis_and_chart(symbol, name):
    try:
        stock = yf.Ticker(symbol)
        hist = stock.history(period="8mo").dropna(subset=['Close'])
        if len(hist) < 40: return None
            
        hist['Date_Key'] = hist.index.date
        if FINMIND_TOKEN and (".TW" in symbol or ".TWO" in symbol):
            chip_df = get_finmind_chip_data(symbol, (datetime.datetime.now() - datetime.timedelta(days=200)).strftime("%Y-%m-%d"))
            if not chip_df.empty:
                hist = hist.merge(chip_df, left_on='Date_Key', right_index=True, how='left').fillna({'Foreign_Inv': 0, 'Trust_Inv': 0, 'Dealer_Inv': 0})
        hist = calculate_chip_signals(hist)
        
        # 集中計算所有指標
        hist['5MA'] = hist['Close'].rolling(5).mean(); hist['20MA'] = hist['Close'].rolling(20).mean()
        hist['5VMA'] = hist['Volume'].rolling(5).mean()
        
        l9 = hist['Low'].rolling(9).min(); h9 = hist['High'].rolling(9).max()
        hist['K'] = ((hist['Close'] - l9) / (h9 - l9) * 100).ewm(com=2, adjust=False).mean()
        hist['D'] = hist['K'].ewm(com=2, adjust=False).mean()
        
        delta = hist['Close'].diff(); gain = delta.clip(lower=0); loss = -delta.clip(upper=0)
        hist['RSI'] = (100 - (100 / (1 + (gain.ewm(com=13, adjust=False).mean() / loss.ewm(com=13, adjust=False).mean())))).fillna(50)
        
        hist['ATR'] = pd.concat([hist['High'] - hist['Low'], abs(hist['High'] - hist['Close'].shift(1)), abs(hist['Low'] - hist['Close'].shift(1))], axis=1).max(axis=1).rolling(14).mean()
        
        hist['MACD'] = hist['Close'].ewm(span=12, adjust=False).mean() - hist['Close'].ewm(span=26, adjust=False).mean()
        hist['MACD_Hist'] = hist['MACD'] - hist['MACD'].ewm(span=9, adjust=False).mean()
        
        hist['STD20'] = hist['Close'].rolling(20).std()
        hist['BB_Width'] = (4 * hist['STD20']) / hist['20MA']
        
        hist['Is_Bottoming'] = ((hist['Close'] < hist['5MA']) & (hist['MACD_Hist'].shift(2) < hist['MACD_Hist'].shift(1)) & (hist['MACD_Hist'].shift(1) < hist['MACD_Hist']) & (hist['MACD_Hist'] < 0)).astype(int)
        hist['Sniper_Signal'] = hist['Is_Bottoming'].rolling(3).max().fillna(0).astype(bool) & ((hist['Close'].shift(1) < hist['5MA'].shift(1)) & (hist['Close'] > hist['5MA']) & (hist['Volume'] > hist['5VMA'] * 1.2))
        hist['Sniper_Memory_5D'] = hist['Sniper_Signal'].rolling(5).max().fillna(0)
        
        hist['20_High'] = hist['High'].rolling(20).max().shift(1)
        hist['Shadow_Ratio'] = (hist['High'] - hist[['Open', 'Close']].max(axis=1)) / (hist['High'] - hist['Low']).replace(0, 0.001)
        
        chart_file = f"{symbol}_chart.png"
        try:
            mc = mpf.make_marketcolors(up='red', down='green', edge='black', wick='black', volume='gray')
            s = mpf.make_mpf_style(base_mpf_style='yahoo', marketcolors=mc)
            mpf.plot(hist[-60:], type='candle', style=s, volume=True, mav=(5, 20), savefig=chart_file)
        except: mpf.plot(hist[-60:], type='candle', style='yahoo', volume=True, mav=(5, 20), savefig=chart_file)
            
        return hist, chart_file
    except: return None

# === 7. 發送模組 (原封不動) ===
def send_reports(subject, text_body, chart_files):
    if TG_TOKEN and TG_CHAT_ID:
        try:
            for part in [text_body[i:i+4000] for i in range(0, len(text_body), 4000)]:
                requests.post(f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage", json={"chat_id": TG_CHAT_ID, "text": part, "disable_web_page_preview": True}, timeout=10)
        except: pass
    if EMAIL_USER and EMAIL_PASS and EMAIL_TO:
        try:
            msg = MIMEMultipart(); msg['From'] = EMAIL_USER; msg['To'] = EMAIL_TO; msg['Subject'] = subject
            msg.attach(MIMEText(text_body, 'plain'))
            for c in chart_files:
                if os.path.exists(c): msg.attach(MIMEImage(open(c, 'rb').read(), name=os.path.basename(c)))
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s: s.login(EMAIL_USER, EMAIL_PASS); s.send_message(msg)
        except: pass

# === 8. 主程式執行 (維持原手動串接邏輯，確保無感) ===
if __name__ == "__main__":
    tw_tz = datetime.timezone(datetime.timedelta(hours=8))
    curr_date = datetime.datetime.now(tw_tz).date()
    curr_time = datetime.datetime.now(tw_tz).strftime("%Y-%m-%d %H:%M:%S")
    msg_list = []; generated_charts = []; has_data = False
    
    is_bull_market, market_msg = get_market_regime()
    noc_state = load_state()
    msg_list.append(f"🌐 【大盤風向】: {market_msg}\n")

    # === 庫存機櫃 ===
    if MY_PORTFOLIO:
        msg_list.append("━━━━━━━━━━━━━━\n💼 【庫存機櫃 (真實持股動態防禦)】\n━━━━━━━━━━━━━━\n")
        for sym, data in MY_PORTFOLIO.items():
            res = get_analysis_and_chart(sym, data['name'])
            if not res: continue
            hist, chart_file = res; td = hist.iloc[-1]; has_data = True; generated_charts.append(chart_file)
            
            curr_price = td['Close']; atr = td['ATR']; buy_price = data['buy_price']
            roi_pct = ((curr_price - buy_price) / buy_price) * 100
            stop_distance = atr * ATR_MULTIPLIER
            sym_state = noc_state.get(sym, {"status": "NONE"})
            
            if sym_state["status"] != "REAL_HOLD":
                noc_state[sym] = {"status": "REAL_HOLD", "entry": buy_price, "trailing_stop": curr_price - stop_distance}
                sym_state = noc_state[sym]
                
            final_stop = max(sym_state["trailing_stop"], curr_price - stop_distance)
            if curr_price < final_stop: pnl_alert = f"🩸【拔線警戒】跌破防守線 {final_stop:.1f}！"
            else:
                noc_state[sym]["trailing_stop"] = final_stop 
                pnl_alert = f"🔥 獲利巡航 | 📍 防線: {final_stop:.1f}" if roi_pct > 0 else f"🟡 浮虧中 | 📍 防線: {final_stop:.1f}"
                    
            etf_icon, _, _ = get_etf_config(sym, data['name'])
            
            portfolio_msg = f"{etf_icon} {data['name']} ({sym})\n   成本: {buy_price:.2f} | 股數: {data['shares']} | 現價: {curr_price:.2f}\n"
            portfolio_msg += f"   損益: {roi_pct:+.2f}% | 👉 {pnl_alert}\n\n"
            msg_list.append(portfolio_msg)

    # === 觀察網域 ===
    for cat, stocks in STOCK_DICT.items():
        if not stocks: continue 
        cat_printed = False 
        
        for sym, name in stocks.items():
            res = get_analysis_and_chart(sym, name)
            if not res: continue
            hist, chart_file = res; td = hist.iloc[-1]; has_data = True
                
            close = td['Close']; atr = td['ATR']; rsi = td['RSI']
            vma5 = td['5VMA']; ma5 = td['5MA']; ma20 = td['20MA']
            k = td['K']; d = td['D']; pe = get_pe_ratio(sym)
            
            # 🌟 ETF 專屬判定與乖離計算
            etf_icon, bias_limit, etf_desc = get_etf_config(sym, name)
            bias = ((close - ma20) / ma20) * 100 if ma20 else 0
            bias_alert = "🚨過熱" if bias > bias_limit else "✅穩定"
            
            vol_status = "📈 出量" if td['Volume'] > vma5 * 1.2 else "📉 量縮" if td['Volume'] < vma5 * 0.8 else "➖ 量平"
            trend_status = "🔥 多頭" if close > ma5 > ma20 else "🧊 空頭" if close < ma5 < ma20 else "🔄 盤整"
            
            yoy = get_revenue_yoy(sym)
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
            safe_stop_distance = stop_distance if stop_distance > 0 else 999999
            safe_close = close if close > 0 else 1.0
            suggested_shares = min(math.floor((TOTAL_CAPITAL * RISK_PER_TRADE) / safe_stop_distance), math.floor(TOTAL_CAPITAL / safe_close))
            
            sym_state = noc_state.get(sym, {"status": "NONE"})
            alert = "✅ 持股觀望"
            
            if sym_state["status"] == "REAL_HOLD": 
                alert = f"💼 已列入持股防禦區 | 📍 防線: {sym_state['trailing_stop']:.1f}"
            elif sym_state["status"] == "NONE":
                if td['Sniper_Signal']: 
                    if not is_bull_market: alert = "🛡️【大盤攔截】大盤偏空，放棄狙擊。"
                    elif isinstance(yoy, float) and yoy < 0: alert = "🛡️【基本面攔截】營收衰退，避開地雷。"
                    elif is_overvalued: alert = f"🛡️【估值攔截】PE {pe_str} 過高，風險極大。"
                    elif bias > bias_limit: alert = f"🛡️【乖離攔截】乖離 {bias:+.1f}% 已過熱，取消狙擊。"
                    else:
                        stop_price = close - stop_distance
                        noc_state[sym] = {"status": "HOLD", "entry": close, "trailing_stop": stop_price}
                        alert = f"{'⚔️【雙劍合璧：終極狙擊】' if is_yoy_explosion else '🚀【啟動狙擊】'}建議買入 {suggested_shares/1000:.1f} 張，停損 {stop_price:.1f}"
                elif td['Sniper_Memory_5D'] == 1: 
                    alert = "🔥【狙擊延續】站穩5日線！" if close > ma5 else "⚠️【狙擊失效】跌破5日線！"
            elif sym_state["status"] == "HOLD":
                new_stop = max(sym_state["trailing_stop"], close - stop_distance)
                if close < new_stop: 
                    alert = f"🩸【拔線離場】跌破防守線 {new_stop:.1f}！"
                    noc_state[sym] = {"status": "NONE"}
                else: 
                    noc_state[sym]["trailing_stop"] = new_stop
                    alert = f"🔥【波段抱牢】損益: {((close - sym_state['entry']) / sym_state['entry']) * 100:+.2f}% | 防守線: {new_stop:.1f}"

            write_noc_log(curr_date, sym, name, close, rsi, vol_status, trend_status, predict_msg, td['Chip_Status'], alert)
            tips = get_strategy_tips(sym, close, k, ma5, ma20)
            
            is_notable = alert != "✅ 持股觀望" or predict_msg != "無特殊徵兆" or tips != "" or td['Chip_Status'] in ["🤝 土洋齊買", "🏦 投信作帳(連買)"] or bias > bias_limit
            
            if (not SILENT_MODE) or is_notable:
                if not cat_printed: 
                    msg_list.append(f"━━━━━━━━━━━━━━\n📂 【{cat}】\n━━━━━━━━━━━━━━\n")
                    cat_printed = True
                
                if chart_file not in generated_charts: generated_charts.append(chart_file)
                    
                stock_msg = f"{etf_icon} {name} ({sym})\n"
                stock_msg += f"   現價: {close:.2f} | 乖離: {bias:+.1f}% ({bias_alert}) | PE: {pe_str}\n"
                stock_msg += f"   指標: {kd_str} | RSI: {rsi:.1f} | 類型: {etf_desc}\n"
                stock_msg += f"   狀態: {trend_status} | {vol_status} | YoY: {yoy_label}\n"
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
            msg_list.append("\n🔕 【靜默模式】目前觀察網域無特殊狙擊訊號，詳細數據請參閱 CSV 日誌。")
            
        final_text = f"📡 NOC 終極戰情室 v8.7 (核心封裝版)\n📅 {curr_time}\n━━━━━━━━━━━━━━\n" + "".join(msg_list)
        send_reports(f"NOC 戰報", final_text, generated_charts)
        for chart in generated_charts:
            if os.path.exists(chart): os.remove(chart)
