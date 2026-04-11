import yfinance as yf
import requests
import os
import datetime
import pandas as pd
import numpy as np
import csv
import json
import math
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

# === 1.1 量化基金風控參數 (v7.6 全面動態核心) ===
TOTAL_CAPITAL = 1000000  # 預設總資金 100 萬台幣
RISK_PER_TRADE = 0.02    # 單筆風險 2%
ATR_MULTIPLIER = 2.0     # 2倍 ATR 動態停損 (套用於全部持股)

# === 2. 專屬通訊錄 (外部觀察網域) ===
STOCK_DICT = {
    "🛡️ 核心持股 (重倉伺服器)": {"3037.TW": "欣興 (ABF載板)"},
    "🔥 潛力種子 (高頻寬觀察區)": {"3163.TWO": "波若威", "5388.TW": "中磊", "3714.TW": "富采","2337.TW": "旺宏"},
    "👀 常態觀察區 (例行監控節點)": {"2330.TW": "台積電", "0050.TW": "元大台灣50","AAPL": "蘋果","NVDA": "輝達"},
    "💾 記憶體族群 (美光連動網域)": {"MU": "美光", "2408.TW": "南亞科", "3260.TWO": "威剛", "8299.TWO": "群聯" },
    "🔍 YAHOO 觀察區": {"2027.TW": "大成鋼", "2382.TW": "廣達", "2886.TW": "兆豐金", "2409.TW": "友達", "2352.TW": "佳世達","2317.TW": "鴻海", "6116.TW": "彩晶" },
    "真實持股 追蹤區" : {"8431.TWO":"匯鑽科","3231.TW":"緯創" }
}

# === 2.1 真實持股庫存 (實體機房配置) ===
MY_PORTFOLIO = {
    "3231.TW": {"name": "緯創", "buy_price": 130.5, "shares": 1000},
    "8431.TWO": {"name": "匯鑽科", "buy_price": 70.7, "shares": 1000},
    "6116.TW": {"name": "彩晶", "buy_price": 8.4, "shares": 1000},
    "2317.TW": {"name": "鴻海", "buy_price": 201.5, "shares": 1000}
}

# === 3. 實體狀態記憶庫與日誌 (Stateful Database) ===
STATE_FILE = "noc_state.json"

def load_state():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f: return json.load(f)
        except: pass
    return {}

def save_state(state_data):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state_data, f, ensure_ascii=False, indent=4)

def write_noc_log(date, symbol, name, close_price, rsi, vol_status, status, alert, predict, chip_signal):
    log_filename = "noc_trading_log.csv"
    file_exists = os.path.exists(log_filename)
    with open(log_filename, mode='a', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["日期", "代號", "名稱", "收盤價", "RSI", "量能狀態", "趨勢狀態", "戰場預判", "籌碼訊號", "行動指令"])
        writer.writerow([date, symbol, name, f"{close_price:.2f}", f"{rsi:.2f}", vol_status, status, predict, chip_signal, alert])

# === 4. 環境感知：大盤與營收 YoY ===
def get_market_regime():
    try:
        twii = yf.Ticker("^TWII").history(period="1mo")
        twii['20MA'] = twii['Close'].rolling(20).mean()
        if twii['Close'].iloc[-1] > twii['20MA'].iloc[-1]: return True, f"🟢 多頭格局 (站上月線)"
        else: return False, f"🔴 空頭警戒 (跌破月線)"
    except: return True, "🟡 大盤狀態未知"

def get_revenue_yoy(symbol):
    if not FINMIND_TOKEN: return "N/A"
    fm_symbol = symbol.replace(".TW", "").replace(".TWO", "")
    if not fm_symbol.isdigit(): return "N/A"
    try:
        url = "https://api.finmindtrade.com/api/v4/data"
        params = {"dataset": "TaiwanStockMonthRevenue", "data_id": fm_symbol, "start_date": (datetime.datetime.now() - datetime.timedelta(days=90)).strftime("%Y-%m-%d"), "token": FINMIND_TOKEN}
        r = requests.get(url, params=params, timeout=10)
        data = r.json()
        if data.get("msg") == "success" and len(data.get("data", [])) > 0:
            return float(pd.DataFrame(data["data"]).iloc[-1]['revenue_YearOverYear_ratio'])
    except: pass
    return "N/A"

# === 5. FinMind 籌碼分析 ===
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
        conditions = [(hist['Signal_CoBuy'] == True), (hist['Signal_Trust_Trend'] == True), (hist['Total_Institutional'] > 0)]
        choices = ["🤝 土洋齊買", "🏦 投信作帳(連買)", "📈 法人偏多"]
        hist['Chip_Status'] = np.select(conditions, choices, default="➖ 中性/偏空")
    return hist

# === 6. 核心分析引擎 ===
def get_analysis_and_chart(symbol, name):
    try:
        stock = yf.Ticker(symbol)
        hist = stock.history(period="8mo")
        if len(hist) < 40: return None

        hist['Date_Key'] = hist.index.date
        if FINMIND_TOKEN and (".TW" in symbol or ".TWO" in symbol):
            start_date_str = (datetime.datetime.now() - datetime.timedelta(days=200)).strftime("%Y-%m-%d")
            chip_df = get_finmind_chip_data(symbol, start_date_str)
            if not chip_df.empty:
                hist = hist.merge(chip_df, left_on='Date_Key', right_index=True, how='left').fillna({'Foreign_Inv': 0, 'Trust_Inv': 0, 'Dealer_Inv': 0})
        hist = calculate_chip_signals(hist)

        # 基礎指標與 RSI
        hist['5MA'] = hist['Close'].rolling(window=5).mean()
        hist['20MA'] = hist['Close'].rolling(window=20).mean()
        hist['5VMA'] = hist['Volume'].rolling(window=5).mean()
        delta = hist['Close'].diff()
        gain = delta.clip(lower=0); loss = -1 * delta.clip(upper=0)
        hist['RSI'] = 100 - (100 / (1 + (gain.ewm(com=13, adjust=False).mean() / loss.ewm(com=13, adjust=False).mean())))
        hist['RSI'] = hist['RSI'].fillna(50)

        # ATR 動態波動率 (v7.6)
        hist['H-L'] = hist['High'] - hist['Low']
        hist['H-PC'] = abs(hist['High'] - hist['Close'].shift(1))
        hist['L-PC'] = abs(hist['Low'] - hist['Close'].shift(1))
        hist['TR'] = hist[['H-L', 'H-PC', 'L-PC']].max(axis=1)
        hist['ATR'] = hist['TR'].rolling(window=14).mean()

        # 底部分析與狙擊記憶
        hist['MACD'] = hist['Close'].ewm(span=12, adjust=False).mean() - hist['Close'].ewm(span=26, adjust=False).mean()
        hist['Signal'] = hist['MACD'].ewm(span=9, adjust=False).mean()
        hist['MACD_Hist'] = hist['MACD'] - hist['Signal']
        hist['STD20'] = hist['Close'].rolling(window=20).std()
        hist['BB_Width'] = (4 * hist['STD20']) / hist['20MA']
        
        hist['Is_Bottoming'] = ((hist['Close'] < hist['5MA']) & (hist['MACD_Hist'].shift(2) < hist['MACD_Hist'].shift(1)) & (hist['MACD_Hist'].shift(1) < hist['MACD_Hist']) & (hist['MACD_Hist'] < 0)).astype(int)
        hist['Recent_Bottoming'] = hist['Is_Bottoming'].rolling(window=3).max().fillna(0).astype(bool)
        hist['Is_Breakout'] = (hist['Close'].shift(1) < hist['5MA'].shift(1)) & (hist['Close'] > hist['5MA']) & (hist['Volume'] > hist['5VMA'] * 1.2)
        hist['Sniper_Signal'] = hist['Recent_Bottoming'] & hist['Is_Breakout']
        hist['Sniper_Memory_5D'] = hist['Sniper_Signal'].rolling(window=5).max().fillna(0)

        # 進階特徵工程 (IPS 防禦)
        hist['20_High'] = hist['High'].rolling(window=20).max().shift(1)
        hist['Body_Top'] = hist[['Open', 'Close']].max(axis=1)
        hist['Upper_Shadow'] = hist['High'] - hist['Body_Top']
        hist['K_Length'] = (hist['High'] - hist['Low']).replace(0, 0.001) 
        hist['Shadow_Ratio'] = hist['Upper_Shadow'] / hist['K_Length']

        chart_file = f"{symbol}_chart.png"
        try:
            mc = mpf.make_marketcolors(up='red', down='green', edge='black', wick='black', volume='gray')
            tw_style = mpf.make_mpf_style(base_style='yahoo', marketcolors=mc)
            mpf.plot(hist[-60:], type='candle', style=tw_style, volume=True, mav=(5, 20), title=f"Stock: {symbol}", savefig=chart_file)
        except:
            mpf.plot(hist[-60:], type='candle', style='yahoo', volume=True, mav=(5, 20), title=f"Stock: {symbol}", savefig=chart_file)

        return hist, chart_file
    except Exception as e:
        print(f"[{symbol}] 分析發生錯誤: {e}")
        return None

# === 7. 發送模組 ===
def send_reports(subject, text_body, chart_files):
    if TG_TOKEN and TG_CHAT_ID:
        requests.post(f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage", json={"chat_id": TG_CHAT_ID, "text": text_body, "disable_web_page_preview": True})
    
    if EMAIL_USER and EMAIL_PASS and EMAIL_TO:
        try:
            msg = MIMEMultipart(); msg['From'] = EMAIL_USER; msg['To'] = EMAIL_TO; msg['Subject'] = subject
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
    msg_list = []; generated_charts = []; has_data = False

    print(f"[{curr_time}] NOC 終極融合版 (v7.6 動態持股防禦) 啟動...")

    # 🌐 載入系統大腦與環境
    is_bull_market, market_msg = get_market_regime()
    noc_state = load_state()
    msg_list.append(f"🌐 【大盤風向】: {market_msg}\n")

    # === 處理真實持股 (MY_PORTFOLIO 導入 ATR 動態防禦) ===
    if MY_PORTFOLIO:
        msg_list.append("━━━━━━━━━━━━━━\n💼 【庫存機櫃 (真實持股動態防禦)】\n━━━━━━━━━━━━━━\n")
        for sym, data in MY_PORTFOLIO.items():
            res = get_analysis_and_chart(sym, data['name'])
            if not res: continue
            hist, chart_file = res; td = hist.iloc[-1]
            has_data = True
            generated_charts.append(chart_file)
            
            curr_price = td['Close']
            atr = td['ATR']
            buy_price = data['buy_price']
            roi_pct = ((curr_price - buy_price) / buy_price) * 100
            
            # 🧠 真實持股的 JSON 狀態機 (ATR 動態防線)
            stop_distance = atr * ATR_MULTIPLIER
            sym_state = noc_state.get(sym, {"status": "NONE"})
            
            # 第一次抓到真實持股，寫入初始防守線
            if sym_state["status"] != "REAL_HOLD":
                initial_stop = curr_price - stop_distance
                noc_state[sym] = {"status": "REAL_HOLD", "entry": buy_price, "trailing_stop": initial_stop}
                sym_state = noc_state[sym]
            
            # 計算只進不退的防線
            new_stop = curr_price - stop_distance
            final_stop = max(sym_state["trailing_stop"], new_stop)
            
            # 判定是否跌破
            if curr_price < final_stop:
                pnl_alert = f"🩸【拔線警戒】跌破動態防守線 {final_stop:.1f}，請嚴格執行停利/停損！"
                # 不清除狀態，讓使用者自行決定是否從 MY_PORTFOLIO 刪除
            else:
                noc_state[sym]["trailing_stop"] = final_stop  # 更新防守線
                if roi_pct > 0: pnl_alert = f"🔥 獲利巡航中 | 📍 動態防線墊高至: {final_stop:.1f}"
                else: pnl_alert = f"🟡 暫時浮虧中 | 📍 死守底線: {final_stop:.1f}"
                
            portfolio_msg = f"🔸 {data['name']} ({sym})\n"
            portfolio_msg += f"   成本: {buy_price:.2f} | 現價: {curr_price:.2f}\n"
            portfolio_msg += f"   損益: {roi_pct:+.2f}% | 👉 {pnl_alert}\n\n"
            msg_list.append(portfolio_msg)

    # === 處理觀察網域 (STOCK_DICT) ===
    for cat, stocks in STOCK_DICT.items():
        if not stocks: continue 
        cat_printed = False 
        
        for sym, name in stocks.items():
            res = get_analysis_and_chart(sym, name)
            if not res: continue
            hist, chart_file = res; td = hist.iloc[-1]
            has_data = True
            if chart_file not in generated_charts: generated_charts.append(chart_file)
            
            if not cat_printed:
                msg_list.append(f"━━━━━━━━━━━━━━\n📂 【{cat}】\n━━━━━━━━━━━━━━\n")
                cat_printed = True

            close = td['Close']; atr = td['ATR']; rsi = td['RSI']
            vol_today = td['Volume']; vma5 = td['5VMA']
            vol_status = "📈 出量" if vol_today > vma5 * 1.2 else "📉 量縮" if vol_today < vma5 * 0.8 else "➖ 量平"
            trend_status = "🔥 多頭" if close > td['5MA'] > td['20MA'] else "🧊 空頭" if close < td['5MA'] < td['20MA'] else "🔄 盤整"
            chip_status = td['Chip_Status']

            # 🛡️ 營收與預判 (IPS)
            yoy = get_revenue_yoy(sym)
            yoy_str = f"{yoy:.2f}%" if isinstance(yoy, float) else yoy
            predict_msg = "無特殊徵兆"
            if vol_today > vma5 * 3 and rsi > 75: predict_msg = "💀【動能竭盡】異常爆量，主力可能倒貨！"
            elif td['Shadow_Ratio'] > 0.5 and vol_today > vma5 * 1.5: predict_msg = "⚠️【避雷針陷阱】高檔爆量長上影線！"
            elif close > td['20_High'] and vol_today > vma5 * 1.2: predict_msg = "🚀【無壓巡航】帶量突破 20 日前高！"
            elif td['BB_Width'] < 0.08: predict_msg = "⚠️【大變盤預警】布林通道極度壓縮！"
            elif td['Is_Bottoming'] == 1: predict_msg = "📈【築底預判】空方動能連續收斂！"

            # 🧠 資金控管與量化大腦邏輯
            stop_distance = atr * ATR_MULTIPLIER
            suggested_shares = min(math.floor((TOTAL_CAPITAL * RISK_PER_TRADE) / stop_distance) if stop_distance > 0 else 0, math.floor(TOTAL_CAPITAL / close))
            sym_state = noc_state.get(sym, {"status": "NONE"})
            alert = "✅ 持股觀望"

            # 確保不會跟 MY_PORTFOLIO 衝突 (如果這檔股票剛好也放在觀察區)
            if sym_state["status"] == "REAL_HOLD":
                alert = f"💼 已列入真實持股防禦區 | 📍 動態防線: {sym_state['trailing_stop']:.1f}"
            elif sym_state["status"] == "NONE":
                if td['Sniper_Signal']: 
                    if not is_bull_market: alert = "🛡️【大盤攔截】大盤偏空，放棄狙擊。"
                    elif isinstance(yoy, float) and yoy < 0: alert = f"🛡️【基本面攔截】營收 YoY {yoy_str}，避開地雷。"
                    else:
                        stop_price = close - stop_distance
                        noc_state[sym] = {"status": "HOLD", "entry": close, "trailing_stop": stop_price}
                        alert = f"🚀【啟動狙擊】建議買入 {suggested_shares/1000:.1f} 張，停損設 {stop_price:.1f}"
                elif td['Sniper_Memory_5D'] == 1 and not td['Sniper_Signal']: # v6.1 記憶追蹤
                    if close > td['5MA']: alert = "🔥【狙擊延續：強勢巡航】近期突破，站穩5日線！"
                    else: alert = "⚠️【狙擊失效：跌破防線】跌破5日線，請觀望！"
                elif rsi > 80: alert = "💰【獲利了結】短線過熱，注意回檔。"
                elif close < td['5MA'] < td['20MA'] and vol_today > vma5 * 1.2: alert = "💀【強制退場】空頭確認，大單砸盤！"
            elif sym_state["status"] == "HOLD":
                new_stop = max(sym_state["trailing_stop"], close - stop_distance)
                if close < new_stop:
                    alert = f"🩸【拔線離場】跌破防守線 {new_stop:.1f}，強制停損/停利！"
                    noc_state[sym] = {"status": "NONE"} 
                else:
                    noc_state[sym]["trailing_stop"] = new_stop
                    alert = f"🔥【波段抱牢】未實現損益: {((close - sym_state['entry']) / sym_state['entry']) * 100:+.2f}% | 防守線: {new_stop:.1f}"

            write_noc_log(curr_date, sym, name, close, rsi, vol_status, trend_status, predict_msg, chip_status, alert)
            
            # 資訊面板
            stock_msg = f"🔸 {name} ({sym})\n   現價: {close:.2f} | RSI: {rsi:.1f} | 營收YoY: {yoy_str}\n   狀態: {trend_status} | {vol_status}\n"
            if chip_status != "無資料": stock_msg += f"   💰 籌碼: {chip_status}\n"
            stock_msg += f"   🔮 預判: {predict_msg}\n   👉 指令: {alert}\n\n"
            msg_list.append(stock_msg)

    if has_data or len(msg_list) > 0:
        save_state(noc_state) # 儲存 JSON 記憶
        final_text = f"📡 【NOC 終極融合版 v7.6】\n📅 時間：{curr_time}\n━━━━━━━━━━━━━━\n" + "".join(msg_list)
        send_reports(f"NOC 戰情報告 {curr_date}", final_text, generated_charts)
        for chart in generated_charts:
            if os.path.exists(chart): os.remove(chart)
    else:
        print("休市或資料讀取失敗，伺服器待命。")
