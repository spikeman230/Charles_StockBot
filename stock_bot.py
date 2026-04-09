import yfinance as yf
import requests
import os
import datetime
import pandas as pd
import numpy as np
import csv
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

# === 2. 專屬通訊錄 (外部觀察網域) ===
STOCK_DICT = {
    "🛡️ 核心持股 (重倉伺服器)": {"3037.TW": "欣興 (ABF載板)"},
    "🔥 潛力種子 (高頻寬觀察區)": {"3163.TW": "波若威", "5388.TW": "中磊", "3714.TW": "富采"},
    "👀 常態觀察區 (例行監控節點)": {"2330.TW": "台積電", "2317.TW": "鴻海", "0050.TW": "元大台灣50"},
    "💾 YAHOO 觀察區": {"2027.TW": "大成鋼", "2382.TW": "廣達", "2886.TW": "兆豐金", "6116.TW": "彩晶", "3231.TW": "緯創","2352.TW": "佳世達", "NVDA": "輝達"}
}

# === 2.1 真實持股庫存 (您的實體機房配置) ===
# 📝 請在這裡填入您的真實買進成本與股數
MY_PORTFOLIO = {
    "3231.TW": {"name": "緯創", "buy_price": 130.5, "shares": 1000},
    "8431.TWO": {"name": "匯鑽科", "buy_price": 70.7, "shares": 1000}
}
# ⚙️ 設定自動停利/停損的閥值 (Threshold)
TAKE_PROFIT_PCT = 20.0  # 當報酬率達到 +20% 時，發出獲利了結警報
STOP_LOSS_PCT = -10.0   # 當報酬率跌至 -10% 時，發出停損拔線警報

# === 3. 🛸 自動拓荒雷達 ===
def scan_top_trust_buy(limit=5):
    if not FINMIND_TOKEN: return {}
    print("📡 啟動全網掃描：尋找投信最新認養目標...")
    for i in range(1, 6):
        target_date = (datetime.datetime.now() - datetime.timedelta(days=i)).strftime("%Y-%m-%d")
        url = "https://api.finmindtrade.com/api/v4/data"
        params = {"dataset": "TaiwanStockInstitutionalInvestorsBuySell", "date": target_date, "token": FINMIND_TOKEN}
        try:
            r = requests.get(url, params=params, timeout=15)
            data = r.json()
            if data.get("msg") == "success" and len(data.get("data", [])) > 0:
                df = pd.DataFrame(data["data"])
                trust_df = df[df['name'].str.contains('投信', na=False)].copy()
                if not trust_df.empty:
                    trust_df['net_buy'] = trust_df['buy'] - trust_df['sell']
                    top_df = trust_df.sort_values(by='net_buy', ascending=False)
                    existing_symbols = [sym.replace('.TW', '').replace('.TWO', '') for stocks in STOCK_DICT.values() for sym in stocks.keys()]
                    radar_stocks = {}
                    count = 0
                    for _, row in top_df.iterrows():
                        stock_id = str(row['stock_id'])
                        if stock_id.isdigit() and len(stock_id) == 4 and stock_id not in existing_symbols:
                            radar_stocks[f"{stock_id}.TW"] = f"投信新寵 ({stock_id})"
                            count += 1
                        if count >= limit: break
                    print(f"✅ 雷達掃描完成！鎖定 {len(radar_stocks)} 檔標的。")
                    return radar_stocks
        except Exception as e:
            print(f"雷達掃描失敗 ({target_date}): {e}")
            continue
    return {}

# === 4. 持久化日誌 ===
def write_noc_log(date, symbol, name, close_price, rsi, vol_status, status, alert, predict, chip_signal):
    log_filename = "noc_trading_log.csv"
    file_exists = os.path.exists(log_filename)
    with open(log_filename, mode='a', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["日期", "代號", "名稱", "收盤價", "RSI", "量能狀態", "趨勢狀態", "戰場預判", "籌碼訊號", "行動指令"])
        writer.writerow([date, symbol, name, f"{close_price:.2f}", f"{rsi:.2f}", vol_status, status, predict, chip_signal, alert])

# === 5. FinMind 籌碼串接 ===
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

# === 6. 籌碼判定 ===
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

# === 7. 分析與預判模組 ===
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
                hist = hist.merge(chip_df, left_on='Date_Key', right_index=True, how='left')
                hist = hist.fillna({'Foreign_Inv': 0, 'Trust_Inv': 0, 'Dealer_Inv': 0})

        hist = calculate_chip_signals(hist)

        hist['5MA'] = hist['Close'].rolling(window=5).mean()
        hist['20MA'] = hist['Close'].rolling(window=20).mean()
        hist['5VMA'] = hist['Volume'].rolling(window=5).mean()
        
        delta = hist['Close'].diff()
        gain = delta.clip(lower=0); loss = -1 * delta.clip(upper=0)
        ema_gain = gain.ewm(com=13, adjust=False).mean()
        ema_loss = loss.ewm(com=13, adjust=False).mean()
        hist['RSI'] = 100 - (100 / (1 + (ema_gain / ema_loss)))
        hist['RSI'] = hist['RSI'].fillna(50)

        hist['EMA12'] = hist['Close'].ewm(span=12, adjust=False).mean()
        hist['EMA26'] = hist['Close'].ewm(span=26, adjust=False).mean()
        hist['MACD'] = hist['EMA12'] - hist['EMA26']
        hist['Signal'] = hist['MACD'].ewm(span=9, adjust=False).mean()
        hist['MACD_Hist'] = hist['MACD'] - hist['Signal']

        hist['STD20'] = hist['Close'].rolling(window=20).std()
        hist['BB_Width'] = (4 * hist['STD20']) / hist['20MA']

        hist['Is_Bottoming'] = ((hist['Close'] < hist['5MA']) & \
                               (hist['MACD_Hist'].shift(2) < hist['MACD_Hist'].shift(1)) & \
                               (hist['MACD_Hist'].shift(1) < hist['MACD_Hist']) & \
                               (hist['MACD_Hist'] < 0)).astype(int)
        
        hist['Recent_Bottoming'] = hist['Is_Bottoming'].rolling(window=3).max().fillna(0).astype(bool)

        chart_file = f"{symbol}_chart.png"
        try:
            mc = mpf.make_marketcolors(up='red', down='green', edge='black', wick='black', volume='gray')
            tw_style = mpf.make_mpf_style(base_style='yahoo', marketcolors=mc)
            mpf.plot(hist[-60:], type='candle', style=tw_style, volume=True, mav=(5, 20), title=f"Stock: {symbol}", savefig=chart_file)
        except:
            mpf.plot(hist[-60:], type='candle', style='yahoo', volume=True, mav=(5, 20), title=f"Stock: {symbol}", savefig=chart_file)

        return hist, chart_file
    except Exception as e:
        print(f"[{symbol}] 分析核心發生嚴重錯誤: {e}")
        return None

# === 8. 發送模組 ===
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

# === 9. 主程式執行 ===
if __name__ == "__main__":
    tw_tz = datetime.timezone(datetime.timedelta(hours=8))
    curr_date = datetime.datetime.now(tw_tz).date()
    curr_time = datetime.datetime.now(tw_tz).strftime("%Y-%m-%d %H:%M:%S")
    msg_list = []; generated_charts = []; has_data = False

    print(f"[{curr_time}] NOC 戰情室 v5.6 (庫存損益盤點版) 啟動...")

    # 💼 優先盤點：實體機房配置 (真實持股)
    if MY_PORTFOLIO:
        msg_list.append("━━━━━━━━━━━━━━\n💼 【庫存機櫃 (真實持股盤點)】\n━━━━━━━━━━━━━━\n")
        for sym, data in MY_PORTFOLIO.items():
            res = get_analysis_and_chart(sym, data['name'])
            if not res: continue
            hist, chart_file = res
            td = hist.iloc[-1]
            has_data = True
            generated_charts.append(chart_file)
            
            curr_price = td['Close']
            buy_price = data['buy_price']
            roi_pct = ((curr_price - buy_price) / buy_price) * 100
            
            # 損益狀態閥值判定
            if roi_pct >= TAKE_PROFIT_PCT:
                pnl_alert = f"💰【達標警戒】建議分批獲利入袋！"
            elif roi_pct <= STOP_LOSS_PCT:
                pnl_alert = f"🩸【破網警戒】請嚴格執行停損拔線！"
            elif roi_pct > 0:
                pnl_alert = f"🟢 獲利巡航中，持續抱牢。"
            else:
                pnl_alert = f"🟡 暫時浮虧，注意防守。"
                
            portfolio_msg = f"🔸 {data['name']} ({sym})\n"
            portfolio_msg += f"   成本: {buy_price:.2f} | 現價: {curr_price:.2f}\n"
            portfolio_msg += f"   損益: {roi_pct:+.2f}% | 👉 {pnl_alert}\n\n"
            msg_list.append(portfolio_msg)

    # 🛸 自動拓荒雷達掃描
    radar_targets = scan_top_trust_buy(limit=5)
    if radar_targets:
        STOCK_DICT["🛸 自動雷達 (投信最新重倉)"] = radar_targets
    else:
        msg_list.append("━━━━━━━━━━━━━━\n📂 【🛸 自動雷達 (投信最新重倉)】\n━━━━━━━━━━━━━━\n🔸 狀態: 今日掃描無符合條件標的或 API 無回應。\n\n")

    # 👀 一般外部網域監控
    for cat, stocks in STOCK_DICT.items():
        cat_printed = False 
        for sym, name in stocks.items():
            res = get_analysis_and_chart(sym, name)
            if not res: continue
            
            hist, chart_file = res
            td = hist.iloc[-1]; yd = hist.iloc[-2]
            has_data = True
            if chart_file not in generated_charts:
                generated_charts.append(chart_file)
            
            if not cat_printed and cat != "🛸 自動雷達 (投信最新重倉)":
                msg_list.append(f"━━━━━━━━━━━━━━\n📂 【{cat}】\n━━━━━━━━━━━━━━\n")
                cat_printed = True

            vol_today = td['Volume']; vma5 = td['5VMA']
            vol_status = "📈 出量" if vol_today > vma5 * 1.2 else "📉 量縮" if vol_today < vma5 * 0.8 else "➖ 量平"
            trend_status = "🔥 多頭" if td['Close'] > td['5MA'] > td['20MA'] else "🧊 空頭" if td['Close'] < td['5MA'] < td['20MA'] else "🔄 盤整"
            chip_status = td['Chip_Status']

            predict_msg = "無特殊徵兆"
            if td['BB_Width'] < 0.08: predict_msg = "⚠️【大變盤預警】布林通道極度壓縮！"
            elif td['Is_Bottoming'] == 1: predict_msg = "📈【築底預判】空方動能連續收斂！"

            is_breakout = (yd['Close'] < yd['5MA']) and (td['Close'] > td['5MA']) and (vol_today > vma5 * 1.2)
            if td['Recent_Bottoming'] and is_breakout: alert = "🚀【狙擊模式：強烈買進】底部完成且帶量突破！"
            elif td['RSI'] > 80: alert = "💰【獲利了結】短線過熱，注意回檔。"
            elif td['Close'] < td['5MA'] < td['20MA'] and vol_today > vma5 * 1.2: alert = "💀【強制退場】空頭確認，大單砸盤！"
            else: alert = "✅【持股續抱】順勢操作，等待訊號。"

            write_noc_log(curr_date, sym, name, td['Close'], td['RSI'], vol_status, trend_status, predict_msg, chip_status, alert)
            
            stock_msg = f"🔸 {name} ({sym})\n   現價: {td['Close']:.2f} | RSI: {td['RSI']:.1f}\n   狀態: {trend_status} | {vol_status}\n"
            if chip_status != "無資料": stock_msg += f"   💰 籌碼: {chip_status}\n"
            stock_msg += f"   🔮 預判: {predict_msg}\n   👉 指令: {alert}\n\n"
            msg_list.append(stock_msg)

    if has_data or len(msg_list) > 0:
        final_text = f"📡 【NOC 戰情室 v5.6：資產損益控管】\n📅 時間：{curr_time}\n━━━━━━━━━━━━━━\n" + "".join(msg_list)
        send_reports(f"NOC 戰情報告 {curr_date}", final_text, generated_charts)
        for chart in generated_charts:
            if os.path.exists(chart): os.remove(chart)
    else:
        print("休市或資料讀取失敗，伺服器待命。")
