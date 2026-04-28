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
    "🔥 潛力種子 (高頻寬觀察區)": {"3163.TW": "波若威", "5388.TW": "中磊", "3714.TW": "富采", "6269.TW": "台郡"},
    "👀 常態觀察區 (例行監控節點)": {"2330.TW": "台積電", "0050.TW": "元大台灣50","AAPL": "蘋果","NVDA": "輝達"},
    "💾 記憶體族群 (美光連動網域)": {"MU": "美光", "2408.TW": "南亞科", "3260.TW": "威剛", "8299.TW": "群聯", },
    "🔍 YAHOO 觀察區": {"2027.TW": "大成鋼", "2382.TW": "廣達", "2886.TW": "兆豐金", "2409.TW": "友達", "2352.TW": "佳世達", },
    "真實持股 追蹤區" : {"8431.TWO":"匯鑽科","3231.TW":"緯創" }
}

# === 2.1 真實持股庫存 (實體機房配置) ===
MY_PORTFOLIO = {
    "3231.TW": {"name": "緯創", "buy_price": 130.5, "shares": 1000},
    "8431.TWO": {"name": "匯鑽科", "buy_price": 70.7, "shares": 1000},
    "6116.TW": {"name": "彩晶", "buy_price": 8.4, "shares": 1000},
    "2317.TW": {"name": "鴻海", "buy_price": 201.5, "shares": 1000}
}

TAKE_PROFIT_PCT = 20.0  
STOP_LOSS_PCT = -10.0   

# === 3. 功能函式庫 ===

def get_etf_tag(symbol, name):
    """自動判定 ETF 類型並回傳標籤與乖離門檻"""
    div_keys = ["高股息", "優息", "0056", "00878", "00919", "00929", "00940"]
    market_keys = ["0050", "006208", "50", "市值", "AAPL", "NVDA", "TSM"]
    
    if any(k in name or k in symbol for k in div_keys):
        return "💰", 5.0, "高股息(控管殖利率)"
    elif any(k in name or k in symbol for k in market_keys):
        return "🚀", 10.0, "市值型(移動停損)"
    return "⚙️", 10.0, "一般個股(趨勢防守)"

def write_noc_log(date, symbol, name, close_price, rsi, vol_status, status, alert, predict, chip_signal):
    log_filename = "noc_trading_log.csv"
    file_exists = os.path.exists(log_filename)
    with open(log_filename, mode='a', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["日期", "代號", "名稱", "收盤價", "RSI", "量能狀態", "趨勢狀態", "戰場預判", "籌碼訊號", "行動指令"])
        writer.writerow([date, symbol, name, f"{close_price:.2f}", f"{rsi:.2f}", vol_status, status, predict, chip_signal, alert])

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
        ema_gain = gain.ewm(com=13, adjust=False).mean(); ema_loss = loss.ewm(com=13, adjust=False).mean()
        hist['RSI'] = 100 - (100 / (1 + (ema_gain / ema_loss)))
        hist['RSI'] = hist['RSI'].fillna(50)
        hist['20_High'] = hist['High'].rolling(window=20).max().shift(1)
        hist['Body_Top'] = hist[['Open', 'Close']].max(axis=1)
        hist['Upper_Shadow'] = hist['High'] - hist['Body_Top']
        hist['K_Length'] = (hist['High'] - hist['Low']).replace(0, 0.001)
        hist['Shadow_Ratio'] = hist['Upper_Shadow'] / hist['K_Length']
        
        chart_file = f"{symbol}_chart.png"
        try:
            # 修正相容性問題：移除 base_style，改用 try-except 包裝
            mc = mpf.make_marketcolors(up='red', down='green', edge='black', wick='black', volume='gray')
            try:
                s = mpf.make_mpf_style(base_style='yahoo', marketcolors=mc)
            except:
                s = mpf.make_mpf_style(marketcolors=mc)
            mpf.plot(hist[-60:], type='candle', style=s, volume=True, mav=(5, 20), savefig=chart_file)
        except Exception as chart_err:
            print(f"[{symbol}] 畫圖模組跳過: {chart_err}")
            # 建立空的檔名確保後續不崩潰，或者用最簡風格
            mpf.plot(hist[-60:], type='candle', savefig=chart_file)

        return hist, chart_file
    except Exception as e:
        print(f"[{symbol}] 分析核心錯誤: {e}"); return None

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
                server.login(EMAIL_USER, EMAIL_PASS); server.send_message(msg)
        except Exception as e: print(f"❌ Email 發送失敗: {e}")

# === 4. 主程式執行 ===
if __name__ == "__main__":
    tw_tz = datetime.timezone(datetime.timedelta(hours=8))
    curr_date = datetime.datetime.now(tw_tz).date()
    curr_time = datetime.datetime.now(tw_tz).strftime("%Y-%m-%d %H:%M:%S")
    msg_list = []; generated_charts = []; has_data = False

    print(f"[{curr_time}] NOC 戰情室 v6.1 (相容修復版) 啟動...")

    # A. 💼 實體機房配置 (真實持股)
    if MY_PORTFOLIO:
        msg_list.append("━━━━━━━━━━━━━━\n💼 【庫存機櫃 (實體持股)】\n━━━━━━━━━━━━━━\n")
        for sym, data in MY_PORTFOLIO.items():
            res = get_analysis_and_chart(sym, data['name'])
            if not res: continue
            hist, chart_file = res; td = hist.iloc[-1]
            has_data = True; generated_charts.append(chart_file)
            
            curr_price = td['Close']; buy_price = data['buy_price']
            roi_pct = ((curr_price - buy_price) / buy_price) * 100
            tag, b_limit, b_desc = get_etf_tag(sym, data['name'])
            
            ma20 = td['20MA']
            bias = ((curr_price - ma20) / ma20) * 100
            bias_alert = "🚨乖離過大" if bias > b_limit else "✅穩定"

            pnl_alert = "💰達標" if roi_pct >= TAKE_PROFIT_PCT else "🩸破網" if roi_pct <= STOP_LOSS_PCT else "🟢持平"
            
            vol_status = "📈出量" if td['Volume'] > td['5VMA']*1.2 else "➖量縮"
            trend_status = "🔥多頭" if td['Close'] > td['5MA'] > td['20MA'] else "🔄盤整"

            write_noc_log(curr_date, sym, f"{tag}{data['name']}", curr_price, td['RSI'], vol_status, trend_status, pnl_alert, bias_alert, td['Chip_Status'])
            
            msg_list.append(f"{tag} {data['name']} ({sym})\n   損益: {roi_pct:+.2f}% | 乖離: {bias:+.1f}% ({bias_alert})\n   👉 {pnl_alert}\n\n")

    # B. 👀 一般外部網域監控
    for cat, stocks in STOCK_DICT.items():
        if not stocks: continue
        msg_list.append(f"━━━━━━━━━━━━━━\n📂 【{cat}】\n━━━━━━━━━━━━━━\n")
        for sym, name in stocks.items():
            res = get_analysis_and_chart(sym, name)
            if not res: continue
            hist, chart_file = res; td = hist.iloc[-1]; yd = hist.iloc[-2]
            has_data = True; generated_charts.append(chart_file)
            
            tag, b_limit, b_desc = get_etf_tag(sym, name)
            ma20 = td['20MA']
            bias = ((td['Close'] - ma20) / ma20) * 100
            bias_status = "🚨過熱" if bias > b_limit else "✅穩定"
            
            vol_status = "📈出量" if td['Volume'] > td['5VMA']*1.2 else "➖量平"
            trend_status = "🔥多頭" if td['Close'] > td['5MA'] > td['20MA'] else "🔄盤整"
            
            predict = "🚀帶量突破" if td['Close'] > td['20_High'] and td['Volume'] > td['5VMA'] else "無特殊徵兆"
            alert = "🚀進場" if bias <= b_limit and td['Close'] > td['5MA'] else "✅續抱"

            write_noc_log(curr_date, sym, f"{tag}{name}", td['Close'], td['RSI'], vol_status, trend_status, alert, predict, td['Chip_Status'])
            
            msg_list.append(f"{tag} {name} ({sym})\n   現價: {td['Close']:.2f} | 乖離: {bias:+.1f}% ({bias_status})\n   👉 {alert} | {predict}\n\n")

    if has_data:
        final_text = f"📡 NOC 戰情室 v6.1\n📅 時間：{curr_time}\n" + "".join(msg_list)
        send_reports(f"NOC 戰情報告 {curr_date}", final_text, generated_charts)
        for c in generated_charts: 
            if os.path.exists(c): os.remove(c)
    else:
        print("休市或資料讀取失敗。")
