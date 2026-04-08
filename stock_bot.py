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
EMAIL_USER = os.environ.get("EMAIL_USER")  # 你的 Gmail
EMAIL_PASS = os.environ.get("EMAIL_PASS")  # Gmail 應用程式密碼
EMAIL_TO = os.environ.get("EMAIL_TO")      # 收件者信箱
FINMIND_TOKEN = os.environ.get("FINMIND_TOKEN")  # 🔑 FinMind Token

# === 2. 專屬通訊錄 (18檔監控名單) ===
STOCK_DICT = {
    "🛡️ 核心持股 (重倉伺服器)": {"3037.TW": "欣興 (ABF載板)"},
    "🔥 潛力種子 (高頻寬觀察區)": {"3163.TW": "波若威", "5388.TW": "中磊", "3714.TW": "富采"},
    "👀 常態觀察區 (例行監控節點)": {"2330.TW": "台積電", "2317.TW": "鴻海", "0050.TW": "元大台灣50"},
    "💾 記憶體族群 (美光連動網域)": {"MU": "美光", "2408.TW": "南亞科", "3260.TW": "威剛", "8299.TW": "群聯", "AAPL": "蘋果", "NVDA": "輝達"}
}

# === 3. 持久化日誌 ===
def write_noc_log(date, symbol, name, close_price, rsi, vol_status, status, alert, predict, chip_signal):
    log_filename = "noc_trading_log.csv"
    file_exists = os.path.exists(log_filename)
    with open(log_filename, mode='a', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["日期", "代號", "名稱", "收盤價", "RSI", "量能狀態", "趨勢狀態", "戰場預判", "籌碼訊號", "行動指令"])
        writer.writerow([date, symbol, name, f"{close_price:.2f}", f"{rsi:.2f}", vol_status, status, predict, chip_signal, alert])

# === 4. 🔌 FinMind API 串接模組 ===
def get_finmind_chip_data(symbol, start_date_str):
    if not FINMIND_TOKEN:
        return pd.DataFrame()
    
    fm_symbol = symbol.replace(".TW", "").replace(".TWO", "")
    if not fm_symbol.isdigit(): # 遇到美股 (如 MU, AAPL) 直接略過
        return pd.DataFrame()

    url = "https://api.finmindtrade.com/api/v4/data"
    params = {
        "dataset": "TaiwanStockInstitutionalInvestorsBuySell",
        "data_id": fm_symbol,
        "start_date": start_date_str,
        "token": FINMIND_TOKEN
    }
    
    try:
        r = requests.get(url, params=params, timeout=10)
        data = r.json()
        if data.get("msg") == "success" and len(data.get("data", [])) > 0:
            df = pd.DataFrame(data["data"])
            df['net_buy'] = df['buy'] - df['sell']
            
            # 分類三大法人
            df['type'] = 'Other'
            df.loc[df['name'].str.contains('外資'), 'type'] = 'Foreign_Inv'
            df.loc[df['name'].str.contains('投信'), 'type'] = 'Trust_Inv'
            df.loc[df['name'].str.contains('自營商'), 'type'] = 'Dealer_Inv'
            
            # 樞紐分析轉換為日線欄位
            pivot_df = df.groupby(['date', 'type'])['net_buy'].sum().unstack(fill_value=0).reset_index()
            
            # 確保欄位存在
            for col in ['Foreign_Inv', 'Trust_Inv', 'Dealer_Inv']:
                if col not in pivot_df.columns: 
                    pivot_df[col] = 0
            
            pivot_df['Date'] = pd.to_datetime(pivot_df['date']).dt.date
            pivot_df.set_index('Date', inplace=True)
            return pivot_df[['Foreign_Inv', 'Trust_Inv', 'Dealer_Inv']]
    except Exception as e:
        print(f"[{symbol}] FinMind 籌碼撈取失敗: {e}")
        
    return pd.DataFrame()

# === 5. Phase 2 籌碼運算模組 ===
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
        
        # 簡易籌碼狀態標註
        conditions = [
            (hist['Signal_CoBuy'] == True),
            (hist['Signal_Trust_Trend'] == True),
            (hist['Total_Institutional'] > 0)
        ]
        choices = ["🤝 土洋齊買", "🏦 投信連買", "📈 法人偏多"]
        hist['Chip_Status'] = np.select(conditions, choices, default="➖ 中性/偏空")
        
    return hist

# === 6. 分析與預判模組 (含嚴格狙擊) ===
def get_analysis_and_chart(symbol, name):
    try:
        stock = yf.Ticker(symbol)
        hist = stock.history(period="6mo")
        if len(hist) < 30: 
            return None

        hist['Date_Key'] = hist.index.date
        
        # 🔌 撈取籌碼並合併
        if FINMIND_TOKEN and (".TW" in symbol or ".TWO" in symbol):
            start_date_str = (datetime.datetime.now() - datetime.timedelta(days=180)).strftime("%Y-%m-%d")
            chip_df = get_finmind_chip_data(symbol, start_date_str)
            if not chip_df.empty:
                hist = hist.merge(chip_df, left_on='Date_Key', right_index=True, how='left')
                hist.fillna({'Foreign_Inv': 0, 'Trust_Inv': 0, 'Dealer_Inv': 0}, inplace=True)

        hist = calculate_chip_signals(hist)

        # 基本 MA 與 Volume
        hist['5MA'] = hist['Close'].rolling(window=5).mean()
        hist['20MA'] = hist['Close'].rolling(window=20).mean()
        hist['5VMA'] = hist['Volume'].rolling(window=5).mean()
        
        # RSI
        delta = hist['Close'].diff()
        gain = delta.clip(lower=0)
        loss = -1 * delta.clip(upper=0)
        ema_gain = gain.ewm(com=13, adjust=False).mean()
        ema_loss = loss.ewm(com=13, adjust=False).mean()
        hist['RSI'] = 100 - (100 / (1 + (ema_gain / ema_loss)))

        # 🔮 MACD 預判動能
        hist['EMA12'] = hist['Close'].ewm(span=12, adjust=False).mean()
        hist['EMA26'] = hist['Close'].ewm(span=26, adjust=False).mean()
        hist['MACD'] = hist['EMA12'] - hist['EMA26']
        hist['Signal'] = hist['MACD'].ewm(span=9, adjust=False).mean()
        hist['MACD_Hist'] = hist['MACD'] - hist['Signal']

        # 🔮 布林通道壓縮 (BB Squeeze)
        hist['STD20'] = hist['Close'].rolling(window=20).std()
        hist['BB_Upper'] = hist['20MA'] + (2 * hist['STD20'])
        hist['BB_Lower'] = hist['20MA'] - (2 * hist['STD20'])
        hist['BB_Width'] = (hist['BB_Upper'] - hist['BB_Lower']) / hist['20MA']

        # ---------------------------------------------------------
        # 🎯 核心：狙擊模式偵測
        # ---------------------------------------------------------
        # 條件 1：價格在5MA下且MACD綠柱連三天收斂
        hist['Is_Bottoming'] = (hist['Close'] < hist['5MA']) & \
                               (hist['MACD_Hist'].shift(2) < hist['MACD_Hist'].shift(1)) & \
                               (hist['MACD_Hist'].shift(1) < hist['MACD_Hist']) & \
                               (hist['MACD_Hist'] < 0)
        
        # 條件 2：檢查過去 3 天(含今天) 是否有築底訊號
        hist['Recent_Bottoming'] = hist['Is_Bottoming'].rolling(window=3).max().fillna(0).astype(bool)

        # 繪圖
        chart_file = f"{symbol}_chart.png"
        mpf.plot(hist[-60:], type='candle', style='yahoo', volume=True, 
                 mav=(5, 20), title=f"{name} ({symbol})", savefig=chart_file)

        return hist, chart_file
    except Exception as e:
        print(f"[{symbol}] 分析發生錯誤: {e}")
        return None

# === 7. 發送模組 (Telegram + Email) ===
def send_telegram(msg):
    if TG_TOKEN and TG_CHAT_ID:
        requests.post(f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage", 
                      json={"chat_id": TG_CHAT_ID, "text": msg, "disable_web_page_preview": True})

def send_email_report(subject, text_body, chart_files):
    if not EMAIL_USER or not EMAIL_PASS or not EMAIL_TO:
        print("未設定 Email 環境變數，略過發送。")
        return
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_USER
        msg['To'] = EMAIL_TO
        msg['Subject'] = subject
        msg.attach(MIMEText(text_body, 'plain'))
        
        for chart in chart_files:
            if os.path.exists(chart):
                with open(chart, 'rb') as f:
                    img = MIMEImage(f.read(), name=os.path.basename(chart))
                    msg.attach(img)
        
        log_file = "noc_trading_log.csv"
        if os.path.exists(log_file):
            with open(log_file, 'rb') as f:
                csv_part = MIMEApplication(f.read(), Name=log_file)
                csv_part.add_header('Content-Disposition', f'attachment; filename="{log_file}"')
                msg.attach(csv_part)
        
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(EMAIL_USER, EMAIL_PASS)
            server.send_message(msg)
        print("✅ Email 戰報 (含圖表與 CSV 日誌) 發送成功！")
    except Exception as e:
        print(f"❌ Email 發送失敗: {e}")

# === 8. 主程式執行 ===
if __name__ == "__main__":
    tw_tz = datetime.timezone(datetime.timedelta(hours=8))
    curr_date = datetime.datetime.now(tw_tz).date()
    curr_time = datetime.datetime.now(tw_tz).strftime("%Y-%m-%d %H:%M:%S")
    
    msg_list = []
    generated_charts = []
    has_data = False

    print(f"[{curr_time}] NOC 戰情室 v4.0 (嚴格狙擊 + FinMind籌碼) 啟動...")

    for cat, stocks in STOCK_DICT.items():
        msg_list.append(f"━━━━━━━━━━━━━━\n📂 【{cat}】\n━━━━━━━━━━━━━━\n")
        for sym, name in stocks.items():
            res = get_analysis_and_chart(sym, name)
            if not res: 
                continue
            
            hist, chart_file = res
            td = hist.iloc[-1]
            yd = hist.iloc[-2]
            last_date = hist.index[-1].date()
            
            if last_date != curr_date and ".TW" in sym: 
                continue
            
            has_data = True
            generated_charts.append(chart_file)

            # 基本狀態判斷
            vol_today = td['Volume']
            vma5 = td['5VMA']
            vol_status = "📈 出量" if vol_today > vma5 * 1.2 else "📉 量縮" if vol_today < vma5 * 0.8 else "➖ 量平"
            trend_status = "🔥 多頭" if td['Close'] > td['5MA'] > td['20MA'] else "🧊 空頭" if td['Close'] < td['5MA'] < td['20MA'] else "🔄 盤整"

            chip_status = td['Chip_Status']

            # 預判雷達邏輯
            predict_msg = "無特殊徵兆"
            if td['BB_Width'] < 0.08:
                predict_msg = "⚠️【大變盤預警】布林通道極度壓縮！"
            elif td['Is_Bottoming']:
                predict_msg = "📈【築底預判】空方動能連續收斂！"

            # 🛡️ 狙擊指令邏輯 (需同時滿足近3日築底 + 今日突破出量)
            is_breakout = (yd['Close'] < yd['5MA']) and (td['Close'] > td['5MA']) and (vol_today > vma5 * 1.2)
            
            if td['Recent_Bottoming'] and is_breakout:
                alert = "🚀【狙擊模式：強烈買進】底部完成且帶量突破！"
            elif td['RSI'] > 80:
                alert = "💰【獲利了結】短線過熱，注意回檔。"
            elif td['Close'] < td['5MA'] < td['20MA'] and vol_today > vma5 * 1.2:
                alert = "💀【強制退場】空頭確認，大單砸盤！"
            else:
                alert = "✅【持股續抱】順勢操作，等待訊號。"

            # 寫入 CSV
            write_noc_log(curr_date, sym, name, td['Close'], td['RSI'], vol_status, trend_status, predict_msg, chip_status, alert)

            # 排版字串
            stock_msg = f"🔸 {name} ({sym})\n"
            stock_msg += f"   現價: {td['Close']:.2f} | RSI: {td['RSI']:.1f}\n"
            stock_msg += f"   狀態: {trend_status} | {vol_status}\n"
            if chip_status != "無資料":
                stock_msg += f"   💰 籌碼: {chip_status}\n"
            stock_msg += f"   🔮 預判: {predict_msg}\n"
            stock_msg += f"   👉 指令: {alert}\n\n"
            msg_list.append(stock_msg)

    # 戰報發送
    if has_data:
        final_text = f"📡 【NOC 戰情室 v4.0：嚴格狙擊版】\n📅 時間：{curr_time}\n━━━━━━━━━━━━━━\n" + "".join(msg_list)
        
        send_telegram(final_text)
        send_email_report(f"NOC 戰情報告 {curr_date}", final_text, generated_charts)
        
        for chart in generated_charts:
            if os.path.exists(chart): 
                os.remove(chart)
    else:
        print("休市，伺服器待命。")
