import yfinance as yf
import requests
import os
import datetime
import pandas as pd
import numpy as np
import mplfinance as mpf
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

# === 1. 從 GitHub 保險箱抓取機密 ===
TG_TOKEN = os.environ.get("TG_TOKEN")
TG_CHAT_ID = os.environ.get("TG_CHAT_ID")
EMAIL_USER = os.environ.get("EMAIL_USER")
EMAIL_PASS = os.environ.get("EMAIL_PASS")
EMAIL_TO = os.environ.get("EMAIL_TO")

# === 2. 專屬通訊錄 (四層 VLAN 網域版) ===
STOCK_DICT = {
    "🛡️ 核心持股 (重倉伺服器)": {
        "3037.TW": "欣興 (ABF載板)"
    },
    "🔥 潛力種子 (高頻寬觀察區)": {
        "3163.TW": "波若威 (光通訊)",
        "5388.TW": "中磊 (網通設備)",
        "3714.TW": "富采 (LED光電)"
    },
    "👀 常態觀察區 (例行監控節點)": {
        "2330.TW": "台積電",
        "2317.TW": "鴻海",
        "0050.TW": "元大台灣50",
        "009816.TW": "凱基台灣TOP50",
        "8431.TWO": "匯鑽科",
        "AAPL": "蘋果 (Apple)",
        "NVDA": "輝達 (NVIDIA)"
    },
    "💾 記憶體族群 (美光連動網域)": {
        "MU": "美光 (Micron)", 
        "2408.TW": "南亞科 (DRAM製造)",
        "2344.TW": "華邦電 (利基記憶體)",
        "6239.TW": "力成 (記憶體封測)", 
        "3260.TW": "威剛 (記憶體模組)", 
        "8299.TW": "群聯 (Flash控制IC)", 
        "4967.TW": "十銓 (電競模組)"
    }
}

# === 3. Phase 2: 籌碼面分析與訊號生成模組 ===
def calculate_chip_signals(hist: pd.DataFrame) -> pd.DataFrame:
    """
    計算籌碼與價量指標。
    防呆機制：若無法人資料，僅計算 MA、VMA 等價量指標。
    """
    # 確保不會發生 SyntaxError 的 5VMA 寫法
    hist['5MA'] = hist['Close'].rolling(window=5).mean()
    hist['20MA'] = hist['Close'].rolling(window=20).mean()
    hist['5VMA'] = hist['Volume'].rolling(window=5).mean()

    # RSI 計算
    delta = hist['Close'].diff()
    gain = delta.clip(lower=0)
    loss = -1 * delta.clip(upper=0)
    ema_gain = gain.ewm(com=13, adjust=False).mean()
    ema_loss = loss.ewm(com=13, adjust=False).mean()
    rs = ema_gain / ema_loss
    hist['RSI'] = 100 - (100 / (1 + rs))

    # 檢查是否具備籌碼欄位 (為未來串接台股 API 鋪路)
    required_chip_cols = ['Foreign_Inv', 'Trust_Inv', 'Dealer_Inv']
    if all(col in hist.columns for col in required_chip_cols):
        # 計算三大法人買賣超佔總成交量的比例
        hist['Total_Institutional'] = hist['Foreign_Inv'] + hist['Trust_Inv'] + hist['Dealer_Inv']
        hist['Inst_Volume_Ratio'] = np.where(
            hist['Volume'] > 0, 
            hist['Total_Institutional'] / hist['Volume'], 
            0
        )
        
        # 連買訊號計算
        hist['Foreign_Buy_Flag'] = (hist['Foreign_Inv'] > 0).astype(int)
        hist['Trust_Buy_Flag'] = (hist['Trust_Inv'] > 0).astype(int)
        hist['Trust_Buy_Days_5d'] = hist['Trust_Buy_Flag'].rolling(window=5).sum()

        # 訊號生成
        hist['Signal_CoBuy'] = (hist['Foreign_Inv'] > 0) & (hist['Trust_Inv'] > 0)
        hist['Signal_Trust_Trend'] = (hist['Trust_Buy_Days_5d'] >= 4) & (hist['Trust_Buy_Flag'] == 1)
    
    return hist

# === 4. 分析、繪圖與情報抓取模組 ===
def get_analysis_and_draw_chart(symbol, name):
    try:
        stock = yf.Ticker(symbol)
        hist = stock.history(period="6mo")
        if len(hist) < 30: return None

        # 呼叫 Phase 2 模組進行指標運算
        hist = calculate_chip_signals(hist)

        chart_filename = f"{symbol}_chart.png"
        mpf.plot(hist, type='candle', style='yahoo', volume=True, 
                 mav=(5, 20), title=f"{name} ({symbol})", 
                 savefig=chart_filename)

        td = hist.iloc[-1]
        yd = hist.iloc[-2]
        last_trade_date = hist.index[-1].date()

        # 抓取新聞情報
        latest_news = None
        try:
            news_list = stock.news
            if news_list and len(news_list) > 0:
                latest_news = news_list[0]
        except:
            pass 

        return td, yd, last_trade_date, chart_filename, latest_news
    except Exception as e:
        print(f"[{symbol}] 分析失敗: {e}")
        return None

# === 5. Telegram 發送模組 ===
def send_telegram_msg(msg):
    # 若無設定 TG Token 則略過 (防呆)
    if not TG_TOKEN or not TG_CHAT_ID:
        print("未設定 Telegram Token，略過發送。")
        return
    url = f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage"
    payload = {"chat_id": TG_CHAT_ID, "text": msg, "disable_web_page_preview": True}
    requests.post(url, json=payload)

# === 6. Email 發送模組 ===
def send_email_report(subject, text_body, image_files):
    if not EMAIL_USER or not EMAIL_PASS:
        return
    msg = MIMEMultipart()
    msg['From'] = EMAIL_USER
    msg['To'] = EMAIL_TO
    msg['Subject'] = subject
    msg.attach(MIMEText(text_body, 'plain'))
    
    for img_file in image_files:
        if os.path.exists(img_file):
            with open(img_file, 'rb') as f:
                img_data = f.read()
            image = MIMEImage(img_data, name=os.path.basename(img_file))
            msg.attach(image)
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_USER, EMAIL_PASS)
        server.send_message(msg)
        server.quit()
        print("✅ Email 發送成功！")
    except Exception as e:
        print(f"❌ Email 發送失敗: {e}")

# === 7. 執行主程式 ===
if __name__ == "__main__":
    tw_tz = datetime.timezone(datetime.timedelta(hours=8))
    
    # 用來計算休市邏輯的日期
    current_tw_date = datetime.datetime.now(tw_tz).date()
    # 🕒 升級版：精確到秒的時間戳記
    current_tw_time = datetime.datetime.now(tw_tz).strftime("%Y-%m-%d %H:%M:%S")

    message_list = []
    generated_charts = []
    has_new_data = False 

    # Log 印出精確時間
    print(f"[{current_tw_time}] 開始執行 NOC 暴力防呆戰情室 (v2.0 籌碼升級版)...")

    for category, stocks in STOCK_DICT.items():
        message_list.append(f"━━━━━━━━━━━━━━\n📂 【{category}】\n━━━━━━━━━━━━━━\n")
        
        for symbol, name in stocks.items():
            result = get_analysis_and_draw_chart(symbol, name)
            if result:
                td, yd, last_trade_date, chart_file, news = result
                
                # 判斷是否有最新交易日資料 (簡易防呆)
                if last_trade_date == current_tw_date:
                    has_new_data = True
                
                # 這裡可以根據 td (今日數據) 組裝您要的訊息格式
                # 範例：顯示收盤價與 RSI
                close_price = td['Close']
                rsi_value = td['RSI'] if not np.isnan(td['RSI']) else 0
                message_list.append(f"🔹 {name} ({symbol}): 收盤 {close_price:.2f} | RSI: {rsi_value:.1f}\n")
                
                if chart_file:
                    generated_charts.append(chart_file)

    # 📡 Timestamp 升級與心跳封包 (Heartbeat) 架構
    if has_new_data and len(message_list) > 0:
        final_text = f"📡 【老網管 NOC 指揮中心：行動清單】\n📅 系統時間：{current_tw_time}\n━━━━━━━━━━━━━━\n" + "".join(message_list)
        final_text += "⚠️ 老網管提醒：收到指令請馬上動作，猶豫就會敗北！"
        send_telegram_msg(final_text)
        print("✅ 戰情報告已發送至 Telegram。")
    else:
        sleep_msg = f"📡 【老網管 NOC 指揮中心：休市回報】\n📅 系統時間：{current_tw_time}\n😴 報告：今日台股休市或無最新資料，戰情室伺服器進入待命模式 (Standby)。"
        print(f"[{current_tw_time}] 今日休市或無更新，已發送待命通知至 Telegram。")
        send_telegram_msg(sleep_msg)
