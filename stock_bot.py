import yfinance as yf
import requests
import os
import datetime
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

# === 2. 專屬通訊錄 ===
STOCK_DICT = {
    "2330.TW": "台積電",
    "2317.TW": "鴻海",
    "0050.TW": "元大台灣50",
    "009816.TW": "凱基台灣TOP50",
    "8431.TWO": "匯鑽科",
    "AAPL": "蘋果 (Apple)",
    "NVDA": "輝達 (NVIDIA)"
}

# === 3. 分析與繪圖模組 ===
def get_analysis_and_draw_chart(symbol, name):
    try:
        stock = yf.Ticker(symbol)
        hist = stock.history(period="6mo")
        if len(hist) < 30: return None
        
        # 繪製 K線圖並存成圖片檔
        # mplfinance 會自動幫我們畫出 K線、成交量，以及我們指定的均線(5, 20)
        chart_filename = f"{symbol}_chart.png"
        mpf.plot(hist, type='candle', style='yahoo', volume=True, 
                 mav=(5, 20), title=f"{name} ({symbol})", 
                 savefig=chart_filename)
        
        # 數據計算 (與之前相同)
        hist['5MA'] = hist['Close'].rolling(window=5).mean()
        hist['20MA'] = hist['Close'].rolling(window=20).mean()
        hist['5VMA'] = hist['Volume'].rolling(window=5).mean()
        
        delta = hist['Close'].diff()
        gain = delta.clip(lower=0)
        loss = -1 * delta.clip(upper=0)
        ema_gain = gain.ewm(com=13, adjust=False).mean()
        ema_loss = loss.ewm(com=13, adjust=False).mean()
        rs = ema_gain / ema_loss
        hist['RSI'] = 100 - (100 / (1 + rs))
        
        td = hist.iloc[-1]
        yd = hist.iloc[-2]
        last_trade_date = hist.index[-1].date()
        
        return td, yd, last_trade_date, chart_filename
    except Exception as e:
        print(f"[{symbol}] 分析或繪圖失敗: {e}")
        return None

# === 4. Telegram 發送模組 ===
def send_telegram_msg(msg):
    url = f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage"
    requests.post(url, json={"chat_id": TG_CHAT_ID, "text": msg})

# === 5. Email 發送模組 (含圖片附件) ===
def send_email_report(subject, text_body, image_files):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_USER
    msg['To'] = EMAIL_TO
    msg['Subject'] = subject

    # 加入文字內容
    msg.attach(MIMEText(text_body, 'plain'))

    # 將剛才畫好的圖表一張一張夾帶進去
    for img_file in image_files:
        if os.path.exists(img_file):
            with open(img_file, 'rb') as f:
                img_data = f.read()
            image = MIMEImage(img_data, name=os.path.basename(img_file))
            msg.attach(image)

    # 透過 Gmail 的 SMTP 伺服器發送
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_USER, EMAIL_PASS)
        server.send_message(msg)
        server.quit()
        print("✅ Email 戰情報告發送成功！")
    except Exception as e:
        print(f"❌ Email 發送失敗: {e}")

# === 6. 執行主程式 ===
tw_tz = datetime.timezone(datetime.timedelta(hours=8))
current_tw_date = datetime.datetime.now(tw_tz).date()

message_list = []
generated_charts = []
has_new_data = False 

print(f"今天是 {current_tw_date}，開始執行 NOC 戰情室掃描...")

for symbol, name in STOCK_DICT.items():
    data = get_analysis_and_draw_chart(symbol, name)
    
    if data is not None:
        td, yd, last_trade_date, chart_file = data
        
        # 假日防禦
        is_market_open = False
        if ".TW" in symbol:
            if last_trade_date == current_tw_date: is_market_open = True
        else:
            if last_trade_date >= current_tw_date - datetime.timedelta(days=1): is_market_open = True
                
        if not is_market_open:
            continue 
            
        has_new_data = True 
        generated_charts.append(chart_file) # 把畫好的圖表記錄下來
        
        # 數據判定
        rsi_value = td['RSI']
        vol_today = td['Volume']
        vma5 = td['5VMA']
        
        td_strong = td['Close'] > td['5MA'] > td['20MA']
        yd_strong = yd['Close'] > yd['5MA'] > yd['20MA']
        td_weak = td['Close'] < td['5MA'] < td['20MA']

        if vol_today > vma5 * 1.2: vol_status = "📈 出量"
        elif vol_today < vma5 * 0.8: vol_status = "📉 量縮"
        else: vol_status = "➖ 量平"

        if td_strong: status = "🔥 多頭排列"
        elif td_weak: status = "🧊 空頭排列"
        else: status = "🔄 盤整震盪"
        
        alert = ""
        if td_weak and vol_today > vma5 * 1.2: alert = "💀【恐慌殺盤】切勿接刀！"
        elif yd['Close'] < yd['5MA'] and td['Close'] > td['5MA']:
            if vol_today > vma5 * 1.2: alert = "🚀【強力反轉】帶量站回5日線！"
            else: alert = "📈【弱勢反彈】站回5日線但無量。"
        elif yd['5MA'] < yd['20MA'] and td['5MA'] > td['20MA']: alert = "🌟【黃金交叉】"
        elif yd_strong and not td_strong: alert = "⚠️【警戒】由強轉弱。"
        elif rsi_value < 30: alert = "🟢【超跌】RSI低於30。"
        elif rsi_value > 70: alert = "🔴【過熱】RSI高於70。"
        else: alert = "✅ 穩定運行中"

        # 排版
        stock_msg = f"🔸 {name} ({symbol})\n   現價: {td['Close']:.2f} | RSI: {rsi_value:.1f}\n   量能: {vol_status}\n   狀態: {status}\n   訊號: {alert}\n\n"
        message_list.append(stock_msg)

if has_new_data and len(message_list) > 0:
    final_text = "📡 【老網管 NOC 戰情室：全方位轉折預警】\n\n" + "".join(message_list)
    
    # 1. 發送簡短文字到 Telegram 讓你第一時間知道
    send_telegram_msg(final_text + "💡 詳細技術線圖已寄至您的 Email。")
    
    # 2. 發送包含「文字 + 實體 K 線圖附件」到 Email
    email_subject = f"📊 理財儀表板戰情日報 ({current_tw_date})"
    send_email_report(email_subject, final_text, generated_charts)
else:
    print("😴 判定為休市，暫停發送報告。")
