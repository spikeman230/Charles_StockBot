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

# === 3. 分析、繪圖與【情報抓取】模組 ===
def get_analysis_and_draw_chart(symbol, name):
    try:
        stock = yf.Ticker(symbol)
        hist = stock.history(period="6mo")
        if len(hist) < 30: return None
        
        # 繪製 K線圖並存成圖片檔
        chart_filename = f"{symbol}_chart.png"
        mpf.plot(hist, type='candle', style='yahoo', volume=True, 
                 mav=(5, 20), title=f"{name} ({symbol})", 
                 savefig=chart_filename)
        
        # 數據計算
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
        
        # 🔍 情報抓取：取得最新一則相關新聞
        latest_news = None
        try:
            news_list = stock.news
            if news_list and len(news_list) > 0:
                latest_news = news_list[0] # 只拿最新的一筆，避免洗版
        except:
            pass # 如果抓不到新聞就算了，不影響主程式運行
        
        return td, yd, last_trade_date, chart_filename, latest_news
    except Exception as e:
        print(f"[{symbol}] 分析或繪圖失敗: {e}")
        return None

# === 4. Telegram 發送模組 ===
def send_telegram_msg(msg):
    url = f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage"
    # 設定 disable_web_page_preview=True，避免 Telegram 因為網址自動產生一大堆預覽圖洗版
    payload = {"chat_id": TG_CHAT_ID, "text": msg, "disable_web_page_preview": True}
    requests.post(url, json=payload)

# === 5. Email 發送模組 ===
def send_email_report(subject, text_body, image_files):
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
        print("✅ Email 戰情報告發送成功！")
    except Exception as e:
        print(f"❌ Email 發送失敗: {e}")

# === 6. 執行主程式 ===
# 設定手錶為台灣時間 (UTC+8)
tw_tz = datetime.timezone(datetime.timedelta(hours=8))
current_tw_date = datetime.datetime.now(tw_tz).date()

message_list = []
generated_charts = []
has_new_data = False 

print(f"今天是 {current_tw_date}，開始執行 NOC 情報戰情室掃描...")

for symbol, name in STOCK_DICT.items():
    data = get_analysis_and_draw_chart(symbol, name)
    
    if data is not None:
        td, yd, last_trade_date, chart_file, news = data
        
        # --- 🛡️ 假日防禦邏輯 ---
        is_market_open = False
        if ".TW" in symbol:
            if last_trade_date == current_tw_date: is_market_open = True
        else:
            if last_trade_date >= current_tw_date - datetime.timedelta(days=1): is_market_open = True
                
        if not is_market_open:
            print(f"{name} ({symbol}) 今日無新交易數據，判定為休市。")
            continue 
            
        has_new_data = True 
        generated_charts.append(chart_file) 
        
        # --- 📊 取得指標數據 ---
        rsi_value = td['RSI']
        vol_today = td['Volume']
        vma5 = td['5VMA']
        
        td_strong = td['Close'] > td['5MA'] > td['20MA']
        yd_strong = yd['Close'] > yd['5MA'] > yd['20MA']
        td_weak = td['Close'] < td['5MA'] < td['20MA']

        # 量能判定
        if vol_today > vma5 * 1.2: vol_status = "📈 出量 (大於5日均量)"
        elif vol_today < vma5 * 0.8: vol_status = "📉 量縮 (交投清淡)"
        else: vol_status = "➖ 量平 (維持均量)"

        # 狀態判定
        if td_strong: status = "🔥 多頭排列 (趨勢強勢)"
        elif td_weak: status = "🧊 空頭排列 (趨勢疲弱)"
        else: status = "🔄 盤整震盪"
        
        # 訊號判定
        alert = ""
        if td_weak and vol_today > vma5 * 1.2:
            alert = "💀【恐慌殺盤】跌勢加速且帶量！主力倒貨中，切勿徒手接刀！"
        elif yd['Close'] < yd['5MA'] and td['Close'] > td['5MA']:
            if vol_today > vma5 * 1.2: alert = "🚀【強力反轉】帶量站回5日線！系統重啟，底部轉強訊號！"
            else: alert = "📈【弱勢反彈】站回5日線但量能不足，暫視為反彈。"
        elif yd['5MA'] < yd['20MA'] and td['5MA'] > td['20MA']:
            alert = "🌟【黃金交叉】長短線趨勢翻轉向上！"
        elif yd_strong and not td_strong:
            alert = "⚠️【警戒】趨勢由強轉弱，支撐失守，建議減碼觀察。"
        elif rsi_value < 30:
            alert = "🟢【超跌】RSI低於30，隨時可能跌深反彈。"
        elif rsi_value > 70:
            alert = "🔴【過熱】RSI高於70，系統過熱，隨時有修正風險。"
        else:
            alert = "✅ 狀態穩定，目前無特殊轉折訊號。"

        # --- D. 單檔股票排版 (新增情報推播) ---
        stock_msg = f"🔸 {name} ({symbol})\n"
        stock_msg += f"   現價: {td['Close']:.2f} | RSI: {rsi_value:.1f}\n"
        stock_msg += f"   量能: {vol_status}\n"
        stock_msg += f"   狀態: {status}\n"
        stock_msg += f"   訊號: {alert}\n"
        
        # 如果有抓到新聞，就把它接在最下面
        if news:
            news_title = news.get("title", "無標題")
            news_link = news.get("link", "")
            stock_msg += f"   📰 情報: {news_title}\n"
            stock_msg += f"   🔗 連結: {news_link}\n"
            
        stock_msg += "\n"
        message_list.append(stock_msg)
        print(f"{name} 掃描完成")

# === 最終檢查：發送 Telegram 與 Email ===
if has_new_data and len(message_list) > 0:
    final_text = "📡 【老網管 NOC 戰情室：全方位轉折與情報預警】\n\n" + "".join(message_list)
    final_text += "老網管提醒：消息面僅供輔助，請以量價結構為主！詳細線圖已寄至 Email！🛡️"
    
    send_telegram_msg(final_text)
    send_email_report(f"📊 理財儀表板戰情日報 ({current_tw_date})", final_text, generated_charts)
else:
    print("😴 今日台美股均判定為休市(或國定假日)，暫停發送報告。")
