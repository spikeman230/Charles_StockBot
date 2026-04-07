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
        "MU"  : "美光 (Micron)", 
        "2408.TW": "南亞科 (DRAM製造)",
        "2344.TW": "華邦電 (利基記憶體)",
        "6239.TW": "力成 (記憶體封測)", 
        "3260.TW": "威剛 (記憶體模組)", 
        "8299.TW": "群聯 (Flash控制IC)", 
        "4967.TW": "十銓 (電競模組)"
    }
}

# === 3. 分析、繪圖與情報抓取模組 ===
def get_analysis_and_draw_chart(symbol, name):
    try:
        stock = yf.Ticker(symbol)
        hist = stock.history(period="6mo")
        if len(hist) < 30: return None
        
        chart_filename = f"{symbol}_chart.png"
        mpf.plot(hist, type='candle', style='yahoo', volume=True, 
                 mav=(5, 20), title=f"{name} ({symbol})", 
                 savefig=chart_filename)
        
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

# === 4. Telegram 發送模組 ===
def send_telegram_msg(msg):
    url = f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage"
    payload = {"chat_id": TG_CHAT_ID, "text": msg, "disable_web_page_preview": True}
    requests.post(url, json=payload)

# === 5. Email 發送模組 (若不需要可略過設定) ===
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
        print("✅ Email 發送成功！")
    except Exception as e:
        print(f"❌ Email 發送失敗: {e}")

# === 6. 執行主程式 ===
if __name__ == "__main__":
    tw_tz = datetime.timezone(datetime.timedelta(hours=8))
    current_tw_date = datetime.datetime.now(tw_tz).date()

    message_list = []
    generated_charts = []
    has_new_data = False 

    print(f"今天是 {current_tw_date}，開始執行 NOC 暴力防呆戰情室 (v2.0 量能升級版)...")

    for category, stocks in STOCK_DICT.items():
        message_list.append(f"━━━━━━━━━━━━━━\n📂 【{category}】\n━━━━━━━━━━━━━━\n")
        
        for symbol, name in stocks.items():
            data = get_analysis_and_draw_chart(symbol, name)
            
            if data is not None:
                td, yd, last_trade_date, chart_file, news = data
                
                # 假日防禦
                is_market_open = False
                if ".TW" in symbol and last_trade_date == current_tw_date: is_market_open = True
                elif ".TW" not in symbol and last_trade_date >= current_tw_date - datetime.timedelta(days=1): is_market_open = True
                        
                if not is_market_open:
                    continue 
                    
                has_new_data = True 
                generated_charts.append(chart_file) 
                
                # 📊 取得指標數據
                rsi_value = td['RSI']
                vol_today = td['Volume']
                vma5 = td['5VMA']
                
                td_strong = td['Close'] > td['5MA'] > td['20MA']
                yd_strong = yd['Close'] > yd['5MA'] > yd['20MA']
                td_weak = td['Close'] < td['5MA'] < td['20MA']

                # 🚀 Phase 1 升級：雙重嚴格量能判定
                if vol_today > vma5 * 2.0: vol_status = "🌋 爆量 (大於5日均量2倍以上)"
                elif vol_today > vma5 * 1.2: vol_status = "📈 出量 (溫和放量)"
                elif vol_today < vma5 * 0.8: vol_status = "📉 量縮 (交投清淡)"
                else: vol_status = "➖ 量平 (維持均量)"

                if td_strong: status = "🔥 多頭排列 (趨勢強勢)"
                elif td_weak: status = "🧊 空頭排列 (趨勢疲弱)"
                else: status = "🔄 盤整震盪"

                # 🚀 Phase 1 升級：暴力防呆指令邏輯 (加入 2倍量條件)
                alert = ""
                if td_weak and vol_today > vma5 * 2.0:
                    alert = "💀【強制退場】2倍爆量狂砸！主力出貨，立刻清倉拔線！"
                elif td_weak and vol_today > vma5 * 1.2:
                    alert = "⚠️【警戒退場】出量跌破均線！建議減碼防守！"
                elif yd['Close'] < yd['5MA'] and td['Close'] > td['5MA']:
                    if vol_today > vma5 * 2.0:
                        alert = "🚀【強烈買進】2倍爆量站回5日線！超級強勢，立刻進場！"
                    elif vol_today > vma5 * 1.2:
                        alert = "📈【試探買進】溫和出量站回5日線，可小買試單。"
                    else:
                        alert = "📊【觀望買進】站回5日線但量縮。可小買，破線即跑！"
                elif yd['5MA'] < yd['20MA'] and td['5MA'] > td['20MA']:
                    alert = "🌟【加碼買進】短中線黃金交叉！空手者快買，有持股者加碼！"
                elif yd_strong and not td_strong:
                    alert = "✂️【建議減持】跌破5日線熄火！立刻賣出一半鎖定利潤！"
                elif rsi_value < 30:
                    alert = "🟢【準備抄底】RSI超賣。先加進觀察名單，隨時準備進場！"
                elif rsi_value > 70:
                    alert = "💰【獲利了結】RSI過熱！不要貪，立刻賣出一半入袋為安！"
                else:
                    alert = "✅【持股續抱】目前無轉折。空手別追，有持股就繼續抱著！"

                # 📝 完整排版 (包含量能、狀態與新聞)
                stock_msg = f"🔸 {name} ({symbol})\n"
                stock_msg += f"   現價: {td['Close']:.2f} | RSI: {rsi_value:.1f}\n"
                stock_msg += f"   量能: {vol_status}\n"
                stock_msg += f"   狀態: {status}\n"
                stock_msg += f"   👉 指令: {alert}\n"
                
                if news:
                    news_title = news.get("title", "無標題")
                    news_link = news.get("link", "")
                    stock_msg += f"   📰 情報: {news_title}\n"
                    stock_msg += f"   🔗 連結: {news_link}\n"
                    
                stock_msg += "\n"
                message_list.append(stock_msg)

    # 📡 Timestamp 升級與心跳封包 (Heartbeat) 架構
    if has_new_data and len(message_list) > 0:
        final_text = f"📡 【老網管 NOC 指揮中心：行動清單】\n📅 系統時間：{current_tw_date}\n━━━━━━━━━━━━━━\n" + "".join(message_list)
        final_text += "⚠️ 老網管提醒：收到指令請馬上動作，猶豫就會敗北！"
        
        send_telegram_msg(final_text)
    else:
        sleep_msg = f"📡 【老網管 NOC 指揮中心：休市回報】\n📅 系統時間：{current_tw_date}\n😴 報告：今日台股休市，戰情室伺服器進入待命模式 (Standby)。"
        print("今日休市，已發送待命通知至 Telegram。")
        send_telegram_msg(sleep_msg)
