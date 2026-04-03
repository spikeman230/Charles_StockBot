import yfinance as yf
import requests
import os

# === 1. 從 GitHub 保險箱抓取機密 ===
TG_TOKEN = os.environ.get("TG_TOKEN")
TG_CHAT_ID = os.environ.get("TG_CHAT_ID")

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

# === 3. 旗艦分析模組：包含均線與 RSI 計算 ===
def get_advanced_analysis(symbol):
    try:
        stock = yf.Ticker(symbol)
        # 為了算 RSI，我們改抓 6 個月的資料比較準確
        hist = stock.history(period="6mo")
        if len(hist) < 30: return None 
        
        # 計算 5MA 與 20MA
        hist['5MA'] = hist['Close'].rolling(window=5).mean()
        hist['20MA'] = hist['Close'].rolling(window=20).mean()
        
        # 計算 14日 RSI (經典參數)
        delta = hist['Close'].diff()
        gain = delta.clip(lower=0)
        loss = -1 * delta.clip(upper=0)
        # 使用指數移動平均計算 RSI
        ema_gain = gain.ewm(com=13, adjust=False).mean()
        ema_loss = loss.ewm(com=13, adjust=False).mean()
        rs = ema_gain / ema_loss
        hist['RSI'] = 100 - (100 / (1 + rs))
        
        today = hist.iloc[-1]
        yesterday = hist.iloc[-2]
        
        return today, yesterday
    except Exception as e:
        return None

# === 4. 發送 Telegram 通知模組 ===
def send_telegram_msg(msg):
    url = f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage"
    payload = {"chat_id": TG_CHAT_ID, "text": msg}
    requests.post(url, json=payload)

# === 5. 執行主程式：抄底與避險全監控 ===
message = "📡 【老網管理財雷達：轉折與抄底監控】\n\n"
print("開始執行 RSI 超跌監控...")

for symbol, name in STOCK_DICT.items():
    data = get_advanced_analysis(symbol)
    
    if data is not None:
        td, yd = data
        
        # 判斷 RSI 狀態
        rsi_value = td['RSI']
        
        # 判斷趨勢預警
        is_strong_today = td['Close'] > td['5MA'] > td['20MA']
        is_strong_yesterday = yd['Close'] > yd['5MA'] > yd['20MA']
        
        alert = ""
        # 1. 抄底訊號 (優先顯示)
        if rsi_value < 30:
            alert = f"🟢【超跌買點】RSI 降至 {rsi_value:.1f}，市場極度恐慌，可關注抄底機會！"
        elif rsi_value > 70:
            alert = f"🔴【過熱警報】RSI 飆至 {rsi_value:.1f}，隨時可能回檔，請勿追高！"
        # 2. 轉弱預警
        elif is_strong_yesterday and not is_strong_today:
            alert = "⚠️【由強轉弱】跌破短線支撐，請注意風險。"
            
        # 組合訊息
        message += f"🔸 {name} ({symbol})\n"
        message += f"   現價: {td['Close']:.2f} | RSI: {rsi_value:.1f}\n"
        
        if alert:
            message += f"   🎯 訊號: {alert}\n"
        else:
            message += f"   🎯 訊號: 數值正常，穩定運行中。\n"
            
        message += "\n"
        print(f"{name} 掃描完成")
    else:
        message += f"🔸 {name} ({symbol}) : 資料抓取失敗 ❌\n\n"

message += "老網管碎碎念：超跌有時候還會更跌，RSI 低於 30 只是把它放入『觀察名單』，千萬不要一次把資金 All-in 啊！🛡️"

try:
    send_telegram_msg(message)
    print("✅ 抄底預警報告發送成功！")
except Exception as e:
    print(f"❌ 傳送 Telegram 失敗: {e}")
