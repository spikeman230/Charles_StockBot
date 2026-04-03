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

# === 3. 旗艦分析模組 (抓取歷史數據並計算指標) ===
def get_advanced_analysis(symbol):
    try:
        stock = yf.Ticker(symbol)
        hist = stock.history(period="6mo")
        if len(hist) < 30: return None 
        
        # 計算 5MA 與 20MA
        hist['5MA'] = hist['Close'].rolling(window=5).mean()
        hist['20MA'] = hist['Close'].rolling(window=20).mean()
        
        # 計算 14日 RSI
        delta = hist['Close'].diff()
        gain = delta.clip(lower=0)
        loss = -1 * delta.clip(upper=0)
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

# === 5. 執行主程式：狀態與訊號全合併 ===
message = "📡 【老網管終極儀表板：趨勢與轉折監控】\n\n"
print("開始執行全方位掃描...")

for symbol, name in STOCK_DICT.items():
    data = get_advanced_analysis(symbol)
    
    if data is not None:
        td, yd = data
        rsi_value = td['RSI']
        
        # --- A. 判斷【狀態】(均線多空排列) ---
        is_strong_today = td['Close'] > td['5MA'] > td['20MA']
        is_strong_yesterday = yd['Close'] > yd['5MA'] > yd['20MA']
        
        if is_strong_today:
            status = "🔥 趨勢仍強 (多頭排列)"
        elif td['Close'] < td['5MA'] < td['20MA']:
            status = "🧊 趨勢疲弱 (空頭排列)"
        else:
            status = "🔄 盤整震盪中"

        # --- B. 判斷【訊號】(RSI 與 轉折預警) ---
        alert = ""
        # 1. RSI 抄底或過熱 (優先級最高)
        if rsi_value < 30:
            alert = "🟢【超跌買點】市場極度恐慌，關注抄底機會！"
        elif rsi_value > 70:
            alert = "🔴【過熱警報】隨時可能回檔，請勿盲目追高！"
        # 2. 由強轉弱預警
        elif is_strong_yesterday and not is_strong_today:
            if td['Close'] < td['5MA']:
                alert = "⚠️【由強轉弱】跌破5日線，短線支撐轉弱！"
            if td['5MA'] < td['20MA'] and yd['5MA'] >= yd['20MA']:
                alert = "❌【危險訊號】發生死亡交叉！趨勢可能反轉！"
        # 3. 無特殊狀況
        else:
            alert = "✅ 數值正常，穩定運行中"
            
        # --- C. 組合完美的排版訊息 ---
        message += f"🔸 {name} ({symbol})\n"
        message += f"   現價: {td['Close']:.2f} | RSI: {rsi_value:.1f}\n"
        message += f"   狀態: {status}\n"
        message += f"   訊號: {alert}\n\n"
        
        print(f"{name} 掃描完成")
    else:
        message += f"🔸 {name} ({symbol}) : 資料抓取失敗 ❌\n\n"

message += "老網管提醒：指標皆為歷史數據計算，請搭配實體量能與市場消息綜合判斷！🛡️"

try:
    send_telegram_msg(message)
    print("✅ 終極報告發送成功！")
except Exception as e:
    print(f"❌ 傳送 Telegram 失敗: {e}")
