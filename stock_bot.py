import yfinance as yf
import requests
import os

# === 1. 從 GitHub 保險箱抓取機密 ===
TG_TOKEN = os.environ.get("TG_TOKEN")
TG_CHAT_ID = os.environ.get("TG_CHAT_ID")

# === 2. 專屬通訊錄 (可自行增減) ===
STOCK_DICT = {
    "2330.TW": "台積電",
    "2317.TW": "鴻海",
    "0050.TW": "元大台灣50",
    "009816.TW": "凱基台灣TOP50",
    "8431.TWO": "匯鑽科",
    "AAPL": "蘋果 (Apple)",
    "NVDA": "輝達 (NVIDIA)"

}

# === 3. 進階抓取與技術分析模組 ===
def get_stock_analysis(symbol):
    try:
        stock = yf.Ticker(symbol)
        # 抓取過去 3 個月的歷史資料，確保有足夠的天數來算 20日線
        hist = stock.history(period="3mo")
        
        if hist.empty:
            return None
        
        # 取得最新收盤價
        current_price = hist['Close'].iloc[-1]
        
        # 讓程式自己算移動平均線 (MA)
        # rolling(window=5).mean() 意思是：把過去 5 天的數字加起來除以 5
        hist['5MA'] = hist['Close'].rolling(window=5).mean()
        hist['20MA'] = hist['Close'].rolling(window=20).mean()
        
        # 抓出最新一天的 MA 數值
        ma5 = hist['5MA'].iloc[-1]
        ma20 = hist['20MA'].iloc[-1]
        
        return current_price, ma5, ma20
    except Exception as e:
        return None

# === 4. 發送 Telegram 通知模組 ===
def send_telegram_msg(msg):
    url = f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage"
    payload = {"chat_id": TG_CHAT_ID, "text": msg}
    requests.post(url, json=payload)

# === 5. 執行主程式：組合技術面早報 ===
message = "📊 【進階技術面早報】老闆早安！\n\n"

print("開始執行技術分析掃描...")

for symbol, name in STOCK_DICT.items():
    data = get_stock_analysis(symbol)
    
    if data is not None:
        price, ma5, ma20 = data
        
        # 老網管寫的簡單趨勢判斷邏輯
        if price > ma5 and ma5 > ma20:
            trend = "🔥 強勢多頭 (站上均線)"
        elif price < ma5 and ma5 < ma20:
            trend = "🧊 弱勢空頭 (跌破均線)"
        else:
            trend = "🔄 盤整震盪"
            
        # 將計算結果排版並加入訊息中
        message += f"🔸 {name} ({symbol})\n"
        message += f"   現價: {price:.2f}\n"
        message += f"   5日線(週): {ma5:.2f} | 20日線(月): {ma20:.2f}\n"
        message += f"   趨勢: {trend}\n\n"
        print(f"{name} 分析完成")
    else:
        message += f"🔸 {name} ({symbol}) : 資料抓取失敗 ❌\n\n"
        print(f"{name} 抓取失敗")

message += "提醒：技術指標由程式自動運算，僅供參考，請謹慎投資！💸"

# 最後發送
try:
    send_telegram_msg(message)
    print("✅ 技術分析早報發送成功！")
except Exception as e:
    print(f"❌ 傳送 Telegram 失敗: {e}")
