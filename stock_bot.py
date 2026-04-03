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

# === 3. 價量分析模組 (加入成交量防騙機制) ===
def get_stock_analysis(symbol):
    try:
        stock = yf.Ticker(symbol)
        hist = stock.history(period="3mo")
        
        if hist.empty:
            return None
        
        # 取得最新收盤價與今日成交量
        current_price = hist['Close'].iloc[-1]
        current_vol = hist['Volume'].iloc[-1]
        
        # 計算價格均線 (MA)
        hist['5MA'] = hist['Close'].rolling(window=5).mean()
        hist['20MA'] = hist['Close'].rolling(window=20).mean()
        
        # 計算成交量均線 (VMA)：過去 5 天的平均成交量
        hist['5VMA'] = hist['Volume'].rolling(window=5).mean()
        
        ma5 = hist['5MA'].iloc[-1]
        ma20 = hist['20MA'].iloc[-1]
        vma5 = hist['5VMA'].iloc[-1]
        
        return current_price, ma5, ma20, current_vol, vma5
    except Exception as e:
        return None

# === 4. 發送 Telegram 通知模組 ===
def send_telegram_msg(msg):
    url = f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage"
    payload = {"chat_id": TG_CHAT_ID, "text": msg}
    requests.post(url, json=payload)

# === 5. 執行主程式：組合量價早報 ===
message = "📊 【量價分析雷達】老闆早安！\n\n"

print("開始執行量價分析掃描...")

for symbol, name in STOCK_DICT.items():
    data = get_stock_analysis(symbol)
    
    if data is not None:
        price, ma5, ma20, vol, vma5 = data
        
        # 判斷成交量是否放大 (今日成交量 > 5日平均成交量)
        is_vol_up = vol > vma5
        
        # 進階趨勢與防騙線判斷邏輯
        if price > ma5 and ma5 > ma20:
            if is_vol_up:
                trend = "🔥 強勢多頭【量增】(價量齊揚，趨勢明確)"
            else:
                trend = "⚠️ 強勢多頭【量縮】(小心主力騙線假突破)"
                
        elif price < ma5 and ma5 < ma20:
            if is_vol_up:
                trend = "🧊 弱勢空頭【量增】(恐慌殺盤，賣壓沉重)"
            else:
                trend = "📉 弱勢空頭【量縮】(無量陰跌，人氣退潮)"
                
        else:
            trend = "🔄 盤整震盪 (趨勢不明，建議觀望)"
            
        message += f"🔸 {name} ({symbol})\n"
        message += f"   現價: {price:.2f}\n"
        # 為了版面簡潔，我們就不顯示冗長的成交量數字，直接顯示結論
        message += f"   分析: {trend}\n\n"
        print(f"{name} 分析完成")
    else:
        message += f"🔸 {name} ({symbol}) : 資料抓取失敗 ❌\n\n"
        print(f"{name} 抓取失敗")

message += "系統提示：已啟動成交量防禦機制，過濾虛假突破。投資有風險，下單請謹慎！🛡️"

try:
    send_telegram_msg(message)
    print("✅ 量價分析早報發送成功！")
except Exception as e:
    print(f"❌ 傳送 Telegram 失敗: {e}")
