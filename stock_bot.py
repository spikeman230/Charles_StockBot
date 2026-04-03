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

# === 3. 進階分析模組：包含今日與昨日對比 ===
def get_advanced_analysis(symbol):
    try:
        stock = yf.Ticker(symbol)
        hist = stock.history(period="3mo")
        if len(hist) < 21: return None # 確保資料夠算 20MA
        
        # 計算技術指標
        hist['5MA'] = hist['Close'].rolling(window=5).mean()
        hist['20MA'] = hist['Close'].rolling(window=20).mean()
        
        # 抓取今天 (Last) 與 昨天 (Last-1) 的資料
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

# === 5. 執行主程式：加入預警邏輯 ===
message = "🚨 【量價與轉折預警報告】\n\n"

print("開始執行全自動趨勢監控...")

for symbol, name in STOCK_DICT.items():
    data = get_advanced_analysis(symbol)
    
    if data is not None:
        td, yd = data # 今日與昨日資料
        
        # 判斷趨勢狀態
        # 強勢定義：股價 > 5MA > 20MA
        is_strong_today = td['Close'] > td['5MA'] > td['20MA']
        is_strong_yesterday = yd['Close'] > yd['5MA'] > yd['20MA']
        
        # 判斷是否有「轉弱預警」
        alert = ""
        if is_strong_yesterday and not is_strong_today:
            if td['Close'] < td['5MA']:
                alert = "⚠️【由強轉弱】股價跌破5日線，短線支撐轉弱！"
            if td['5MA'] < td['20MA'] and yd['5MA'] >= yd['20MA']:
                alert = "❌【危險訊號】發生死亡交叉！趨勢可能正式反轉！"
        
        # 組合訊息
        message += f"🔸 {name} ({symbol})\n"
        message += f"   現價: {td['Close']:.2f}\n"
        
        # 顯示目前狀態
        if is_strong_today:
            status = "🔥 趨勢仍強 (多頭排列)"
        elif td['Close'] < td['5MA'] < td['20MA']:
            status = "🧊 趨勢疲弱 (空頭排列)"
        else:
            status = "🔄 盤整震盪中"
            
        message += f"   狀態: {status}\n"
        
        # 如果有預警訊息，加粗顯示
        if alert:
            message += f"   🛑 預警: {alert}\n"
            
        message += "\n"
        print(f"{name} 掃描完成")
    else:
        message += f"🔸 {name} ({symbol}) : 資料抓取失敗 ❌\n\n"

message += "老網管提醒：預警訊號出現時建議減碼或觀察，切勿盲目追高！🛡️"

try:
    send_telegram_msg(message)
    print("✅ 預警報告發送成功！")
except Exception as e:
    print(f"❌ 傳送 Telegram 失敗: {e}")
