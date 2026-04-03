import yfinance as yf
import requests
import os

# === 1. 從 GitHub 保險箱抓取機密 ===
TG_TOKEN = os.environ.get("TG_TOKEN")
TG_CHAT_ID = os.environ.get("TG_CHAT_ID")

# === 2. 你的「專屬通訊錄」 (這在程式裡叫做 Dictionary 字典) ===
# 左邊是代碼，右邊是你想要顯示的名稱。想怎麼取名都可以！
STOCK_DICT = {
    "2330.TW": "台積電",
    "2317.TW": "鴻海",
    "0050.TW": "元大台灣50",
    "009816.TW": "凱基台灣TOP50",  # 補上剛才修正的正確代碼
    "8431.TWO": "匯鑽科",
    "AAPL": "蘋果 (Apple)",
    "NVDA": "輝達 (NVIDIA)"
}

# === 3. 抓取目前股價模組 ===
def get_stock_price(symbol):
    try:
        stock = yf.Ticker(symbol)
        current_price = stock.history(period="1d")['Close'].iloc[-1]
        return current_price
    except Exception as e:
        return None

# === 4. 發送 Telegram 通知模組 ===
def send_telegram_msg(msg):
    url = f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage"
    payload = {"chat_id": TG_CHAT_ID, "text": msg}
    requests.post(url, json=payload)

# === 5. 執行主程式：組合你的專屬早報 ===
message = "📊 【每日早報】老闆早安！您的關注清單現價如下：\n\n"

print("開始抓取多檔股票...")

# 注意這裡！我們把字典裡的「代碼(symbol)」跟「名稱(name)」一起抓出來用
for symbol, name in STOCK_DICT.items():
    price = get_stock_price(symbol)
    
    if price is not None:
        # 顯示格式升級：名稱 + (代碼) + 價格
        message += f"🔸 {name} ({symbol}) : {price:.2f}\n"
        print(f"{name} 抓取成功")
    else:
        message += f"🔸 {name} ({symbol}) : 抓取失敗 ❌\n"
        print(f"{name} 抓取失敗")

message += "\n祝您今天投資順利！💸"

# 最後發送
try:
    send_telegram_msg(message)
    print("✅ 升級版早報發送成功！")
except Exception as e:
    print(f"❌ 傳送 Telegram 失敗: {e}")
