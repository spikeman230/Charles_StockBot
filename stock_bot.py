import yfinance as yf
import requests
import os

# === 1. 從 GitHub 保險箱抓取機密 ===
TG_TOKEN = os.environ.get("TG_TOKEN")
TG_CHAT_ID = os.environ.get("TG_CHAT_ID")

# === 2. 你的「多檔股票」火力展示區 (用中括號包起來，逗號隔開) ===
# 台股記得加 .TW，美股直接寫代碼。你可以無限往下加！
STOCK_LIST = [
    "2330.TW",  # 台積電
    "2454.TW",  # 聯發科
    "2317.TW",  # 鴻海
    "9816.TW",  # 凱基台灣50
    "8431.TW",  # 匯鑽科
    "7733U.TW",  # 台積中信售
    "AAPL",     # 蘋果
    "NVDA"      # 輝達
]

# === 3. 抓取目前股價模組 ===
def get_stock_price(symbol):
    try:
        stock = yf.Ticker(symbol)
        current_price = stock.history(period="1d")['Close'].iloc[-1]
        return current_price
    except Exception as e:
        return None  # 如果某檔股票下市或代碼打錯，回傳 None 避免整個程式崩潰

# === 4. 發送 Telegram 通知模組 ===
def send_telegram_msg(msg):
    url = f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage"
    payload = {"chat_id": TG_CHAT_ID, "text": msg}
    requests.post(url, json=payload)

# === 5. 執行主程式：組合你的專屬早報 ===
# 先建立一個開頭問候語
message = "📊 【每日早報】老闆早安！您的關注清單現價如下：\n\n"

print("開始抓取多檔股票...")

# 讓程式一個一個去查價 (這就是迴圈)
for symbol in STOCK_LIST:
    price = get_stock_price(symbol)
    
    if price is not None:
        # 如果抓取成功，就把這行字「接」在原本的 message 後面
        message += f"🔸 {symbol} : {price:.2f}\n"
        print(f"{symbol} 抓取成功")
    else:
        # 如果抓取失敗，也回報一下
        message += f"🔸 {symbol} : 抓取失敗 ❌\n"
        print(f"{symbol} 抓取失敗")

# 在報告最底下加個結語
message += "\n祝您今天投資順利！💸"

# 最後，把這包整理好的大訊息，一次發送到 Telegram
try:
    send_telegram_msg(message)
    print("✅ 多檔股票報告發送成功！")
except Exception as e:
    print(f"❌ 傳送 Telegram 失敗: {e}")
