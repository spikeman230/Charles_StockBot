import yfinance as yf
import requests
import os

# === 1. 從 GitHub 保險箱抓取機密 (資安 100 分！) ===
STOCK_SYMBOL = "2330.TW"
TG_TOKEN = os.environ.get("TG_TOKEN")
TG_CHAT_ID = os.environ.get("TG_CHAT_ID")

# === 2. 抓取目前股價 ===
def get_stock_price(symbol):
    stock = yf.Ticker(symbol)
    current_price = stock.history(period="1d")['Close'].iloc[-1]
    return current_price

# === 3. 發送 Telegram 通知 ===
def send_telegram_msg(msg):
    url = f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage"
    payload = {"chat_id": TG_CHAT_ID, "text": msg}
    requests.post(url, json=payload)

# === 4. 執行主程式 ===
try:
    price = get_stock_price(STOCK_SYMBOL)
    message = f"📊 每日早報：老闆早安！目前 {STOCK_SYMBOL} 股價為 {price:.2f} 元！"
    send_telegram_msg(message)
    print("✅ 報告發送成功！")
except Exception as e:
    print(f"❌ 抓取失敗: {e}")
