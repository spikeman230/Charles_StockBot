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

# === 3. 核心分析模組 ===
def get_advanced_analysis(symbol):
    try:
        stock = yf.Ticker(symbol)
        hist = stock.history(period="6mo")
        if len(hist) < 30: return None 
        
        hist['5MA'] = hist['Close'].rolling(window=5).mean()
        hist['20MA'] = hist['Close'].rolling(window=20).mean()
        hist['5VMA'] = hist['Volume'].rolling(window=5).mean()
        
        # RSI 計算
        delta = hist['Close'].diff()
        gain = delta.clip(lower=0)
        loss = -1 * delta.clip(upper=0)
        ema_gain = gain.ewm(com=13, adjust=False).mean()
        ema_loss = loss.ewm(com=13, adjust=False).mean()
        rs = ema_gain / ema_loss
        hist['RSI'] = 100 - (100 / (1 + rs))
        
        return hist.iloc[-1], hist.iloc[-2]
    except Exception as e:
        return None

# === 4. Telegram 發送模組 ===
def send_telegram_msg(msg):
    url = f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage"
    payload = {"chat_id": TG_CHAT_ID, "text": msg}
    requests.post(url, json=payload)

# === 5. 執行主程式：轉折偵測邏輯 ===
message = "🚨 【老網管：全方位轉折預警儀表板】\n\n"
print("開始執行深度量價掃描...")

for symbol, name in STOCK_DICT.items():
    data = get_advanced_analysis(symbol)
    if data is None: continue
    
    td, yd = data # 今日與昨日
    vol_up = td['Volume'] > td['5VMA'] * 1.2
    
    # 趨勢判定
    td_strong = td['Close'] > td['5MA'] > td['20MA']
    yd_strong = yd['Close'] > yd['5MA'] > yd['20MA']
    td_weak = td['Close'] < td['5MA'] < td['20MA']
    
    # --- 核心訊號邏輯 ---
    alert = ""
    
    # A. 恐慌殺盤偵測
    if td_weak and vol_up:
        alert = "💀【恐慌殺盤】跌勢加速且帶量！主力倒貨中，切勿徒手接刀！"
    
    # B. 由弱轉強 (反轉) 偵測
    elif yd['Close'] < yd['5MA'] and td['Close'] > td['5MA']:
        if vol_up:
            alert = "🚀【強力反轉】帶量站回5日線！系統重啟，底部轉強訊號！"
        else:
            alert = "📈【弱勢反彈】站回5日線但量能不足，暫視為反彈。"
            
    # C. 黃金交叉偵測
    elif yd['5MA'] < yd['20MA'] and td['5MA'] > td['20MA']:
        alert = "🌟【黃金交叉】長短線趋势翻轉向上！"
        
    # D. 由強轉弱預警
    elif yd_strong and not td_strong:
        alert = "⚠️【警戒】趨勢由強轉弱，支撐失守，建議減碼觀察。"
        
    # E. RSI 超跌/過熱 (保底監控)
    elif td['RSI'] < 30:
        alert = "🟢【超跌】RSI低於30，雖然還在跌，但隨時可能跌深反彈。"
    elif td['RSI'] > 70:
        alert = "🔴【過熱】RSI高於70，系統過熱，隨時有修正風險。"
    else:
        alert = "✅ 狀態穩定，目前無特殊轉折訊號。"

    # 排版
    message += f"🔸 {name} ({symbol})\n"
    message += f"   現價: {td['Close']:.2f} | RSI: {td['RSI']:.1f}\n"
    message += f"   訊號: {alert}\n\n"

message += "老網管碎碎念：看到 💀 跑得快，看到 🚀 慢慢買。轉折點是致富關鍵，也是破產開端，請冷靜執行！🛡️"

try:
    send_telegram_msg(message)
    print("✅ 終極轉折報告發送成功！")
except Exception as e:
    print(f"❌ 傳送失敗: {e}")
