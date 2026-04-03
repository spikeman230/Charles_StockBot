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

# === 3. 終極分析模組 (均線 + RSI + 量能) ===
def get_advanced_analysis(symbol):
    try:
        stock = yf.Ticker(symbol)
        hist = stock.history(period="6mo")
        if len(hist) < 30: return None 
        
        # 價格均線 (MA)
        hist['5MA'] = hist['Close'].rolling(window=5).mean()
        hist['20MA'] = hist['Close'].rolling(window=20).mean()
        
        # 量能均線 (VMA)
        hist['5VMA'] = hist['Volume'].rolling(window=5).mean()
        
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

# === 5. 執行主程式 ===
message = "📡 【老網管終極儀表板：全方位量價監控】\n\n"
print("開始執行全方位量價掃描...")

for symbol, name in STOCK_DICT.items():
    data = get_advanced_analysis(symbol)
    
    if data is not None:
        td, yd = data
        rsi_value = td['RSI']
        vol_today = td['Volume']
        vma5 = td['5VMA']
        
        # --- A. 判斷【量能】 ---
        # 設定超過均量 1.2 倍叫出量，低於 0.8 倍叫量縮
        if vol_today > vma5 * 1.2:
            vol_status = "📈 出量 (動能充沛，大於5日均量)"
        elif vol_today < vma5 * 0.8:
            vol_status = "📉 量縮 (交投清淡，觀望濃厚)"
        else:
            vol_status = "➖ 量平 (維持近期平均動能)"

        # --- B. 判斷【狀態】(均線多空排列) ---
        is_strong_today = td['Close'] > td['5MA'] > td['20MA']
        is_strong_yesterday = yd['Close'] > yd['5MA'] > yd['20MA']
        
        if is_strong_today:
            status = "🔥 多頭排列 (趨勢強勢)"
        elif td['Close'] < td['5MA'] < td['20MA']:
            status = "🧊 空頭排列 (趨勢疲弱)"
        else:
            status = "🔄 盤整震盪"

        # --- C. 判斷【訊號】(綜合研判) ---
        alert = ""
        if rsi_value < 30:
            alert = "🟢【超跌買點】RSI 極度恐慌，關注反彈契機！"
        elif rsi_value > 70:
            alert = "🔴【過熱警報】RSI 進入超買，注意獲利了結賣壓！"
        elif is_strong_yesterday and not is_strong_today:
            alert = "⚠️【由強轉弱】跌破5日線，短線支撐失守！"
        else:
            # 結合量能給出常態建議
            if is_strong_today and vol_today > vma5 * 1.2:
                alert = "⭐【價量齊揚】多頭上攻且帶量，健康輪動！"
            else:
                alert = "✅ 數值正常，穩定運行中"
            
        # --- D. 組合完美的排版訊息 ---
        message += f"🔸 {name} ({symbol})\n"
        message += f"   現價: {td['Close']:.2f} | RSI: {rsi_value:.1f}\n"
        message += f"   量能: {vol_status}\n"
        message += f"   狀態: {status}\n"
        message += f"   訊號: {alert}\n\n"
        
        print(f"{name} 掃描完成")
    else:
        message += f"🔸 {name} ({symbol}) : 資料抓取失敗 ❌\n\n"

message += "老網管提醒：量價結構是市場最真實的足跡，搭配服用效果最佳！🛡️"

try:
    send_telegram_msg(message)
    print("✅ 終極報告發送成功！")
except Exception as e:
    print(f"❌ 傳送 Telegram 失敗: {e}")
