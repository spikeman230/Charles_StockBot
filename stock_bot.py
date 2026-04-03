import yfinance as yf
import requests
import os
import datetime

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
        
        delta = hist['Close'].diff()
        gain = delta.clip(lower=0)
        loss = -1 * delta.clip(upper=0)
        ema_gain = gain.ewm(com=13, adjust=False).mean()
        ema_loss = loss.ewm(com=13, adjust=False).mean()
        rs = ema_gain / ema_loss
        hist['RSI'] = 100 - (100 / (1 + rs))
        
        td = hist.iloc[-1]
        yd = hist.iloc[-2]
        
        # 抓取這筆資料的「真實交易日期」
        last_trade_date = hist.index[-1].date()
        
        return td, yd, last_trade_date
    except Exception as e:
        return None

# === 4. Telegram 發送模組 ===
def send_telegram_msg(msg):
    url = f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage"
    payload = {"chat_id": TG_CHAT_ID, "text": msg}
    requests.post(url, json=payload)

# === 5. 執行主程式 ===
# 設定手錶為台灣時間 (UTC+8)
tw_tz = datetime.timezone(datetime.timedelta(hours=8))
current_tw_date = datetime.datetime.now(tw_tz).date()

message_list = []
has_new_data = False 

print(f"今天是台灣時間 {current_tw_date}，開始執行掃描...")

for symbol, name in STOCK_DICT.items():
    data = get_advanced_analysis(symbol)
    
    if data is not None:
        td, yd, last_trade_date = data
        
        # --- 🛡️ 假日防禦邏輯 ---
        is_market_open = False
        if ".TW" in symbol:
            if last_trade_date == current_tw_date:
                is_market_open = True
        else:
            if last_trade_date >= current_tw_date - datetime.timedelta(days=1):
                is_market_open = True
                
        if not is_market_open:
            print(f"{name} ({symbol}) 今日無新交易數據，判定為休市。")
            continue 
            
        has_new_data = True 
        
        # --- 📊 取得指標數據 ---
        rsi_value = td['RSI']
        vol_today = td['Volume']
        vma5 = td['5VMA']
        
        td_strong = td['Close'] > td['5MA'] > td['20MA']
        yd_strong = yd['Close'] > yd['5MA'] > yd['20MA']
        td_weak = td['Close'] < td['5MA'] < td['20MA']

        # --- A. 量能判定 ---
        if vol_today > vma5 * 1.2: vol_status = "📈 出量 (大於5日均量)"
        elif vol_today < vma5 * 0.8: vol_status = "📉 量縮 (交投清淡)"
        else: vol_status = "➖ 量平 (維持均量)"

        # --- B. 狀態判定 (成功補回！) ---
        if td_strong:
            status = "🔥 多頭排列 (趨勢強勢)"
        elif td_weak:
            status = "🧊 空頭排列 (趨勢疲弱)"
        else:
            status = "🔄 盤整震盪"
        
        # --- C. 訊號判定 ---
        alert = ""
        if td_weak and vol_today > vma5 * 1.2:
            alert = "💀【恐慌殺盤】主力倒貨中，切勿徒手接刀！"
        elif yd['Close'] < yd['5MA'] and td['Close'] > td['5MA']:
            if vol_today > vma5 * 1.2: alert = "🚀【強力反轉】帶量站回5日線！底部轉強訊號！"
            else: alert = "📈【弱勢反彈】站回5日線但量能不足。"
        elif yd['5MA'] < yd['20MA'] and td['5MA'] > td['20MA']:
            alert = "🌟【黃金交叉】長短線趨勢翻轉向上！"
        elif yd_strong and not td_strong:
            alert = "⚠️【警戒】趨勢由強轉弱，支撐失守。"
        elif rsi_value < 30:
            alert = "🟢【超跌】RSI低於30，隨時可能反彈。"
        elif rsi_value > 70:
            alert = "🔴【過熱】RSI高於70，注意修正風險。"
        else:
            alert = "✅ 狀態穩定，無特殊轉折訊號。"

        # --- D. 單檔股票排版 (4行標準格式) ---
        stock_msg = f"🔸 {name} ({symbol})\n"
        stock_msg += f"   現價: {td['Close']:.2f} | RSI: {rsi_value:.1f}\n"
        stock_msg += f"   量能: {vol_status}\n"
        stock_msg += f"   狀態: {status}\n"
        stock_msg += f"   訊號: {alert}\n\n"
        message_list.append(stock_msg)
        print(f"{name} 掃描完成")
    else:
        print(f"{name} 抓取失敗")

# === 最終檢查：發送 Telegram ===
if has_new_data and len(message_list) > 0:
    final_message = "📡 【老網管終極儀表板：全方位轉折預警】\n\n"
    final_message += "".join(message_list)
    final_message += "老網管提醒：休市日防禦機制已上線，假日安心補眠！🛡️"
    
    try:
        send_telegram_msg(final_message)
        print("✅ 交易日報告發送成功！")
    except Exception as e:
        print(f"❌ 傳送失敗: {e}")
else:
    print("😴 今日台美股均判定為休市(或國定假日)，暫停發送報告。")
