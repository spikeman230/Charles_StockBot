import yfinance as yf
import requests
import os
import datetime
import pandas as pd
import numpy as np
import csv
import mplfinance as mpf

# === 1. 從 GitHub 保險箱抓取機密 ===
TG_TOKEN = os.environ.get("TG_TOKEN")
TG_CHAT_ID = os.environ.get("TG_CHAT_ID")

# === 2. 專屬通訊錄 ===
STOCK_DICT = {
    "🛡️ 核心持股 (重倉伺服器)": {"3037.TW": "欣興 (ABF載板)"},
    "🔥 潛力種子 (高頻寬觀察區)": {
        "3163.TW": "波若威 (光通訊)",
        "5388.TW": "中磊 (網通設備)",
        "3714.TW": "富采 (LED光電)"
    },
    "👀 常態觀察區 (例行監控節點)": {
        "2330.TW": "台積電", "2317.TW": "鴻海", "0050.TW": "元大台灣50",
        "AAPL": "蘋果 (Apple)", "NVDA": "輝達 (NVIDIA)"
    },
    "💾 記憶體族群 (美光連動網域)": {
        "MU": "美光 (Micron)", "2408.TW": "南亞科", "2344.TW": "華邦電",
        "6239.TW": "力成", "3260.TW": "威剛", "8299.TW": "群聯", "4967.TW": "十銓"
    }
}

# === 3. 數據與訊號模組 ===
def write_noc_log(date, symbol, name, close_price, rsi, vol_status, status, alert):
    log_filename = "noc_trading_log.csv"
    file_exists = os.path.exists(log_filename)
    with open(log_filename, mode='a', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["日期", "代號", "名稱", "收盤價", "RSI", "量能狀態", "趨勢狀態", "戰情室指令"])
        writer.writerow([date, symbol, name, f"{close_price:.2f}", f"{rsi:.2f}", vol_status, status, alert])

def get_analysis(symbol, name):
    try:
        stock = yf.Ticker(symbol)
        hist = stock.history(period="6mo")
        if len(hist) < 30: return None

        # 計算指標
        hist['5MA'] = hist['Close'].rolling(window=5).mean()
        hist['20MA'] = hist['Close'].rolling(window=20).mean()
        hist['5VMA'] = hist['Volume'].rolling(window=5).mean()
        
        delta = hist['Close'].diff()
        gain = delta.clip(lower=0); loss = -1 * delta.clip(upper=0)
        ema_gain = gain.ewm(com=13, adjust=False).mean()
        ema_loss = loss.ewm(com=13, adjust=False).mean()
        hist['RSI'] = 100 - (100 / (1 + (ema_gain / ema_loss)))

        td = hist.iloc[-1]; yd = hist.iloc[-2]
        
        # 抓新聞
        news = stock.news[0] if stock.news else None
        return td, yd, hist.index[-1].date(), news
    except: return None

def send_telegram_msg(msg):
    url = f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage"
    requests.post(url, json={"chat_id": TG_CHAT_ID, "text": msg, "disable_web_page_preview": True})

# === 4. 執行主程式 ===
if __name__ == "__main__":
    tw_tz = datetime.timezone(datetime.timedelta(hours=8))
    curr_date = datetime.datetime.now(tw_tz).date()
    curr_time = datetime.datetime.now(tw_tz).strftime("%Y-%m-%d %H:%M:%S")
    
    msg_list = []; has_data = False
    print(f"[{curr_time}] NOC 戰情室啟動...")

    for cat, stocks in STOCK_DICT.items():
        msg_list.append(f"━━━━━━━━━━━━━━\n📂 【{cat}】\n━━━━━━━━━━━━━━\n")
        for sym, name in stocks.items():
            res = get_analysis(sym, name)
            if not res: continue
            td, yd, last_date, news = res
            
            # 假日防禦 (只在有新資料時執行)
            if last_date != curr_date and ".TW" in sym: continue
            has_data = True

            # --- 強化版暴力防呆邏輯 ---
            vol_today = td['Volume']; vma5 = td['5VMA']
            td_strong = td['Close'] > td['5MA'] > td['20MA']
            yd_strong = yd['Close'] > yd['5MA'] > yd['20MA']
            td_weak = td['Close'] < td['5MA'] < td['20MA']

            vol_status = "📈 出量 (大於5日均量)" if vol_today > vma5 * 1.2 else "📉 量縮" if vol_today < vma5 * 0.8 else "➖ 量平"
            trend_status = "🔥 多頭排列 (強勢)" if td_strong else "🧊 空頭排列 (疲弱)" if td_weak else "🔄 盤整震盪"

            if td_weak and vol_today > vma5 * 1.2:
                alert = "💀【強制退場】大單狂砸！立刻清倉停損，拔掉網路線保命！"
            elif yd['Close'] < yd['5MA'] and td['Close'] > td['5MA']:
                alert = "🚀【強烈買進】流量爆發站回5日線！立刻進場！" if vol_today > vma5 * 1.2 else "📈【試探買進】站回5日線但量縮。"
            elif td['RSI'] > 80:
                alert = "💰【獲利了結】RSI過熱！不要標，立刻賣出一半入袋為安！"
            elif td['RSI'] < 30:
                alert = "🟢【準備抄底】RSI超賣。先加進觀察名單！"
            else:
                alert = "✅【持股續抱】目前無轉折。空手別追，有持股就繼續抱著！"

            # 寫入日誌
            write_noc_log(curr_date, sym, name, td['Close'], td['RSI'], vol_status, trend_status, alert)

            # --- 還原排版格式 (加入空格縮排) ---
            stock_msg = f"🔸 {name} ({sym})\n"
            stock_msg += f"   現價: {td['Close']:.2f} | RSI: {td['RSI']:.1f}\n"
            stock_msg += f"   量能: {vol_status}\n"
            stock_msg += f"   狀態: {trend_status}\n"
            stock_msg += f"   👉 指令: {alert}\n"
            if news:
                stock_msg += f"   📰 情報: {news.get('title', '無')}\n"
                stock_msg += f"   🔗 連結: {news.get('link', '')}\n"
            msg_list.append(stock_msg + "\n")

    if has_data:
        final_text = f"📡 【老網管 NOC 指揮中心：行動清單】\n📅 系統時間：{curr_time}\n━━━━━━━━━━━━━━\n" + "".join(msg_list)
        final_text += "⚠️ 老網管提醒：收到指令請馬上動作，猶豫就會敗北！"
        send_telegram_msg(final_text)
    else:
        send_telegram_msg(f"📡 【NOC 休市回報】\n📅 {curr_time}\n😴 報告：今日台股休市，伺服器待命。")
