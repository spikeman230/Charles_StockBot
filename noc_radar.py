import yfinance as yf
import requests
import datetime
import pandas as pd
import os

# === 1. 設定掃描池 (可自行增減) ===
SCAN_LIST = [
    "2330.TW", "2317.TW", "2454.TW", "2382.TW", "2308.TW", 
    "2881.TW", "2882.TW", "3231.TW", "3037.TW", "3714.TW",
    "2603.TW", "2609.TW", "2615.TW", "2352.TW", "2409.TW",
    "3163.TWO", "5388.TW", "2408.TW", "8299.TWO", "3260.TWO"
]

FINMIND_TOKEN = os.environ.get("FINMIND_TOKEN", "") # 有就用，沒有也沒關係

def get_revenue_yoy(symbol):
    if not FINMIND_TOKEN: return None
    fm_symbol = symbol.replace(".TW", "").replace(".TWO", "")
    try:
        url = "https://api.finmindtrade.com/api/v4/data"
        params = {"dataset": "TaiwanStockMonthRevenue", "data_id": fm_symbol, 
                  "start_date": (datetime.datetime.now() - datetime.timedelta(days=400)).strftime("%Y-%m-%d"), 
                  "token": FINMIND_TOKEN}
        data = requests.get(url, params=params, timeout=5).json()
        if data.get("msg") == "success" and len(data.get("data", [])) > 0:
            df = pd.DataFrame(data["data"])
            latest = df.iloc[-1]
            last_year = df[(df['revenue_year'] == latest['revenue_year'] - 1) & (df['revenue_month'] == latest['revenue_month'])]
            if not last_year.empty and last_year.iloc[-1]['revenue'] > 0:
                return ((latest['revenue'] - last_year.iloc[-1]['revenue']) / last_year.iloc[-1]['revenue']) * 100
    except: pass
    return None

def scan_stock(symbol):
    try:
        hist = yf.Ticker(symbol).history(period="3mo").dropna(subset=['Close'])
        if len(hist) < 30: return None
        
        # 計算均線與均量
        hist['20MA'] = hist['Close'].rolling(20).mean()
        hist['5VMA'] = hist['Volume'].rolling(5).mean()
        
        # KD 計算 (9,3,3)
        low_9 = hist['Low'].rolling(9).min()
        high_9 = hist['High'].rolling(9).max()
        hist['K'] = (((hist['Close'] - low_9) / (high_9 - low_9)) * 100).ewm(com=2, adjust=False).mean()
        hist['D'] = hist['K'].ewm(com=2, adjust=False).mean()
        
        td = hist.iloc[-1]
        y_td = hist.iloc[-2] # 昨天
        
        # === 四合一黃金條件判斷 ===
        cond_1 = td['Close'] > td['20MA'] # 站上月線
        cond_2 = td['Volume'] > (td['5VMA'] * 1.2) # 底部出量
        cond_3 = td['K'] < 35 and td['K'] > td['D'] and y_td['K'] <= y_td['D'] # KD低檔金叉 (放寬至35方便掃描)
        
        if cond_1 and cond_2 and cond_3:
            yoy = get_revenue_yoy(symbol)
            if yoy is not None and yoy < 0: return None # 營收衰退者淘汰
            
            yoy_str = f"{yoy:.1f}%" if yoy is not None else "無API資料"
            return {
                "symbol": symbol,
                "close": td['Close'],
                "K": td['K'],
                "D": td['D'],
                "yoy": yoy_str
            }
    except Exception as e:
        return None
    return None

if __name__ == "__main__":
    print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] 🚀 NOC 游擊隊雷達啟動，掃描目標 {len(SCAN_LIST)} 檔...")
    print("=" * 50)
    found_targets = []
    
    for sym in SCAN_LIST:
        print(f"正在掃描 {sym}...", end="\r")
        result = scan_stock(sym)
        if result: found_targets.append(result)
        
    print("=" * 50)
    if not found_targets:
        print("🎯 報告總操盤手，目前無符合【KD低檔金叉 + 站上月線 + 出量 + 營收成長】之標的。")
    else:
        print("🎯 發現符合黃金條件的潛力股：")
        for t in found_targets:
            print(f"🔹 {t['symbol']} | 現價: {t['close']:.1f} | K值: {t['K']:.1f} | 營收YoY: {t['yoy']}")
    print("=" * 50)
