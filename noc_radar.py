import yfinance as yf
import requests
import datetime
import pandas as pd
import os
import json  # 必須引入 json 模組來建立傳令兵

# === 1. 設定掃描池 (台股前150大中型權值股 + 產業指標股) ===
SCAN_LIST = [
    # 半導體、電子與 AI 伺服器
    "2330.TW", "2317.TW", "2454.TW", "2382.TW", "2308.TW", "3231.TW", "3037.TW", "2303.TW", 
    "3008.TW", "3034.TW", "3711.TW", "2357.TW", "2395.TW", "2408.TW", "2353.TW", "2356.TW", 
    "2379.TW", "4938.TW", "2301.TW", "2345.TW", "2324.TW", "3661.TW", "6669.TW", "3714.TW", 
    "3163.TWO", "5388.TW", "8299.TWO", "3260.TWO", "2377.TW", "2383.TW", "3017.TW", "2352.TW", 
    "3443.TW", "3529.TW", "3293.TWO", "6488.TWO", "8069.TWO", "6274.TW", "6239.TW", "3044.TW", 
    "2449.TW", "2344.TW", "2409.TW", "3481.TW", "6116.TW", "4958.TW", "6176.TW", "3532.TW", 
    "2371.TW", "2404.TW", "3702.TW", "8046.TW", "5483.TWO", "3105.TWO", "5347.TWO", "6147.TWO", 
    "6214.TW", "2313.TW", "2368.TW", "3013.TW", "3019.TW", "3042.TW", "3324.TW", "3533.TW", 
    "3583.TW", "3653.TW", "4966.TW", "5269.TW", "6269.TW", "6415.TW", "6531.TW", "8016.TW", 
    "8081.TW", "8150.TW",
    # 金融權值股
    "2881.TW", "2882.TW", "2891.TW", "2886.TW", "2884.TW", "2892.TW", "2885.TW", "2880.TW", 
    "2883.TW", "2887.TW", "5871.TW", "2890.TW", "5880.TW", "2801.TW", "2834.TW", "2838.TW", 
    "2845.TW", "2888.TW", "2889.TW", "6005.TW", "2809.TW", "2812.TW", "2858.TW",
    # 航運、傳產、生技與電信
    "2412.TW", "3045.TW", "4904.TW", "2002.TW", "1216.TW", "1301.TW", "1303.TW", "1326.TW", 
    "2912.TW", "9904.TW", "2603.TW", "2609.TW", "2615.TW", "2207.TW", "1101.TW", "1102.TW", 
    "1229.TW", "1402.TW", "1504.TW", "1513.TW", "1514.TW", "1519.TW", "1590.TW", "1605.TW", 
    "2105.TW", "2606.TW", "2610.TW", "2618.TW", "5522.TW", "8464.TW", "9910.TW", "9914.TW", 
    "9921.TW", "9941.TW", "1108.TW", "1210.TW", "1314.TW", "1319.TW", "1476.TW", "1477.TW", 
    "1536.TW", "1609.TW", "1707.TW", "1717.TW", "1722.TW", "1795.TW", "1802.TW", "2006.TW", 
    "2014.TW", "2027.TW", "2049.TW", "2101.TW", "2106.TW", "2201.TW", "2204.TW", "2231.TW", 
    "2612.TW", "2637.TW", "2707.TW", "2723.TW", "2915.TW", "6505.TW", "8436.TW", "9907.TW", 
    "9933.TW", "9938.TW", "9939.TW", "9945.TW"
]

FINMIND_TOKEN = os.environ.get("FINMIND_TOKEN", "")

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
        
        hist['20MA'] = hist['Close'].rolling(20).mean()
        
        # KD 計算 (9,3,3)
        low_9 = hist['Low'].rolling(9).min()
        high_9 = hist['High'].rolling(9).max()
        hist['K'] = (((hist['Close'] - low_9) / (high_9 - low_9)) * 100).ewm(com=2, adjust=False).mean()
        hist['D'] = hist['K'].ewm(com=2, adjust=False).mean()
        
        td = hist.iloc[-1]
        y_td = hist.iloc[-2]
        
        # === 條件判斷 (站上月線 + KD<50且剛金叉) ===
        cond_1 = td['Close'] > td['20MA']
        cond_3 = td['K'] < 50 and td['K'] > td['D'] and y_td['K'] <= y_td['D']
        
        if cond_1 and cond_3:
            yoy = get_revenue_yoy(symbol)
            if yoy is not None and yoy < 0: return None # 營收衰退者淘汰
            
            yoy_str = f"{yoy:.1f}%" if yoy is not None else "無API資料"
            return {"symbol": symbol, "close": td['Close'], "K": td['K'], "D": td['D'], "yoy": yoy_str}
    except: return None
    return None

if __name__ == "__main__":
    print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] 🚀 NOC 游擊隊雷達 (擴充版) 啟動，掃描目標 {len(SCAN_LIST)} 檔...")
    print("=" * 60)
    found_targets = []
    
    for sym in SCAN_LIST:
        print(f"正在掃描 {sym}...", end="\r")
        result = scan_stock(sym)
        if result: found_targets.append(result)
        
    print("\n" + "=" * 60)
    
    TARGET_FILE = "radar_targets.json"
    
    if not found_targets:
        print("🎯 報告總操盤手，目前無符合【KD < 50 金叉 + 站上月線 + 營收成長】之標的。")
        # 即使沒掃到，也要寫入一個「空字典」，將前一次的雷達名單無情洗掉！
        with open(TARGET_FILE, "w", encoding="utf-8") as f:
            json.dump({}, f, ensure_ascii=False, indent=4)
        print(f"🧹 雷達畫面已清空。戰情室的【🎯 雷達鎖定區】將同步淨空。")
    else:
        print("🎯 發現符合廣域掃描條件的潛力股：")
        
        # --- 覆蓋模式：每次都建立全新的字典 ---
        radar_dict = {}
        for t in found_targets:
            print(f"🔹 {t['symbol']:>9} | 現價: {t['close']:>6.1f} | K值: {t['K']:>4.1f} | 營收YoY: {t['yoy']}")
            radar_dict[t['symbol']] = f"雷達選股 (進場價約 {t['close']:.1f})"
                
        # 無情覆蓋存檔
        with open(TARGET_FILE, "w", encoding="utf-8") as f:
            json.dump(radar_dict, f, ensure_ascii=False, indent=4)
        print(f"✅ 雷達畫面已刷新！最新火種已裝填至 {TARGET_FILE}，待戰情室接手追蹤。")
        
    print("=" * 60)
