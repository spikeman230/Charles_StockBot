import yfinance as yf
import requests
import datetime
import pandas as pd
import os
import json

# === 1. 設定掃描池 (已校正之 160 檔台股中大型/熱門指標股) ===
SCAN_LIST = [
    # 半導體、電子與 AI 伺服器
    "2330.TW", "2317.TW", "2454.TW", "2382.TW", "2308.TW", "3231.TW", "3037.TW", "2303.TW", 
    "3008.TW", "3034.TW", "3711.TW", "2357.TW", "2395.TW", "2408.TW", "2353.TW", "2356.TW", 
    "2379.TW", "4938.TW", "2301.TW", "2345.TW", "2324.TW", "3661.TW", "6669.TW", "3714.TW", 
    "3163.TWO", "5388.TW", "8299.TWO", "3260.TWO", "2377.TW", "2383.TW", "3017.TW", "2352.TW", 
    "3443.TW", "3529.TWO", "3293.TWO", "6488.TWO", "8069.TWO", "6274.TWO", "6239.TW", "3044.TW", 
    "2449.TW", "2344.TW", "2409.TW", "3481.TW", "6116.TW", "4958.TW", "6176.TW", "3532.TW", 
    "2371.TW", "2404.TW", "3702.TW", "8046.TW", "5483.TWO", "3105.TWO", "5347.TWO", "6147.TWO", 
    "6214.TW", "2313.TW", "2368.TW", "3013.TW", "3019.TW", "3042.TW", "3324.TWO", "3533.TW", 
    "3583.TW", "3653.TW", "4966.TWO", "5269.TW", "6269.TW", "6415.TW", "6531.TW", "8016.TW", 
    "8081.TW", "8150.TW",
    # 金融權值股
    "2881.TW", "2882.TW", "2891.TW", "2886.TW", "2884.TW", "2892.TW", "2885.TW", "2880.TW", 
    "2883.TW", "2887.TW", "5871.TW", "2890.TW", "5880.TW", "2801.TW", "2834.TW", "2838.TW", 
    "2845.TW", "2889.TW", "6005.TW", "2812.TW",
    # 航運、傳產、生技與電信
    "2412.TW", "3045.TW", "4904.TW", "2002.TW", "1216.TW", "1301.TW", "1303.TW", "1326.TW", 
    "2912.TW", "9904.TW", "2603.TW", "2609.TW", "2615.TW", "2207.TW", "1101.TW", "1102.TW", 
    "1229.TW", "1402.TW", "1504.TW", "1513.TW", "1514.TW", "1519.TW", "1590.TW", "1605.TW", 
    "2105.TW", "2606.TW", "2610.TW", "2618.TW", "5522.TW", "8464.TWO", "9910.TW", "9914.TW", 
    "9921.TW", "9941.TW", "1108.TW", "1210.TW", "1314.TW", "1319.TW", "1476.TW", "1477.TW", 
    "1536.TW", "1609.TW", "1707.TW", "1717.TW", "1722.TW", "1795.TW", "1802.TW", "2006.TW", 
    "2014.TW", "2027.TW", "2049.TW", "2101.TW", "2106.TW", "2201.TW", "2204.TW", "2231.TW", 
    "2612.TW", "2637.TW", "2707.TW", "2723.TW", "2915.TW", "6505.TW", "9907.TW", "9933.TW", 
    "9938.TW", "9939.TW", "9945.TW"
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
        hist = yf.Ticker(symbol).history(period="2mo").dropna(subset=['Close', 'Volume'])
        if len(hist) < 25: return None
        
        hist['5MA'] = hist['Close'].rolling(5).mean()
        hist['20MA'] = hist['Close'].rolling(20).mean()
        hist['5VMA'] = hist['Volume'].rolling(5).mean()
        
        td = hist.iloc[-1]
        y_td = hist.iloc[-2]
        
        # === 閃電突擊條件 ===
        price_change = ((td['Close'] - y_td['Close']) / y_td['Close']) * 100
        
        cond_1 = td['Close'] > td['5MA'] # 站上5日線
        cond_2 = td['Close'] > td['20MA'] # 確保趨勢偏多
        cond_3 = td['Volume'] > (td['5VMA'] * 2.0) # 爆量: 今日量大於5日均量的2倍
        cond_4 = td['Close'] > td['Open'] and price_change > 3.0 # 實體紅K 且 漲幅大於 3%
        
        if cond_1 and cond_2 and cond_3 and cond_4:
            yoy = get_revenue_yoy(symbol)
            yoy_str = f"{yoy:.1f}%" if yoy is not None else "無API資料"
            return {
                "symbol": symbol, 
                "close": td['Close'], 
                "change": price_change,
                "vol_ratio": td['Volume'] / td['5VMA'], # 爆量倍數
                "yoy": yoy_str
            }
    except: return None
    return None

if __name__ == "__main__":
    print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] ⚡ NOC 閃電突擊雷達 (極短線模式) 啟動...")
    print("=" * 65)
    found_targets = []
    
    for sym in SCAN_LIST:
        print(f"正在掃描 {sym}...", end="\r")
        result = scan_stock(sym)
        if result: found_targets.append(result)
        
    print("\n" + "=" * 65)
    
    # 注意：這裡使用另一個 JSON 檔名，避免與原本的游擊隊雷達打架
    TARGET_FILE = "lightning_targets.json"
    
    if not found_targets:
        print("🎯 報告總操盤手，今日無符合【爆量2倍 + 漲幅>3% + 站上5MA】之突擊標的。")
        with open(TARGET_FILE, "w", encoding="utf-8") as f:
            json.dump({}, f, ensure_ascii=False, indent=4)
        print(f"🧹 閃電突擊畫面已淨空。")
    else:
        print("🎯 發現符合閃電突擊條件的高動能飆股：")
        
        lightning_dict = {}
        for t in found_targets:
            print(f"⚡ {t['symbol']:>9} | 現價:{t['close']:>6.1f} | 漲幅:+{t['change']:>4.1f}% | 爆量:{t['vol_ratio']:>3.1f}倍 | YoY:{t['yoy']}")
            # 寫入戰情室時，加上進場價與 5MA 防守提示
            lightning_dict[t['symbol']] = f"閃電突擊 (參考價 {t['close']:.1f}，跌破5MA停損)"
                
        with open(TARGET_FILE, "w", encoding="utf-8") as f:
            json.dump(lightning_dict, f, ensure_ascii=False, indent=4)
        print(f"✅ 閃電目標已鎖定！裝填至 {TARGET_FILE}。")
        
    print("=" * 65)
