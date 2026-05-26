# update_db.py
import datetime
import os
import time
from dotenv import load_dotenv

from noc_core import NOCDatabase, NOCDataFetcher

load_dotenv()
FINMIND_TOKEN = os.getenv("FINMIND_TOKEN")

# === 1. 直接寫死掃描池 (與 init_db.py 完全相同) ===
SCAN_LIST = [
    # 總司令專屬：黃金淬鍊 170 檔波段觀察池 (此處以高流動性權值與成長股為範例)
    # [權值前 50 大]
    "2330.TW", "2317.TW", "2454.TW", "2382.TW", "2308.TW", "3231.TW", "3037.TW", "2303.TW",
    "3008.TW", "3034.TW", "3711.TW", "2357.TW", "2395.TW", "2408.TW", "2353.TW", "2379.TW",
    "4938.TW", "2301.TW", "2345.TW", "2324.TW", "3661.TW", "6669.TW", "3714.TW", "2881.TW",
    "2882.TW", "2891.TW", "2886.TW", "2884.TW", "2892.TW", "2885.TW", "2880.TW", "2883.TW",
    "2887.TW", "5871.TW", "2890.TW", "5880.TW", "2002.TW", "1216.TW", "1301.TW", "1303.TW",
    "1326.TW", "2912.TW", "9904.TW", "2603.TW", "2609.TW", "2615.TW", "2207.TW", "1101.TW",
    "1102.TW", "2412.TW",
    # [高動能科技 60 檔]
    "2356.TW", "3163.TWO", "5388.TW", "8299.TWO", "3260.TWO", "2377.TW", "2383.TW", "3017.TW",
    "2352.TW", "3443.TW", "3529.TWO", "3293.TWO", "6488.TWO", "8069.TWO", "6274.TWO", "6239.TW",
    "3044.TW", "2449.TW", "2344.TW", "2409.TW", "3481.TW", "6116.TW", "4958.TW", "6176.TW",
    "3532.TW", "2371.TW", "2404.TW", "3702.TW", "8046.TW", "5483.TWO", "3105.TWO", "5347.TWO",
    "6147.TWO", "6214.TW", "2313.TW", "2368.TW", "3013.TW", "3019.TW", "3042.TW", "3324.TWO",
    "3533.TW", "3583.TW", "3653.TW", "4966.TWO", "5269.TW", "6269.TW", "6415.TW", "6531.TW",
    "8016.TW", "8081.TW", "8150.TW", "3376.TW", "3035.TW", "3227.TWO", "3131.TWO", "2451.TW", 
    "5469.TW", "3413.TW", "3450.TW", "4919.TW",
    # [熱門傳產/重電/生技 60 檔] 
    "1513.TW", "1514.TW", "1519.TW", "1605.TW", "1477.TW", "2049.TW", "2610.TW", "2618.TW", 
    "1536.TW", "1795.TW", "2231.TW", "8464.TW", "9910.TW", "9914.TW", "9921.TW", "9941.TW", 
    "1319.TW", "1707.TW", "1722.TW", "1504.TW", "1590.TW", "1609.TW", "1717.TW", "1802.TW", 
    "2006.TW", "2014.TW", "2027.TW", "2101.TW", "2105.TW", "2106.TW", "2201.TW", "2204.TW", 
    "2801.TW", "2834.TW", "2838.TW", "2845.TW", "2889.TW", "6005.TW", "2812.TW", "3045.TW", 
    "4904.TW", "1229.TW", "1402.TW", "2606.TW", "5522.TW", "1108.TW", "1210.TW", "1314.TW", 
    "1476.TW", "2633.TW", "4104.TW", "4105.TWO", "4142.TW", "4736.TW", "6446.TW", "6472.TW", 
    "6550.TW", "6579.TW", "8422.TW", "6806.TW"
]

if __name__ == "__main__":
    print("🚚 開始執行 NOC 盤後戰情資料庫補給作業 (鋼鐵直寫版)...")
    db = NOCDatabase()
    fetcher = NOCDataFetcher(token=FINMIND_TOKEN)
    
    # 每日補給只抓 3 天
    start_date = (datetime.datetime.now() - datetime.timedelta(days=3)).strftime("%Y-%m-%d")
    
    try:
        print("1️⃣ 正在更新大盤指數與外資期貨空單數據...")
        fetcher.fetch_market_health_data(start_date, db)
        
        target_stocks = list(set([s.split('.')[0] for s in SCAN_LIST if s.split('.')[0].isdigit()]))
        print(f"🎯 鎖定 {len(target_stocks)} 檔目標，準備開始日常補給！")

        for sym in target_stocks:
            fetcher.fetch_and_store_stock_data(sym, start_date, db)
            time.sleep(0.5) 
            
        print("\n✅ 戰情資料庫補給完畢！雷達彈藥充足！")
        
    except Exception as e:
        print(f"\n⚠️ 補給過程發生錯誤: {e}")
