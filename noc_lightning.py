# =============================================================================
# NOC 閃電突擊雷達 v13.0 - 極速並行掃描 (60分K即時裝甲版)
# 優化項目：導入 60m 戰鬥K線、API防爆盾、強制斷頭台、yfinance 靜音
# 戰略邏輯：站上 60分K的5MA + 單小時漲幅 > 3% + 單小時爆量 2 倍以上
# =============================================================================

import yfinance as yf
import requests
import datetime
import pandas as pd
import os
import json
import time
import random
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed, TimeoutError
from dotenv import load_dotenv

# =============================================================================
# === 0. 初始化：載入環境變數 & 日誌系統 ===
# =============================================================================
load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# 🛡️ 啟動靜音防護罩：強制拔除 yfinance 內建的錯誤廣播喇叭
logging.getLogger('yfinance').setLevel(logging.CRITICAL)

# === 機密與參數設定 ===
FINMIND_TOKEN = os.environ.get("FINMIND_TOKEN", "")
MAX_WORKERS   = int(os.environ.get("MAX_WORKERS", "5")) # 調降並發數保護免費 API
TARGET_FILE   = "lightning_targets.json"

# === 1. 設定掃描池 (已校正之 170 檔台股中大型/熱門指標股) ===
SCAN_LIST = [
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

# =============================================================================
# === 2. 營收動能解析模組 (搭載 API 防爆盾) ===
# =============================================================================
def get_revenue_yoy(symbol):
    if not FINMIND_TOKEN: 
        return None
        
    fm_symbol = symbol.replace(".TW", "").replace(".TWO", "")
    try:
        url = "https://api.finmindtrade.com/api/v4/data"
        params = {
            "dataset": "TaiwanStockMonthRevenue", 
            "data_id": fm_symbol, 
            "start_date": (datetime.datetime.now() - datetime.timedelta(days=400)).strftime("%Y-%m-%d"), 
            "token": FINMIND_TOKEN
        }
        response = requests.get(url, params=params, timeout=10)
        if response.status_code != 200: return None
        
        # 🛡️ 防禦裝甲：防止 JSON 解析錯誤
        try:
            data = response.json()
        except ValueError:
            return None
        
        if data.get("msg") == "success" and len(data.get("data", [])) > 0:
            df = pd.DataFrame(data["data"])
            latest = df.iloc[-1]
            last_year = df[(df['revenue_year'] == latest['revenue_year'] - 1) & (df['revenue_month'] == latest['revenue_month'])]
            
            if not last_year.empty and last_year.iloc[-1]['revenue'] > 0:
                return ((latest['revenue'] - last_year.iloc[-1]['revenue']) / last_year.iloc[-1]['revenue']) * 100
    except Exception:
        pass
    return None

# =============================================================================
# === 3. 單一標的掃描引擎 (60分鐘K閃電邏輯) ===
# =============================================================================
def scan_stock(symbol):
    try:
        # 🛡️ 輕量級防爬蟲微延遲
        time.sleep(random.uniform(0.3, 1.0))
        
        # 🚀 升級為 60 分鐘 K 線，抓取過去 1 個月的資料
        hist = yf.Ticker(symbol).history(period="1mo", interval="60m").dropna(subset=['Close', 'Volume'])
        if len(hist) < 25: 
            return None
        
        # 這裡的 5MA 變成「5小時均線」，5VMA 變成「5小時均量」
        hist['5MA'] = hist['Close'].rolling(5).mean()
        hist['20MA'] = hist['Close'].rolling(20).mean()
        hist['5VMA'] = hist['Volume'].rolling(5).mean()
        
        td = hist.iloc[-1]
        y_td = hist.iloc[-2]
        
        # === 閃電突擊核心戰法 (盤中短兵相接) ===
        price_change = ((td['Close'] - y_td['Close']) / y_td['Close']) * 100
        
        cond_1 = td['Close'] > td['5MA']           # 站上 5小時線 (短線極強)
        cond_2 = td['Close'] > td['20MA']          # 站上 20小時線 (趨勢保護)
        cond_3 = td['Volume'] > (td['5VMA'] * 2.0) # 爆量：單小時成交量大於前5小時均量的2倍
        cond_4 = td['Close'] > td['Open'] and price_change > 3.0 # 實體紅K 且 該小時急拉逾 3%
        
        if cond_1 and cond_2 and cond_3 and cond_4:
            yoy = get_revenue_yoy(symbol)
            if yoy is not None and yoy < 0: 
                return None # 營收衰退的假突破不追
                
            yoy_str = f"{yoy:.1f}%" if yoy is not None else "無API資料"
            return {
                "symbol": symbol, 
                "close": td['Close'], 
                "change": price_change,
                "vol_ratio": td['Volume'] / td['5VMA'], # 爆量倍數
                "yoy": yoy_str
            }
            
    except Exception:
        # 🛡️ 靜音跳過有問題的股票
        return None
    return None

# =============================================================================
# === 4. 主程式 (搭載強制斷頭台) ===
# =============================================================================
if __name__ == "__main__":
    start_time = time.time()
    logger.info(f"⚡ NOC 閃電突擊雷達 (v13.0 60m裝甲版) 啟動，掃描目標 {len(SCAN_LIST)} 檔...")
    logger.info("=" * 65)
    
    found_targets = []
    
    # ⚡ 使用 ThreadPoolExecutor 進行多執行緒掃描
    executor = ThreadPoolExecutor(max_workers=MAX_WORKERS)
    future_to_symbol = {executor.submit(scan_stock, sym): sym for sym in SCAN_LIST}
    
    try:
        # 🛡️ 裝甲升級：設定 300 秒絕對死線
        for future in as_completed(future_to_symbol, timeout=300):
            sym = future_to_symbol[future]
            try:
                result = future.result()
                if result:
                    found_targets.append(result)
                    logger.info(f"⚡ 捕捉到飆股目標: {sym} (現價:{result['close']:.1f}, 漲幅:+{result['change']:.1f}%, 爆量:{result['vol_ratio']:.1f}倍)")
            except Exception:
                pass
                
    except TimeoutError:
        logger.error("🚨 [致命警報] 網路嚴重卡死！已達到 5 分鐘強制斷頭死線，立即中止！")
    finally:
        # 🛡️ 斬斷殭屍執行緒
        executor.shutdown(wait=False, cancel_futures=True)
                
    elapsed = time.time() - start_time
    logger.info("=" * 65)
    logger.info(f"⏱️ 掃描完成！總耗時: {elapsed:.1f} 秒")
    
    # === 寫入戰報 ===
    if not found_targets:
        logger.info("🎯 報告總操盤手，今日盤中無符合【單小時爆量2倍 + 急拉3%】之突擊標的。")
        with open(TARGET_FILE, "w", encoding="utf-8") as f:
            json.dump({}, f, ensure_ascii=False, indent=4)
        logger.info(f"🧹 閃電突擊畫面已淨空。")
    else:
        logger.info(f"🎯 發現 {len(found_targets)} 檔符合閃電突擊條件的高動能飆股：")
        
        lightning_dict = {}
        for t in found_targets:
            logger.info(f"  ⚡ {t['symbol']:>9} | 現價:{t['close']:>6.1f} | 漲幅:+{t['change']:>4.1f}% | 爆量:{t['vol_ratio']:>3.1f}倍 | YoY:{t['yoy']}")
            lightning_dict[t['symbol']] = f"閃電突擊 (參考價 {t['close']:.1f}，跌破60分5MA停損)"
                
        with open(TARGET_FILE, "w", encoding="utf-8") as f:
            json.dump(lightning_dict, f, ensure_ascii=False, indent=4)
            
        logger.info(f"✅ 閃電目標已鎖定！裝填至 {TARGET_FILE}，待戰情室接手追蹤。")
        
    logger.info("=" * 65)
