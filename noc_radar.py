# =============================================================================
# NOC 游擊隊雷達 v12.0 - 極速並行掃描版
# 優化項目：多執行緒掃描、防爬蟲微延遲、.env 環境變數、專業日誌系統
# 戰略邏輯：站上 20MA + KD(9,3,3) 低檔金叉 + 營收 YoY 正成長
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
from concurrent.futures import ThreadPoolExecutor, as_completed
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

# === 機密與參數設定 ===
FINMIND_TOKEN = os.environ.get("FINMIND_TOKEN", "")
MAX_WORKERS   = int(os.environ.get("MAX_WORKERS", "8"))
TARGET_FILE   = "radar_targets.json"

# === 1. 設定掃描池 (台股前150大中型權值股 + 產業指標股 - 已校正TW/TWO) ===
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
    "2105.TW", "2606.TW", "2610.TW", "2618.TW", "5522.TW", "8464.TW", "9910.TW", "9914.TW", 
    "9921.TW", "9941.TW", "1108.TW", "1210.TW", "1314.TW", "1319.TW", "1476.TW", "1477.TW", 
    "1536.TW", "1609.TW", "1707.TW", "1717.TW", "1722.TW", "1795.TW", "1802.TW", "2006.TW", 
    "2014.TW", "2027.TW", "2049.TW", "2101.TW", "2106.TW", "2201.TW", "2204.TW", "2231.TW", 
    "2612.TW", "2637.TW", "2707.TW", "2723.TW", "2915.TW", "6505.TW", "8436.TWO", "9907.TW", 
    "9933.TW", "9938.TW", "9939.TW", "9945.TW"
]

# =============================================================================
# === 2. 營收動能解析模組 ===
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
        response.raise_for_status()
        data = response.json()
        
        if data.get("msg") == "success" and len(data.get("data", [])) > 0:
            df = pd.DataFrame(data["data"])
            latest = df.iloc[-1]
            last_year = df[(df['revenue_year'] == latest['revenue_year'] - 1) & (df['revenue_month'] == latest['revenue_month'])]
            
            if not last_year.empty and last_year.iloc[-1]['revenue'] > 0:
                return ((latest['revenue'] - last_year.iloc[-1]['revenue']) / last_year.iloc[-1]['revenue']) * 100
    except Exception as e:
        logger.debug(f"[{symbol}] FinMind API 錯誤或無資料: {e}")
    return None

# =============================================================================
# === 3. 單一標的掃描引擎 ===
# =============================================================================
def scan_stock(symbol):
    try:
        # 🛡️ 加入防爬蟲隨機微延遲 (0.2 ~ 1.2 秒)
        time.sleep(random.uniform(0.2, 1.2))
        
        hist = yf.Ticker(symbol).history(period="3mo").dropna(subset=['Close'])
        if len(hist) < 30: 
            return None
        
        hist['20MA'] = hist['Close'].rolling(20).mean()
        
        # KD 計算 (9,3,3)
        low_9 = hist['Low'].rolling(9).min()
        high_9 = hist['High'].rolling(9).max()
        kd_range = (high_9 - low_9).replace(0, float('nan'))
        hist['K'] = (((hist['Close'] - low_9) / kd_range) * 100).ewm(com=2, adjust=False).mean()
        hist['D'] = hist['K'].ewm(com=2, adjust=False).mean()
        
        td = hist.iloc[-1]
        y_td = hist.iloc[-2]
        
        # === 核心戰法：站上月線 + KD<50且剛金叉 ===
        cond_1 = td['Close'] > td['20MA']
        cond_3 = (td['K'] < 50) and (td['K'] > td['D']) and (y_td['K'] <= y_td['D'])
        
        if cond_1 and cond_3:
            yoy = get_revenue_yoy(symbol)
            # 營收衰退者無情淘汰
            if yoy is not None and yoy < 0: 
                return None 
            
            yoy_str = f"{yoy:.1f}%" if yoy is not None else "無API資料"
            return {
                "symbol": symbol, 
                "close": td['Close'], 
                "K": td['K'], 
                "D": td['D'], 
                "yoy": yoy_str
            }
            
    except Exception as e:
        logger.debug(f"[{symbol}] 分析過程發生錯誤: {e}")
        return None
    return None

# =============================================================================
# === 4. 主程式 (多執行緒並行調度) ===
# =============================================================================
if __name__ == "__main__":
    start_time = time.time()
    logger.info(f"🚀 NOC 游擊隊雷達 (極速並行版) 啟動，掃描目標 {len(SCAN_LIST)} 檔...")
    logger.info("=" * 60)
    
    found_targets = []
    
    # ⚡ 使用 ThreadPoolExecutor 進行多執行緒掃描
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        # 將任務交給工作池
        future_to_symbol = {executor.submit(scan_stock, sym): sym for sym in SCAN_LIST}
        
        # 收集完成的結果
        for future in as_completed(future_to_symbol):
            sym = future_to_symbol[future]
            try:
                result = future.result()
                if result:
                    found_targets.append(result)
                    logger.info(f"⚡ 捕捉到潛力目標: {sym} (現價: {result['close']:.1f}, KD: {result['K']:.1f}/{result['D']:.1f})")
            except Exception as e:
                logger.error(f"❌ 處理 {sym} 時發生不可預期的錯誤: {e}")
                
    elapsed = time.time() - start_time
    logger.info("=" * 60)
    logger.info(f"⏱️ 掃描完成！總耗時: {elapsed:.1f} 秒")
    
    # === 寫入戰報 ===
    if not found_targets:
        logger.info("🎯 報告總操盤手，目前無符合【KD < 50 金叉 + 站上月線 + 營收成長】之標的。")
        # 即使沒掃到，也要寫入一個「空字典」，將前一次的雷達名單洗掉
        with open(TARGET_FILE, "w", encoding="utf-8") as f:
            json.dump({}, f, ensure_ascii=False, indent=4)
        logger.info(f"🧹 雷達畫面已清空。戰情室的【🎯 雷達鎖定區】將同步淨空。")
    else:
        logger.info(f"🎯 發現 {len(found_targets)} 檔符合條件的潛力股：")
        
        # 建立全新的字典並覆蓋存檔
        radar_dict = {}
        for t in found_targets:
            logger.info(f"  🔹 {t['symbol']:>9} | 現價: {t['close']:>6.1f} | K值: {t['K']:>4.1f} | 營收YoY: {t['yoy']}")
            radar_dict[t['symbol']] = f"雷達選股 (進場價約 {t['close']:.1f})"
                
        with open(TARGET_FILE, "w", encoding="utf-8") as f:
            json.dump(radar_dict, f, ensure_ascii=False, indent=4)
            
        logger.info(f"✅ 雷達畫面已刷新！最新火種已裝填至 {TARGET_FILE}，待戰情室接手追蹤。")
        
    logger.info("=" * 60)
