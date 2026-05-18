# =============================================================================
# NOC 游擊隊雷達 v13.0 - 雙引擎突擊版 (60分K盤中即時 + 動能突破/低檔伏擊)
# 優化項目：導入 60m 戰鬥K線、RSI 動能指標、雙軌制戰術 (突擊/伏擊)
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
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

# 強制靜音 yfinance 錯誤日誌 (保留 v12.5 的抗干擾防護)
logging.getLogger('yfinance').setLevel(logging.CRITICAL)

FINMIND_TOKEN = os.environ.get("FINMIND_TOKEN", "")
MAX_WORKERS   = int(os.environ.get("MAX_WORKERS", "5"))
TARGET_FILE   = "radar_targets.json"

# === 1. 設定掃描池 (黃金淬鍊 170 檔 - 保留原設定) ===
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
# === 2. 營收動能解析模組 (保持不變) ===
# =============================================================================
def get_revenue_yoy(symbol):
    if not FINMIND_TOKEN: return None
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
# === 3. 雙軌制戰鬥掃描引擎 (全面升級) ===
# =============================================================================
def scan_stock(symbol):
    try:
        time.sleep(random.uniform(0.3, 1.0))
        
        # 🛡️ 重大變更：抓取「過去 1 個月」的「60 分鐘 K 線」
        hist = yf.Ticker(symbol).history(period="1mo", interval="60m").dropna(subset=['Close'])
        if len(hist) < 30: 
            return None
        
        # 計算均線
        hist['20MA'] = hist['Close'].rolling(20).mean()
        hist['5MA'] = hist['Close'].rolling(5).mean()
        
        # 計算 KD (9,3,3)
        low_9 = hist['Low'].rolling(9).min()
        high_9 = hist['High'].rolling(9).max()
        kd_range = (high_9 - low_9).replace(0, float('nan'))
        hist['K'] = (((hist['Close'] - low_9) / kd_range) * 100).ewm(com=2, adjust=False).mean()
        hist['D'] = hist['K'].ewm(com=2, adjust=False).mean()

        # 計算 RSI (14) 作為動能指標
        delta = hist['Close'].diff()
        gain = (delta.where(delta > 0, 0)).rolling(window=14).mean()
        loss = (-delta.where(delta < 0, 0)).rolling(window=14).mean()
        rs = gain / loss.replace(0, float('nan'))
        hist['RSI'] = 100 - (100 / (1 + rs))
        
        td = hist.iloc[-1]
        y_td = hist.iloc[-2]
        
        # ==========================================
        # 🟢 戰術 A：低檔伏擊 (原本的 KD 金叉邏輯)
        # ==========================================
        cond_a = (td['Close'] > td['20MA']) and (td['K'] < 50) and (td['K'] > td['D']) and (y_td['K'] <= y_td['D'])
        
        # ==========================================
        # 🔥 戰術 B：動能突擊 (不看 KD，專抓起漲飆股)
        # ==========================================
        high_20 = hist['High'].rolling(20).max().iloc[-2] # 過去20根(相當於近4天)的最高點
        # 條件：突破近期高點 + 短均線上彎(5MA>20MA) + RSI>60 強勢
        cond_b = (td['Close'] > high_20) and (td['5MA'] > td['20MA']) and (td['RSI'] > 60)
        
        if cond_a or cond_b:
            yoy = get_revenue_yoy(symbol)
            if yoy is not None and yoy < 0: 
                return None # 營收衰退者依舊無情淘汰
            
            yoy_str = f"{yoy:.1f}%" if yoy is not None else "無API資料"
            tactics = "🔥 突擊 (動能突破)" if cond_b else "🟢 伏擊 (低檔金叉)"
            
            return {
                "symbol": symbol, 
                "close": td['Close'], 
                "tactics": tactics,
                "RSI": td['RSI'],
                "yoy": yoy_str
            }
            
    except Exception as e:
        return None
    return None

# =============================================================================
# === 4. 主程式 ===
# =============================================================================
if __name__ == "__main__":
    start_time = time.time()
    logger.info(f"🚀 NOC 游擊隊雷達 (v13.0 雙引擎突擊版 - 60分K) 啟動，掃描目標 {len(SCAN_LIST)} 檔...")
    logger.info("=" * 60)
    
    found_targets = []
    
    executor = ThreadPoolExecutor(max_workers=MAX_WORKERS)
    future_to_symbol = {executor.submit(scan_stock, sym): sym for sym in SCAN_LIST}
    
    try:
        for future in as_completed(future_to_symbol, timeout=300):
            sym = future_to_symbol[future]
            try:
                result = future.result()
                if result:
                    found_targets.append(result)
                    logger.info(f"🎯 鎖定: {sym} | 戰術: {result['tactics']} | 現價: {result['close']:.1f} | RSI: {result['RSI']:.1f}")
            except Exception:
                pass
    except TimeoutError:
        logger.error("🚨 偵測到網路卡死！達到 5 分鐘斷頭死線，中止剩餘掃描！")
    finally:
        executor.shutdown(wait=False, cancel_futures=True)
                
    elapsed = time.time() - start_time
    logger.info("=" * 60)
    logger.info(f"⏱️ 戰鬥掃描結束！總耗時: {elapsed:.1f} 秒")
    
    if not found_targets:
        logger.info("📡 報告總操盤手，盤中無符合【突擊/伏擊】之標的。")
        with open(TARGET_FILE, "w", encoding="utf-8") as f:
            json.dump({}, f, ensure_ascii=False, indent=4)
    else:
        logger.info(f"🎯 發現 {len(found_targets)} 檔盤中發動的潛力股！")
        radar_dict = {}
        for t in found_targets:
            radar_dict[t['symbol']] = f"[{t['tactics']}] 現價 {t['close']:.1f} | 營收 {t['yoy']}"
                
        with open(TARGET_FILE, "w", encoding="utf-8") as f:
            json.dump(radar_dict, f, ensure_ascii=False, indent=4)
            
        logger.info(f"✅ 雷達刷新！最新火種已裝填至 {TARGET_FILE}。")
