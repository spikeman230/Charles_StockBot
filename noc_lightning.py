# =============================================================================
# NOC 閃電突擊雷達 (noc_lightning.py) v14.5 - 波段戰備望遠鏡版
# 核心戰術：
# 1. 物理閹割盤中自動化交易權限，純粹作為「籌碼異常波動」的觀測望遠鏡。
# 2. 強制寫入「嚴禁追高」之戰略警語，所有資料僅供收盤後的主戰情室評估。
# 3. 對接 NOCStrategy，大盤紅燈時自動強制休眠，防禦至上。
# 4. 程式碼嚴格遵守不精簡、完整展開之軍規鐵律。
# =============================================================================

import yfinance as yf
import datetime
import pandas as pd
import os
import json
import time
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed, TimeoutError
from dotenv import load_dotenv

# 🌟 深度引入 NOC 核心防禦模組
from noc_core import NOCStrategy

# =============================================================================
# === 0. 初始化：載入環境變數 & 靜音防護罩 ===
# =============================================================================
load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

# 強制靜音 yfinance 錯誤日誌，維持戰情主控台純淨
logging.getLogger('yfinance').setLevel(logging.CRITICAL)

# =============================================================================
# === 1. 雷達全域組態與掃描池設定 ===
# =============================================================================
class LightningConfig:
    MAX_WORKERS : int = int(os.environ.get("MAX_WORKERS", "5"))
    TARGET_FILE : str = "lightning_targets.json"
    SCAN_LIST : list = [
        # 總司令專屬：200 檔旗艦級波段觀察池
        # [權值前 50 大]
        "0050.TW", "2330.TW", "2317.TW", "2454.TW", "2382.TW", "2308.TW", "3231.TW", "3037.TW",
        "2303.TW", "3008.TW", "3034.TW", "3711.TW", "2357.TW", "2395.TW", "2408.TW", "2353.TW",
        "2379.TW", "4938.TW", "2301.TW", "2345.TW", "2324.TW", "3661.TW", "6669.TW", "3714.TW",
        "2881.TW", "2882.TW", "2891.TW", "2886.TW", "2884.TW", "2892.TW", "2885.TW", "2880.TW",
        "2883.TW", "2887.TW", "5871.TW", "2890.TW", "5880.TW", "2002.TW", "1216.TW", "1301.TW",
        "1303.TW", "1326.TW", "2912.TW", "9904.TW", "2603.TW", "2609.TW", "2615.TW", "2207.TW",
        "1101.TW", "1102.TW",
        # [高動能科技 60 檔]
        "2356.TW", "3163.TWO", "5388.TW", "8299.TWO", "3260.TWO", "2377.TW", "2383.TW", "3017.TW",
        "2352.TW", "3443.TW", "3529.TWO", "3293.TWO", "6488.TWO", "8069.TWO", "6274.TWO", "6239.TW",
        "3044.TW", "2449.TW", "2344.TW", "2409.TW", "3481.TW", "6116.TW", "4958.TW", "6176.TW",
        "3532.TW", "2371.TW", "2404.TW", "3702.TW", "8046.TW", "5483.TWO", "3105.TWO", "5347.TWO",
        "6147.TWO", "6214.TW", "2313.TW", "2368.TW", "3013.TW", "3019.TW", "3042.TW", "3324.TWO",
        "3533.TW", "3583.TW", "3653.TW", "4966.TWO", "5269.TW", "6269.TW", "6415.TW", "6531.TW",
        "8016.TW", "8081.TW", "8150.TW", "3376.TW", "3035.TW", "3227.TWO", "3131.TWO", "2451.TW",
        "5469.TW", "3413.TW", "3450.TW", "4919.TW",
        # [區塊 3：重電/綠能/電纜與生技醫療 - 共 60 檔]
        # 重電綠能 (25檔)
        "1513.TW", "1514.TW", "1519.TW", "1605.TW", "1504.TW", "1503.TW", "1515.TW", "6806.TW",
        "3708.TW", "1609.TW", "1608.TW", "1611.TW", "1612.TW", "1618.TW", "9958.TW", "3712.TW",
        "6409.TW", "1582.TW", "1522.TW", "1532.TW", "4536.TW", "8926.TW", "6869.TW", "1537.TW",
        "1520.TW",
        # [區塊 4：傳產塑化/汽車零組件/造船航太 - 共 30 檔]
        # 汽車零組件 (13檔)
        "1536.TW", "2231.TW", "1521.TW", "1525.TW", "2228.TW", "2115.TW", "2201.TW", "2204.TW",
        "3346.TW", "1339.TW", "6279.TW", "1524.TW", "1568.TW",
        # 傳產塑化化學 (12檔)
        "1314.TW", "1717.TW", "1304.TW", "1308.TW", "1309.TW", "1312.TW", "1305.TW", "1710.TW",
        "1704.TW", "4722.TW", "4739.TW", "1718.TW",
        # 造船與軍工航太 (5檔)
        "2208.TW", "2634.TW", "4541.TW", "8222.TW", "2646.TW"
    ]

cfg = LightningConfig()

# =============================================================================
# === 2. 籌碼異常與爆量觀測引擎 ===
# =============================================================================
def scan_stock_for_anomaly(symbol: str) -> dict:
    """
    籌碼動能異常觀測器 + 旱地拔蔥 Boss 級突破偵測
    - 一般異常：爆量 2 倍以上，漲幅 ≥ 3%
    - 旱地拔蔥：昨日收盤低於季線、今日站上季線 + 爆量 3 倍以上 + 長紅 ≥ 4%
    """
    try:
        # 擴充歷史資料至 6 個月，確保足夠計算 60MA
        stock = yf.Ticker(symbol)
        hist = stock.history(period="6mo").dropna(subset=["Close", "Volume"])
        
        if len(hist) < 60:
            return None
            
        # 計算技術指標
        hist['5VMA'] = hist['Volume'].rolling(5).mean()
        hist['60MA'] = hist['Close'].rolling(60).mean()
        
        # 今日與昨日資料
        current_vol = hist['Volume'].iloc[-1]
        prev_vol_ma5 = hist['5VMA'].iloc[-2]          # 昨日的 5 日均量
        current_close = hist['Close'].iloc[-1]
        prev_close = hist['Close'].iloc[-2]
        current_ma60 = hist['60MA'].iloc[-1]
        prev_ma60 = hist['60MA'].iloc[-2]
        
        # 防呆處理
        if pd.isna(prev_vol_ma5) or prev_vol_ma5 == 0:
            return None
            
        vol_ratio = current_vol / prev_vol_ma5
        price_change_pct = ((current_close - prev_close) / prev_close) * 100
        
        # ========== 雙階層判定 ==========
        # 條件 A：旱地拔蔥 (Boss 級)
        just_crossed_60ma = (current_close > current_ma60) and (prev_close <= prev_ma60)
        is_monster = just_crossed_60ma and (vol_ratio >= 3.0) and (price_change_pct >= 4.0)
        
        # 條件 B：一般籌碼異常
        is_anomaly = (vol_ratio >= 2.0) and (price_change_pct >= 3.0)
        
        if is_monster:
            raw_id = symbol.replace(".TW", "").replace(".TWO", "")
            return {
                "symbol": symbol,
                "name": raw_id,
                "close": round(current_close, 2),
                "vol_ratio": round(vol_ratio, 1),
                "change_pct": round(price_change_pct, 1),
                "tactics": "🔥【旱地拔蔥】底部極端爆量，長紅突破季線起漲！",
                "trello_tip": f"極端爆量 {vol_ratio:.1f} 倍！無懼基本面，強烈建議納入短線追蹤！"
            }
        elif is_anomaly:
            raw_id = symbol.replace(".TW", "").replace(".TWO", "")
            return {
                "symbol": symbol,
                "name": raw_id,
                "close": round(current_close, 2),
                "vol_ratio": round(vol_ratio, 1),
                "change_pct": round(price_change_pct, 1),
                "tactics": "⚡ 籌碼動能異常 (嚴禁追高，僅供波段觀察)",
                "trello_tip": f"爆量 {vol_ratio:.1f} 倍。此為市場雜訊觀測，請交由 NOC 核心執行長線濾網判定。"
            }
        else:
            return None
            
    except Exception as e:
        # 發生任何錯誤均靜默回傳 None，不影響整體掃描
        return None

# =============================================================================
# === 3. 主控作戰執行緒 (Main Execution) ===
# =============================================================================
if __name__ == "__main__":
    logger.info("⚡ NOC 閃電突擊雷達 v14.5 (籌碼觀測閹割版) 啟動...")
    start_time = time.time()
    
    # 呼叫 NOCStrategy 判定大盤風向
    strategy = NOCStrategy()
    macro = strategy.get_macro_status()
    
    # 🛡️ 總體防禦：大盤紅燈時，直接切斷電源，不觀測任何異常
    if macro["status"] == "🔴 紅燈":
        logger.warning("🚨 大盤跌破季線，閃電觀測儀強制關閉。嚴禁於空頭市場尋找任何多頭火種！")
        with open(cfg.TARGET_FILE, "w", encoding="utf-8") as f:
            json.dump({}, f, ensure_ascii=False, indent=4)
        exit(0)

    found_targets = []
    logger.info(f"📡 啟動多執行緒籌碼觀測，掃描 {len(cfg.SCAN_LIST)} 檔標的是否有主力異常點火足跡...")

    executor = ThreadPoolExecutor(max_workers=cfg.MAX_WORKERS)
    future_to_symbol = {executor.submit(scan_stock_for_anomaly, sym): sym for sym in cfg.SCAN_LIST}
    
    try:
        for future in as_completed(future_to_symbol, timeout=300):
            sym = future_to_symbol[future]
            try:
                result = future.result()
                if result:
                    found_targets.append(result)
                    logger.info(f"👁️ 觀測到異常籌碼足跡: {sym} | 漲幅: +{result['change_pct']}% | 爆量: {result['vol_ratio']} 倍")
            except Exception:
                pass
    except TimeoutError:
        logger.error("🚨 網路延遲超時，強制中止觀測任務！")
    finally:
        executor.shutdown(wait=False, cancel_futures=True)
                
    elapsed = time.time() - start_time
    logger.info("=" * 65)
    logger.info(f"⏱️ 籌碼異常觀測結束！總耗時: {elapsed:.1f} 秒")
    
    # === 寫入觀測報告 ===
    if not found_targets:
        logger.info("📡 今日盤面無極端爆量異常標的。")
        with open(cfg.TARGET_FILE, "w", encoding="utf-8") as f:
            json.dump({}, f, ensure_ascii=False, indent=4)
    else:
        logger.info(f"🎯 彙整 {len(found_targets)} 檔籌碼異常標的，已上鎖嚴禁自動交易：")
        
        lightning_dict = {}
        for tgt in found_targets:
            lightning_dict[tgt["symbol"]] = {
                "name": tgt["name"],
                "tactics": tgt["tactics"],
                "trello_tip": tgt["trello_tip"]
            }
            
        try:
            with open(cfg.TARGET_FILE, "w", encoding="utf-8") as f:
                json.dump(lightning_dict, f, ensure_ascii=False, indent=4)
            logger.info(f"✅ 異常觀測清單已同步至 {cfg.TARGET_FILE}，將交由主戰情室進行基本面與長線趨勢之終極審判。")
        except Exception as e:
            logger.error(f"❌ 寫入觀測檔案時發生嚴重錯誤: {e}")
