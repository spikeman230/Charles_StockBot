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
    SCAN_LIST   : list = [
        # 總司令專屬：黃金淬鍊波段觀察池 (與 noc_radar 共用或獨立設定)
        "2330.TW", "2317.TW", "2454.TW", "2308.TW", "2382.TW", "3231.TW", "2376.TW",
        "2356.TW", "2301.TW", "2603.TW", "2609.TW", "2615.TW", "1519.TW", "1504.TW",
        "1513.TW", "1514.TW", "3034.TW", "2379.TW", "3008.TW", "3017.TW", "3443.TW"
    ]

cfg = LightningConfig()

# =============================================================================
# === 2. 籌碼異常與爆量觀測引擎 ===
# =============================================================================
def scan_stock_for_anomaly(symbol: str) -> dict:
    """
    籌碼動能異常觀測器：
    不進行任何交易決策，僅偵測單日爆量 (大於 5 日均量 2 倍) 且價格劇烈波動的標的，
    做為主力籌碼介入的「足跡」觀察。
    """
    try:
        stock = yf.Ticker(symbol)
        hist = stock.history(period="1mo").dropna(subset=["Close", "Volume"])
        
        if len(hist) < 5:
            return None
            
        # 精算基礎均量與現價
        hist['5VMA'] = hist['Volume'].rolling(5).mean()
        current_vol = hist['Volume'].iloc[-1]
        vma5 = hist['5VMA'].iloc[-2] # 取前一日的 5日均量作為基準比較
        
        current_close = hist['Close'].iloc[-1]
        prev_close = hist['Close'].iloc[-2]
        price_change_pct = ((current_close - prev_close) / prev_close) * 100
        
        # 防呆機制：避免除以零
        if pd.isna(vma5) or vma5 == 0:
            return None
            
        vol_ratio = current_vol / vma5
        
        # 🛡️ 觀測條件：爆量 2 倍以上，且漲幅大於 3% (顯示有主力點火跡象)
        if vol_ratio >= 2.0 and price_change_pct >= 3.0:
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
            
    except Exception as e:
        return None
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
