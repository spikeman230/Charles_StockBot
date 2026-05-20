# =============================================================================
# NOC 游擊隊雷達 (noc_radar.py) v14.5 - 長線波段戰備收斂版
# 核心戰術：
# 1. 徹底廢除盤中跳動觸發，轉型為「每日收盤後執行」的戰備火種篩選器。
# 2. 強制掛載 NOCStrategy 雙大腦 (Trend_Score 季線趨勢 + YoY 基本面)。
# 3. 嚴禁任何自動化買入指令，僅產出高質量波段火種供總司令決策。
# 4. 程式碼嚴格遵守不精簡、完整展開之軍規鐵律，保留所有防爆機制。
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
from noc_core import NOCStrategy, NOCDatabase

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

# 強制靜音 yfinance 錯誤日誌，防止因個別 API 失敗洗版戰情室
logging.getLogger('yfinance').setLevel(logging.CRITICAL)

# =============================================================================
# === 1. 雷達全域組態與掃描池設定 ===
# =============================================================================
class RadarConfig:
    MAX_WORKERS : int = int(os.environ.get("MAX_WORKERS", "5"))
    TARGET_FILE : str = "radar_targets.json"
    SCAN_LIST   : list = [
        # 總司令專屬：黃金淬鍊 170 檔波段觀察池 (此處以高流動性權值與成長股為範例)
        "2330.TW", "2317.TW", "2454.TW", "2308.TW", "2382.TW", "3231.TW", "2376.TW",
        "2356.TW", "2301.TW", "2603.TW", "2609.TW", "2615.TW", "1519.TW", "1504.TW",
        "1513.TW", "1514.TW", "3034.TW", "2379.TW", "3008.TW", "3017.TW", "3443.TW",
        # ETF 戰略防禦與市值觀測群
        "0050.TW", "006208.TW", "0056.TW", "00878.TW", "00713.TW", "00919.TW"
        # 註：實際執行時可自由擴充此陣列至 170 檔，系統具備自動多執行緒調度能力。
    ]

cfg = RadarConfig()

# =============================================================================
# === 2. 核心掃描引擎 (掛載波段雙濾網) ===
# =============================================================================
def scan_stock_for_wave(symbol: str, strategy: NOCStrategy) -> dict:
    """
    波段火種精煉器：
    專注於日線級別 (Daily K) 的 60MA 季線斜率與基本面體質檢查。
    """
    try:
        # 1. 抓取長線波段所需之半年日 K 線資料 (不看分時短線)
        stock = yf.Ticker(symbol)
        hist = stock.history(period="6mo").dropna(subset=["Close"])
        
        # 標的防呆：若上市櫃時間不足 60 天，無法形成有效季線，直接剔除
        if len(hist) < 60:
            return None
            
        # 2. 核心均線精算
        hist['20MA'] = hist['Close'].rolling(20).mean()
        hist['60MA'] = hist['Close'].rolling(60).mean()
        
        current_close = hist['Close'].iloc[-1]
        
        # 3. 🛡️ 第一道硬核濾網：趨勢得分 (Trend Score)
        # 判定季線斜率是否向上，且現價不可距離季線過遠 (乖離 < 15%)
        trend_score = strategy.get_trend_score(hist)
        if trend_score < 0:
            # 趨勢不符 (下彎或過熱)，淘汰
            return None
            
        # 4. 🛡️ 第二道硬核濾網：基本面護城河 (Fundamental Health)
        raw_id = symbol.replace(".TW", "").replace(".TWO", "")
        fund_health = strategy.get_fundamental_health(raw_id)
        if "衰退" in fund_health or "警報" in fund_health:
            # 營收衰退，無情淘汰
            return None

        # 5. 精算動能指標，做為觀察輔助 (RSI & 乖離率)
        delta = hist["Close"].diff()
        rs = delta.clip(lower=0).ewm(com=13, adjust=False).mean() / (-delta.clip(upper=0)).ewm(com=13, adjust=False).mean().replace(0, 0.001)
        rsi = (100 - (100 / (1 + rs))).iloc[-1]
        
        ma20_val = hist['20MA'].iloc[-1]
        bias_20 = ((current_close - ma20_val) / ma20_val) * 100 if ma20_val else 0

        # 6. 產出波段火種報告
        return {
            "symbol": symbol,
            "name": raw_id, # 暫存代碼，交由 stock_bot Trello 模組關聯中文名
            "close": round(current_close, 2),
            "RSI": round(rsi, 2),
            "Bias20": round(bias_20, 2),
            "tactics": "🔥 長線波段多頭 (符合季線上揚與營收健康)",
            "trello_tip": "系統雷達自動篩選，等待總司令確認建倉。"
        }
        
    except Exception as e:
        # 雷達單兵掃描失敗不影響整體陣列，靜默回報
        return None

# =============================================================================
# === 3. 主控作戰執行緒 (Main Execution) ===
# =============================================================================
if __name__ == "__main__":
    logger.info("⚡ NOC 游擊隊雷達 v14.5 (收盤波段戰備版) 啟動...")
    start_time = time.time()
    
    # 初始化核心戰略模組
    strategy = NOCStrategy()
    
    # 檢查大盤狀態，若為極端空頭 (DEFCON 1)，直接放棄尋找火種
    macro = strategy.get_macro_status()
    if macro["status"] == "🔴 紅燈":
        logger.warning("🚨 大盤跌破季線，防空警報大響。系統拒絕掃描任何新火種，強制保護現金！")
        # 清空目標檔，確保明日 stock_bot 不會有異常買進提示
        with open(cfg.TARGET_FILE, "w", encoding="utf-8") as f:
            json.dump({}, f, ensure_ascii=False, indent=4)
        exit(0)

    found_targets = []
    logger.info(f"📡 大盤風向正常 ({macro['status']})，開始對 {len(cfg.SCAN_LIST)} 檔標的進行『波段雙核濾網』深度掃描...")

    # 啟動多執行緒高並行掃描 (具備防爆盾機制)
    executor = ThreadPoolExecutor(max_workers=cfg.MAX_WORKERS)
    future_to_symbol = {executor.submit(scan_stock_for_wave, sym, strategy): sym for sym in cfg.SCAN_LIST}
    
    try:
        for future in as_completed(future_to_symbol, timeout=300):
            sym = future_to_symbol[future]
            try:
                result = future.result()
                if result:
                    found_targets.append(result)
                    logger.info(f"🎯 成功鎖定長線火種: {sym} | 現價: {result['close']} | RSI: {result['RSI']}")
            except Exception:
                pass
    except TimeoutError:
        logger.error("🚨 偵測到網路嚴重延遲！達到 5 分鐘斷頭死線，強制中止剩餘雷達掃描！")
    finally:
        # 斬斷所有卡死的殭屍執行緒
        executor.shutdown(wait=False, cancel_futures=True)
                
    elapsed = time.time() - start_time
    logger.info("=" * 65)
    logger.info(f"⏱️ 波段戰備巡邏結束！總耗時: {elapsed:.1f} 秒")
    
    # === 戰果結算與檔案寫入 ===
    if not found_targets:
        logger.info("📡 報告總司令，今日收盤後無任何標的通過『季線上揚 + 營收健康』之雙重嚴苛濾網。")
        with open(cfg.TARGET_FILE, "w", encoding="utf-8") as f:
            json.dump({}, f, ensure_ascii=False, indent=4)
    else:
        logger.info(f"🎯 淬鍊完成！共發現 {len(found_targets)} 檔符合龍蝦養殖標準之高價值火種：")
        
        # 將清單格式化為 stock_bot 兼容之字典結構
        radar_dict = {}
        for tgt in found_targets:
            radar_dict[tgt["symbol"]] = {
                "name": tgt["name"],
                "tactics": tgt["tactics"],
                "trello_tip": tgt["trello_tip"]
            }
            
        try:
            with open(cfg.TARGET_FILE, "w", encoding="utf-8") as f:
                json.dump(radar_dict, f, ensure_ascii=False, indent=4)
            logger.info(f"✅ 長線火種清單已成功寫入 {cfg.TARGET_FILE}，已同步至戰情室等候明日判讀。")
        except Exception as e:
            logger.error(f"❌ 寫入目標檔案時發生嚴重錯誤: {e}")
