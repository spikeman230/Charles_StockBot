# =============================================================================
# NOC 閃電突擊雷達 (noc_lightning.py) v16.0 - 市場廣度聯動版
# 核心戰術：
# 1. 物理閹割盤中自動化交易權限，純粹作為「籌碼異常波動」的觀測望遠鏡。
# 2. 強制寫入「嚴禁追高」之戰略警語，所有資料僅供收盤後的主戰情室評估。
# 3. 對接 NOCStrategy，大盤紅燈時自動強制休眠，防禦至上。
# 4. 掃描清單從獨立設定檔 stock_scan_list.py 載入，實現單一來源管理。
# 5. 整合市場廣度，動態調整爆量門檻，在量價背離時提高標準並加入警語。
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
from typing import List, Union, Dict, Optional

# =============================================================================
# 日誌設定必須在「匯入 SCAN_LIST」之前，確保 logger 已定義
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

# ===== 從獨立設定檔載入監控清單（logger 已初始化） =====
try:
    from stock_scan_list import SCAN_LIST
except ImportError:
    logger.error("❌ 找不到 stock_scan_list.py，請確保該檔案存在於同目錄下。")
    SCAN_LIST = [] # 避免崩潰

# 🌟 深度引入 NOC 核心防禦模組（含資料庫）
from noc_core import NOCStrategy, NOCDatabase

# =============================================================================
# === 1. 雷達全域組態 ===
# =============================================================================
class LightningConfig:
    MAX_WORKERS : int = int(os.environ.get("MAX_WORKERS", "5"))
    TARGET_FILE : str = "lightning_targets.json"

cfg = LightningConfig()

# =============================================================================
# === 2. 輔助函數：解析 SCAN_LIST ===
# =============================================================================
def parse_scan_list(source: Union[Dict, List]) -> List[str]:
    """將 SCAN_LIST 轉為純股票代號列表，並保留名稱資訊（若為 dict）"""
    if isinstance(source, dict):
        symbols = list(source.keys())
        names = source # 保留對照表
    elif isinstance(source, list):
        symbols = source
        names = None
    else:
        raise TypeError("SCAN_LIST 必須是 dict 或 list")
    return symbols, names

# =============================================================================
# === 3. 籌碼異常與爆量觀測引擎（整合市場廣度聯動） ===
# =============================================================================
def scan_stock_for_anomaly(symbol: str, stock_name: Optional[str] = None,
                           breadth_status: str = "動能持平",
                           divergence_ratio: float = 0.0) -> dict:
    """
    籌碼動能異常觀測器 + 旱地拔蔥 Boss 級突破偵測
    - 一般異常：爆量 2 倍以上，漲幅 ≥ 3%
    - 旱地拔蔥：昨日收盤低於季線、今日站上季線 + 爆量 3 倍以上 + 長紅 ≥ 4%
    - 市場量價背離時，門檻自動提高（爆量 > 2.5 倍，漲幅 ≥ 3.5%，旱地拔蔥爆量 ≥ 3.5 倍）
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
        prev_vol_ma5 = hist['5VMA'].iloc[-2] # 昨日的 5 日均量
        current_close = hist['Close'].iloc[-1]
        prev_close = hist['Close'].iloc[-2]
        current_ma60 = hist['60MA'].iloc[-1]
        prev_ma60 = hist['60MA'].iloc[-2]
        
        # 防呆處理
        if pd.isna(prev_vol_ma5) or prev_vol_ma5 == 0:
            return None
            
        vol_ratio = current_vol / prev_vol_ma5
        price_change_pct = ((current_close - prev_close) / prev_close) * 100
        
        # ===== 動態門檻設定（根據市場廣度） =====
        is_divergence = (breadth_status == "量價背離" or divergence_ratio >= 60.0)
        is_healthy = (breadth_status == "健康放量" or divergence_ratio <= 35.0)

        if is_divergence:
            # 嚴格防誘多模式
            min_vol_ratio_anomaly = 2.5 # 原 2.0
            min_change_pct_anomaly = 3.5 # 原 3.0
            min_vol_ratio_monster = 3.5 # 原 3.0
            market_tip = "🔴 大盤量價背離(誘多盤)，此為高風險雜訊，嚴禁追高！"
        elif is_healthy:
            # 健康放量，維持標準
            min_vol_ratio_anomaly = 2.0
            min_change_pct_anomaly = 3.0
            min_vol_ratio_monster = 3.0
            market_tip = "🟢 大盤健康放量順風，主力實彈點火。"
        else:
            # 動能持平，維持標準
            min_vol_ratio_anomaly = 2.0
            min_change_pct_anomaly = 3.0
            min_vol_ratio_monster = 3.0
            market_tip = "➖ 市場動能持平，正常觀測。"
        
        # ========== 雙階層判定（套用動態門檻） ==========
        # 條件 A：旱地拔蔥 (Boss 級)
        just_crossed_60ma = (current_close > current_ma60) and (prev_close <= prev_ma60)
        is_monster = just_crossed_60ma and (vol_ratio >= min_vol_ratio_monster) and (price_change_pct >= 4.0)
        
        # 條件 B：一般籌碼異常
        is_anomaly = (vol_ratio >= min_vol_ratio_anomaly) and (price_change_pct >= min_change_pct_anomaly)
        
        # 決定名稱：優先使用外部提供的名稱，否則去掉後綴
        name = stock_name if stock_name else symbol.replace(".TW", "").replace(".TWO", "")
        
        if is_monster:
            return {
                "symbol": symbol,
                "name": name,
                "close": round(current_close, 2),
                "vol_ratio": round(vol_ratio, 1),
                "change_pct": round(price_change_pct, 1),
                "tactics": "🔥【旱地拔蔥】底部極端爆量，長紅突破季線起漲！",
                "trello_tip": f"極端爆量 {vol_ratio:.1f} 倍！{market_tip}"
            }
        elif is_anomaly:
            return {
                "symbol": symbol,
                "name": name,
                "close": round(current_close, 2),
                "vol_ratio": round(vol_ratio, 1),
                "change_pct": round(price_change_pct, 1),
                "tactics": "⚡ 籌碼動能異常 (嚴禁追高，僅供波段觀察)",
                "trello_tip": f"爆量 {vol_ratio:.1f} 倍。{market_tip}"
            }
        else:
            return None
            
    except Exception as e:
        # 發生任何錯誤均靜默回傳 None，不影響整體掃描
        return None

# =============================================================================
# === 4. 主控作戰執行緒 (Main Execution) ===
# =============================================================================
if __name__ == "__main__":
    logger.info("⚡ NOC 閃電突擊雷達 v16.0 (市場廣度聯動版) 啟動...")
    start_time = time.time()
    
    # ---- 解析 SCAN_LIST ----
    try:
        symbols, name_map = parse_scan_list(SCAN_LIST)
    except Exception as e:
        logger.error(f"❌ SCAN_LIST 格式錯誤: {e}")
        exit(1)

    if not symbols:
        logger.warning("⚠️ SCAN_LIST 為空，結束程式")
        exit(0)

    # 去重保留順序
    symbols = list(dict.fromkeys(symbols))
    logger.info(f"📋 成功載入 {len(symbols)} 檔標的 (來自 {'dict' if name_map else 'list'})")
    
    # ---- 建立資料庫連線並獲取市場廣度 ----
    db = NOCDatabase()
    strategy = NOCStrategy(db=db)
    macro = strategy.get_macro_status()
    
    # 提取市場廣度指標
    breadth_status = macro.get("breadth_status", "動能持平")
    divergence_ratio = macro.get("divergence_ratio", 0.0)
    breadth_summary = macro.get("market_breadth", "➖ 無廣度數據")
    
    logger.info(f"📊 大盤狀態: {macro['status']} | 市場廣度: {breadth_summary}")
    
    # 🛡️ 總體防禦：大盤紅燈時，直接切斷電源，不觀測任何異常
    if macro["status"] == "🔴 紅燈":
        logger.warning("🚨 大盤跌破季線，閃電觀測儀強制關閉。嚴禁於空頭市場尋找任何多頭火種！")
        with open(cfg.TARGET_FILE, "w", encoding="utf-8") as f:
            json.dump({}, f, ensure_ascii=False, indent=4)
        exit(0)

    found_targets = []
    logger.info(f"📡 啟動多執行緒籌碼觀測，掃描 {len(symbols)} 檔標的是否有主力異常點火足跡...")

    executor = ThreadPoolExecutor(max_workers=cfg.MAX_WORKERS)
    # 傳入廣度狀態
    future_to_symbol = {
        executor.submit(scan_stock_for_anomaly, sym,
                        name_map.get(sym) if name_map else None,
                        breadth_status, divergence_ratio): sym
        for sym in symbols
    }
    
    try:
        for future in as_completed(future_to_symbol, timeout=300):
            sym = future_to_symbol[future]
            try:
                result = future.result()
                if result:
                    found_targets.append(result)
                    logger.info(f"👁️ 觀測到異常籌碼足跡: {sym} | 漲幅: +{result['change_pct']}% | 爆量: {result['vol_ratio']} 倍 | {result['trello_tip']}")
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
