# init_db.py
# NOC 戰情室初始建庫腳本 v2.0
# 功能：從 stock_scan_list 載入監控清單，抓取 400 天歷史 K 線與大盤數據至 noc_warroom.db
# =============================================================================
import datetime
import os
import time
import logging
import re
from typing import List, Union, Dict

from dotenv import load_dotenv

from noc_core import NOCDatabase, NOCDataFetcher

# ===== 日誌設定 =====
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

# ===== 從獨立設定檔載入監控清單 =====
try:
    from stock_scan_list import SCAN_LIST
except ImportError:
    logger.error("❌ 找不到 stock_scan_list.py，請確保該檔案存在於同目錄下。")
    SCAN_LIST = []  # 避免崩潰

load_dotenv()
FINMIND_TOKEN = os.getenv("FINMIND_TOKEN")

# ===== 相容 SCAN_LIST 格式 =====
def parse_scan_list(source: Union[Dict, List]) -> List[str]:
    """將 SCAN_LIST 轉為純股票代號列表，並過濾非法代號"""
    if isinstance(source, dict):
        symbols = list(source.keys())
    elif isinstance(source, list):
        symbols = source
    else:
        raise TypeError("SCAN_LIST 必須是 dict 或 list")

    # 過濾：只保留標準台股代號（含 .TW 或 .TWO，且數字部分為純數字）
    valid_symbols = []
    for sym in symbols:
        if not isinstance(sym, str):
            continue
        # 若已包含 .TW 或 .TWO，直接保留；否則嘗試補充（但我們不自動補充，只保留已包含的）
        # 標準台股代號範例：2330.TW, 5347.TWO
        if '.' in sym:
            parts = sym.split('.')
            if len(parts) == 2 and parts[0].isdigit() and parts[1] in ('TW', 'TWO'):
                valid_symbols.append(sym)
            else:
                logger.warning(f"⚠️ 跳過非標準代號: {sym}")
        else:
            # 若無後綴，假設為 TW（但原清單大多有後綴，此處保留但不自動添加）
            logger.warning(f"⚠️ 跳過無後綴代號: {sym}，請補上 .TW 或 .TWO")
            # 可選擇自動添加，但為嚴謹起見，跳過
    return list(dict.fromkeys(valid_symbols))  # 去重保留順序

if __name__ == "__main__":
    logger.info("🚀 啟動 NOC 戰情室「建庫大補丸」歷史資料載入作業 (模組化版)...")
    
    # ---- 解析 SCAN_LIST ----
    try:
        target_stocks = parse_scan_list(SCAN_LIST)
    except Exception as e:
        logger.error(f"❌ SCAN_LIST 格式錯誤: {e}")
        exit(1)

    if not target_stocks:
        logger.warning("⚠️ SCAN_LIST 為空或無效代號，結束程式")
        exit(0)

    logger.info(f"🎯 成功讀取監控清單！最終鎖定 {len(target_stocks)} 檔股票，準備進行 400 天歷史大補給！")
    logger.info("⚠️ 預計耗時 15-25 分鐘，請耐心等候...")

    db = NOCDatabase()
    fetcher = NOCDataFetcher(token=FINMIND_TOKEN)
    start_date = (datetime.datetime.now() - datetime.timedelta(days=400)).strftime("%Y-%m-%d")

    # ---- 1. 更新大盤指數 ----
    try:
        logger.info("📈 正在下載大盤與防空歷史數據 (近400天)...")
        fetcher.fetch_market_health_data(start_date, db)
    except Exception as e:
        logger.error(f"大盤更新失敗: {e}")

    # ---- 2. 逐檔下載個股歷史數據 ----
    total = len(target_stocks)
    success_count = 0
    fail_list = []
    start_time = time.time()

    for i, sym in enumerate(target_stocks, 1):
        try:
            logger.info(f"[{i}/{total}] 正在抓取 {sym} 的歷史戰情數據...")
            fetcher.fetch_and_store_stock_data(sym, start_date, db)
            success_count += 1
            # 間隔 0.8~1.0 秒，避免 API 限流
            time.sleep(0.8 + 0.2 * (i % 3))  # 隨機延遲
        except Exception as e:
            fail_list.append(sym)
            logger.error(f"[{i}/{total}] ❌ {sym} 下載失敗: {e}")
            # 仍延遲，避免連續失敗
            time.sleep(1.0)

    elapsed = time.time() - start_time
    logger.info("\\n✅ 歷史戰情資料庫 (SQLite) 初始建置與灌水完成！")
    logger.info(f"   📊 總標的數: {total}")
    logger.info(f"   ✅ 成功: {success_count}")
    logger.info(f"   ❌ 失敗: {len(fail_list)}")
    if fail_list:
        logger.info(f"   📋 失敗清單: {', '.join(fail_list)}")
    logger.info(f"   ⏱️ 總耗時: {elapsed:.1f} 秒")

    # ---- 統計資料庫實存標的數 ----
    try:
        import sqlite3
        with sqlite3.connect(db.db_path) as conn:
            cur = conn.cursor()
            cur.execute("SELECT COUNT(DISTINCT symbol) FROM stock_prices")
            distinct_count = cur.fetchone()[0]
        logger.info(f"   💾 資料庫目前實存不重複標的：{distinct_count} 檔")
    except:
        pass
