# update_db.py
# NOC 戰情室資料庫補給腳本 v2.0
# 功能：從 stock_scan_list 載入監控清單，並行下載歷史 K 線與股本資訊至 noc_warroom.db
# =============================================================================
import datetime
import os
import time
import re
import logging
import sqlite3
import random
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Union, Dict, Tuple

import yfinance as yf
from dotenv import load_dotenv

from noc_core import NOCDatabase, NOCDataFetcher

# ===== 從獨立設定檔載入監控清單 =====
try:
    from stock_scan_list import SCAN_LIST
except ImportError:
    logging.error("❌ 找不到 stock_scan_list.py，請確保該檔案存在於同目錄下。")
    SCAN_LIST = [] # 避免崩潰

# ===== 日誌設定 =====
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

load_dotenv()
FINMIND_TOKEN = os.getenv("FINMIND_TOKEN")

# ===== 相容 SCAN_LIST 格式 =====
def parse_scan_list(source: Union[Dict, List]) -> List[str]:
    """將 SCAN_LIST 轉為純股票代號列表"""
    if isinstance(source, dict):
        return list(source.keys())
    elif isinstance(source, list):
        return source
    else:
        raise TypeError("SCAN_LIST 必須是 dict 或 list")

# =============================================================================
# 雙重保障抓取機制 (支援並行)
# =============================================================================
def fetch_stock_robust(sym: str, start_date: str, fetcher: NOCDataFetcher, db: NOCDatabase) -> Tuple[bool, int]:
    """
    個股數據抓取，回傳 (是否成功, 資料筆數)
    """
    try:
        # 1. 嘗試透過 NOCDataFetcher (FinMind) 抓取
        fetcher.fetch_and_store_stock_data(sym, start_date, db)
    except Exception as e:
        logger.warning(f"⚠️ {sym} 透過預設 Fetcher 抓取失敗: {e}")

    # 2. 驗證資料庫是否有數據，若無則啟動 yfinance 備援
    try:
        df = db.get_stock_dataframe(sym, days=5)
        if df is None or df.empty:
            logger.info(f"🔄 觸發 yfinance 備援機制補給: {sym}")
            ticker = yf.Ticker(sym)
            # 抓取 8 個月歷史，確保資料充足
            hist = ticker.history(period="8mo")
            if not hist.empty:
                with sqlite3.connect(db.db_path) as conn:
                    for idx, row in hist.iterrows():
                        date_str = idx.strftime("%Y-%m-%d")
                        conn.execute('''
                            INSERT OR REPLACE INTO stock_prices (symbol, date, open, high, low, close, volume, adj_close)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (sym, date_str, row['Open'], row['High'], row['Low'], row['Close'], int(row['Volume']), row['Close']))
                # 補充股本
                try:
                    info = ticker.info
                    shares_out = info.get("sharesOutstanding") or info.get("impliedSharesOutstanding")
                    if shares_out:
                        db.save_shares_out(sym, shares_out)
                except:
                    pass
                logger.info(f"✅ {sym} 透過 yfinance 成功補給寫入資料庫！")
            else:
                logger.error(f"❌ {sym} Yahoo Finance 亦無數據（可能代號不符或興櫃/暫停交易）")
                return False, 0
        else:
            logger.debug(f"✅ {sym} 已存在資料庫中，跳過補給。")
    except Exception as e:
        logger.error(f"❌ 備援寫入 {sym} 時發生錯誤: {e}")
        return False, 0

    # 統計最終筆數
    try:
        with sqlite3.connect(db.db_path) as conn:
            cur = conn.execute("SELECT COUNT(*) FROM stock_prices WHERE symbol = ?", (sym,))
            count = cur.fetchone()[0]
        return True, count
    except:
        return True, 0

# =============================================================================
# 多執行緒執行器
# =============================================================================
def update_all_stocks(symbols: List[str], start_date: str, max_workers: int = 6) -> Dict[str, int]:
    """
    並行更新所有標的，回傳統計結果
    """
    db = NOCDatabase()
    fetcher = NOCDataFetcher(token=FINMIND_TOKEN)
    total = len(symbols)
    success_count = 0
    fail_list = []
    total_records = 0

    logger.info(f"🚀 啟動多執行緒（{max_workers} 個 worker）更新 {total} 檔標的...")

    def worker(sym):
        # 加入隨機延遲，避免 Rate Limit
        time.sleep(random.uniform(0.05, 0.2))
        ok, cnt = fetch_stock_robust(sym, start_date, fetcher, db)
        return sym, ok, cnt

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(worker, sym): sym for sym in symbols}
        for idx, future in enumerate(as_completed(futures), 1):
            sym = futures[future]
            try:
                sym, ok, cnt = future.result()
                if ok:
                    success_count += 1
                    total_records += cnt
                    logger.info(f"[{idx}/{total}] ✅ {sym} 寫入成功，共 {cnt} 筆數據")
                else:
                    fail_list.append(sym)
                    logger.warning(f"[{idx}/{total}] ❌ {sym} 寫入失敗")
            except Exception as e:
                fail_list.append(sym)
                logger.error(f"[{idx}/{total}] ❌ {sym} 發生例外: {e}")

    return {
        "total": total,
        "success": success_count,
        "fail": len(fail_list),
        "fail_list": fail_list,
        "total_records": total_records,
    }
# =============================================================================
# 主程式
# =============================================================================
if __name__ == "__main__":
    logger.info("🚀 開始執行 NOC 盤後戰情資料庫補給作業 (多執行緒版)...")

    # ---- 解析 SCAN_LIST ----
    try:
        symbols = parse_scan_list(SCAN_LIST)
    except Exception as e:
        logger.error(f"❌ SCAN_LIST 格式錯誤: {e}")
        exit(1)

    if not symbols:
        logger.warning("⚠️ SCAN_LIST 為空，結束程式")
        exit(0)

    # 去重保留順序
    symbols = list(dict.fromkeys(symbols))
    logger.info(f"📊 鎖定 {len(symbols)} 檔目標（不重複），準備開始日常補給！")

    db = NOCDatabase()
    fetcher = NOCDataFetcher(token=FINMIND_TOKEN)
    start_date = (datetime.datetime.now() - datetime.timedelta(days=240)).strftime("%Y-%m-%d")

    # ---- 更新大盤指數 ----
    try:
        logger.info("📈 正在更新大盤指數與市場海象數據...")
        fetcher.fetch_market_health_data(start_date, db)
    except Exception as e:
        logger.error(f"大盤更新失敗: {e}")

    # ---- 更新個股 ----
    start_time = time.time()
    stats = update_all_stocks(symbols, start_date, max_workers=6)

    elapsed = time.time() - start_time
    logger.info(f"\n🎉 戰情資料庫補給完畢！")
    logger.info(f" 📊 總標的數: {stats['total']}")
    logger.info(f" ✅ 成功: {stats['success']}")
    logger.info(f" ❌ 失敗: {stats['fail']}")
    if stats['fail_list']:
        logger.info(f" 📋 失敗清單: {', '.join(stats['fail_list'])}")
    logger.info(f" ⏱️ 總耗時: {elapsed:.1f} 秒")
    logger.info(f" 📦 總資料筆數: {stats['total_records']}")

    # ---- 統計資料庫實存標的數 ----
    try:
        with sqlite3.connect(db.db_path) as conn:
            cur = conn.cursor()
            cur.execute("SELECT COUNT(DISTINCT symbol) FROM stock_prices")
            distinct_count = cur.fetchone()[0]
        logger.info(f" 💾 資料庫目前實存不重複標的：{distinct_count} 檔")
    except:
        pass
