# =============================================================================
# NOC 戰情室資料庫每日補給腳本 (update_db.py) v2.0
# 功能：每日定時排程執行，同步 stock_scan_list 進行最新行情與大盤增量更新
# =============================================================================

import datetime
import os
import time
import logging
import sqlite3
import random
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Union, Dict, Tuple
import yfinance as yf
from dotenv import load_dotenv

from noc_core import NOCDatabase, NOCDataFetcher

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)
logging.getLogger('yfinance').setLevel(logging.CRITICAL)

load_dotenv()
FINMIND_TOKEN = os.getenv("FINMIND_TOKEN", "")

# ===== 同一來源載入 =====
try:
    from stock_scan_list import SCAN_LIST
except ImportError:
    logger.error("❌ 找不到 stock_scan_list.py，請確認檔案存在。")
    SCAN_LIST = []

def parse_scan_list(source: Union[Dict, List]) -> List[str]:
    if isinstance(source, dict):
        return list(source.keys())
    elif isinstance(source, list):
        return source
    else:
        raise TypeError("SCAN_LIST 必須是 dict 或 list")

def fetch_stock_daily(sym: str, start_date: str, fetcher: NOCDataFetcher, db: NOCDatabase) -> Tuple[bool, int]:
    """每日增量更新（透過 INSERT OR REPLACE 自動覆蓋最新行情）"""
    try:
        fetcher.fetch_and_store_stock_data(sym, start_date, db)
    except Exception as e:
        logger.warning(f"⚠️ {sym} 透過預設 Fetcher 抓取失敗: {e}")

    try:
        df = db.get_stock_dataframe(sym, days=5)
        if df is None or df.empty:
            logger.info(f"🔄 觸發 yfinance 備援補給: {sym}")
            ticker = yf.Ticker(sym)
            hist = ticker.history(period="8mo")
            if not hist.empty:
                with sqlite3.connect(db.db_path) as conn:
                    for idx, row in hist.iterrows():
                        date_str = idx.strftime("%Y-%m-%d")
                        conn.execute('''
                            INSERT OR REPLACE INTO stock_prices (symbol, date, open, high, low, close, volume, adj_close)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (sym, date_str, row['Open'], row['High'], row['Low'], row['Close'], int(row['Volume']), row['Close']))
                info = ticker.info
                shares_out = info.get("sharesOutstanding") or info.get("impliedSharesOutstanding")
                if shares_out:
                    db.save_shares_out(sym, shares_out)
            else:
                return False, 0

        with sqlite3.connect(db.db_path) as conn:
            cnt = conn.execute("SELECT COUNT(*) FROM stock_prices WHERE symbol = ?", (sym,)).fetchone()[0]
        return True, cnt
    except Exception as e:
        logger.error(f"❌ 更新 {sym} 發生例外: {e}")
        return False, 0

if __name__ == "__main__":
    logger.info("🚀 開始執行 NOC 每日盤後戰情資料庫補給作業...")

    symbols = parse_scan_list(SCAN_LIST)
    symbols = list(dict.fromkeys(symbols))

    if not symbols:
        logger.warning("⚠️ SCAN_LIST 為空，結束作業")
        exit(0)

    logger.info(f"📊 鎖定 {len(symbols)} 檔目標，進行每日行情增量補給！")

    db = NOCDatabase("noc_warroom.db")
    fetcher = NOCDataFetcher(token=FINMIND_TOKEN)
    start_date = (datetime.datetime.now() - datetime.timedelta(days=240)).strftime("%Y-%m-%d")

    # 更新大盤
    try:
        logger.info("📈 正在更新大盤與市場海象數據...")
        fetcher.fetch_market_health_data(start_date, db)
    except Exception as e:
        logger.error(f"大盤更新失敗: {e}")

    # 更新個股
    start_time = time.time()
    success_count = 0
    fail_list = []

    def worker(sym):
        time.sleep(random.uniform(0.05, 0.2))
        return sym, fetch_stock_daily(sym, start_date, fetcher, db)

    with ThreadPoolExecutor(max_workers=6) as executor:
        futures = {executor.submit(worker, sym): sym for sym in symbols}
        for idx, f in enumerate(as_completed(futures), 1):
            sym, (ok, cnt) = f.result()
            if ok:
                success_count += 1
                logger.info(f"[{idx}/{len(symbols)}] ✅ {sym} 更新成功 (總計 {cnt} 筆)")
            else:
                fail_list.append(sym)
                logger.warning(f"[{idx}/{len(symbols)}] ❌ {sym} 更新失敗")

    elapsed = time.time() - start_time
    logger.info("=" * 60)
    logger.info("🎉 每日戰情資料庫補給完畢！")
    logger.info(f" 📊 總目標數: {len(symbols)} | 成功: {success_count} | 失敗: {len(fail_list)}")
    if fail_list:
        logger.info(f" 📋 失敗標的: {', '.join(fail_list)}")
    logger.info(f" ⏱️ 總耗時: {elapsed:.1f} 秒")

    with sqlite3.connect(db.db_path) as conn:
        distinct_count = conn.execute("SELECT COUNT(DISTINCT symbol) FROM stock_prices").fetchone()[0]
    logger.info(f" 💾 資料庫實存有效標的：{distinct_count} 檔")
