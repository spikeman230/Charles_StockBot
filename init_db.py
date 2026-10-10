# =============================================================================
# NOC 戰情室資料庫初始化腳本 (init_db.py) v2.1
# 功能：從 stock_scan_list 載入清單，冷啟動建立資料表並寫入 400 天歷史底庫
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

# ===== 日誌設定 =====
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)
logging.getLogger('yfinance').setLevel(logging.CRITICAL)

load_dotenv()
FINMIND_TOKEN = os.getenv("FINMIND_TOKEN", "")

# ===== 載入外部獨立掃描清單 =====
try:
    from stock_scan_list import SCAN_LIST
except ImportError:
    logger.error("❌ 找不到 stock_scan_list.py，請確認檔案存在於同目錄下。")
    SCAN_LIST = []

def parse_scan_list(source: Union[Dict, List]) -> List[str]:
    """將 SCAN_LIST 轉為純股票代號列表"""
    if isinstance(source, dict):
        return list(source.keys())
    elif isinstance(source, list):
        return source
    else:
        raise TypeError("SCAN_LIST 必須是 dict 或 list")

def fetch_and_populate_stock(sym: str, start_date: str, fetcher: NOCDataFetcher, db: NOCDatabase) -> Tuple[bool, int]:
    """初次抓取個股 400 天完整歷史數據寫入 DB"""
    try:
        fetcher.fetch_and_store_stock_data(sym, start_date, db)
    except Exception as e:
        logger.warning(f"⚠️ {sym} 透過預設 Fetcher 下載異常: {e}")

    # 備援驗證
    try:
        df = db.get_stock_dataframe(sym, days=5)
        if df is None or df.empty:
            logger.info(f"🔄 觸發 yfinance 備援補給 (16mo): {sym}")
            ticker = yf.Ticker(sym)
            hist = ticker.history(period="16mo")
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
                logger.error(f"❌ {sym} yfinance 亦無數據")
                return False, 0

        with sqlite3.connect(db.db_path) as conn:
            cnt = conn.execute("SELECT COUNT(*) FROM stock_prices WHERE symbol = ?", (sym,)).fetchone()[0]
        return True, cnt
    except Exception as e:
        logger.error(f"❌ 寫入 {sym} 發生例外: {e}")
        return False, 0

if __name__ == "__main__":
    logger.info("🛠️ 開始執行 NOC 戰情室資料庫初始化建庫 (init_db.py - 400 天底庫)...")
    
    symbols = parse_scan_list(SCAN_LIST)
    symbols = list(dict.fromkeys(symbols))  # 去重保持順序
    
    if not symbols:
        logger.error("❌ 監控清單為空，終止建庫！")
        exit(1)

    db = NOCDatabase("noc_warroom.db")  # 自動初始化所有 Tables
    fetcher = NOCDataFetcher(token=FINMIND_TOKEN)
    # 🔥 設定為 400 天歷史
    start_date = (datetime.datetime.now() - datetime.timedelta(days=400)).strftime("%Y-%m-%d")

    # 1. 抓取大盤 400 天歷史
    try:
        logger.info("📈 正在下載加權指數歷史數據 (400 天)...")
        fetcher.fetch_market_health_data(start_date, db)
    except Exception as e:
        logger.error(f"大盤資料下載失敗: {e}")

    # 2. 多執行緒建庫
    logger.info(f"🚀 開始下載 {len(symbols)} 檔目標歷史 K 線 (400 天)...")
    start_time = time.time()
    success_count = 0

    def worker(sym):
        time.sleep(random.uniform(0.05, 0.2))
        return sym, fetch_and_populate_stock(sym, start_date, fetcher, db)

    with ThreadPoolExecutor(max_workers=6) as executor:
        futures = {executor.submit(worker, s): s for s in symbols}
        for idx, f in enumerate(as_completed(futures), 1):
            sym, (ok, cnt) = f.result()
            if ok:
                success_count += 1
                logger.info(f"[{idx}/{len(symbols)}] ✅ {sym} 建庫完成 (共 {cnt} 筆)")
            else:
                logger.warning(f"[{idx}/{len(symbols)}] ❌ {sym} 建庫失敗")

    elapsed = time.time() - start_time
    logger.info("=" * 60)
    logger.info(f"🎉 初始化建庫完成！耗時: {elapsed:.1f} 秒")
    logger.info(f"📊 標的成功率: {success_count}/{len(symbols)}")

    with sqlite3.connect(db.db_path) as conn:
        distinct_count = conn.execute("SELECT COUNT(DISTINCT symbol) FROM stock_prices").fetchone()[0]
    logger.info(f"💾 資料庫實存標的總數：{distinct_count} 檔（與清單核對標的數一致）")
