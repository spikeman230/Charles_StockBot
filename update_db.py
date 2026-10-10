# update_db.py
# NOC 戰情室資料庫每日補給腳本 v3.1
# 功能：從 stock_scan_list 載入監控清單，每日增量抓取最近 5 天；首次建庫抓 400 天
# 整合：單一來源維護 + 智慧增量 + 寬容失敗判定 + yfinance 備援 + 涵蓋率驗證
# =============================================================================
import datetime
import os
import time
import logging
import sqlite3
import random
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Tuple

import yfinance as yf
from dotenv import load_dotenv

from noc_core import NOCDatabase, NOCDataFetcher

# ===== 從獨立設定檔載入監控清單（單一來源維護） =====
try:
    from stock_scan_list import SCAN_LIST, parse_scan_list
except ImportError as exc:
    logging.error("❌ 找不到 stock_scan_list.py 或清單解析器。")
    raise SystemExit(1) from exc

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

# =============================================================================
# 智慧起算日決定（DB 有資料 → 5 天；DB 無資料 → 400 天首次建庫）
# =============================================================================
def get_stock_fetch_start_date(symbol: str, db: NOCDatabase,
                               initial_start_date: str, daily_start_date: str) -> str:
    """
    決定該檔標的的抓取起算日：
    - DB 已有資料 → 返回每日增量起算日（近 5 天）
    - DB 無資料   → 返回首次建庫起算日（近 400 天）
    """
    try:
        with sqlite3.connect(db.db_path) as conn:
            row = conn.execute(
                "SELECT MAX(date) FROM stock_prices WHERE symbol = ?", (symbol,)
            ).fetchone()
        if row and row[0]:
            # DB 已有資料，走每日增量（近 5 天）
            return daily_start_date
    except Exception:
        pass
    # DB 無資料，走首次建庫（近 400 天）
    return initial_start_date

# =============================================================================
# 雙重保障抓取機制（增量 5 天 + 寬容判定 + yfinance 備援）
# =============================================================================
def fetch_stock_robust(sym: str, start_date: str, fetcher: NOCDataFetcher, db: NOCDatabase) -> Tuple[bool, int]:
    """
    抓取個股資料（依 start_date 決定範圍）：
    - 第一優先：FinMind（NOCDataFetcher）
    - 寬容判定：不論 FinMind 是否成功，只要 DB 有資料即視為成功
    - yfinance 備援：若 DB 完全無資料，改用 yfinance 補給
    回傳 (是否成功, DB 目前總筆數)
    """
    # ---- 第一層：嘗試 FinMind ----
    try:
        fetcher.fetch_and_store_stock_data(sym, start_date, db)
    except Exception as e:
        logger.warning(f"⚠️ {sym} 透過 FinMind 抓取失敗: {e}")

    # ---- 第二層：驗證 DB 是否有資料，若無則觸發 yfinance 備援 ----
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
                            INSERT OR REPLACE INTO stock_prices
                            (symbol, date, open, high, low, close, volume, adj_close)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (sym, date_str, row['Open'], row['High'], row['Low'],
                              row['Close'], int(row['Volume']), row['Close']))
                # 補充股本
                try:
                    info = ticker.info
                    shares_out = info.get("sharesOutstanding") or info.get("impliedSharesOutstanding")
                    if shares_out:
                        db.save_shares_out(sym, shares_out)
                except Exception:
                    pass
            else:
                logger.error(f"❌ {sym} yfinance 亦無數據（可能代號錯誤或已下市）")
                return False, 0

        # ---- 第三層：統計最終筆數 ----
        with sqlite3.connect(db.db_path) as conn:
            cnt = conn.execute(
                "SELECT COUNT(*) FROM stock_prices WHERE symbol = ?", (sym,)
            ).fetchone()[0]

        # 寬容判定：只要 DB 有資料即視為成功
        if cnt > 0:
            return True, cnt
        else:
            return False, 0

    except Exception as e:
        logger.error(f"❌ 寫入 {sym} 發生例外: {e}")
        return False, 0

# =============================================================================
# 多執行緒執行器（含完整統計）
# =============================================================================
def update_all_stocks(symbols: List[str],
                      initial_start_date: str,
                      daily_start_date: str,
                      max_workers: int = 6) -> dict:
    """
    並行更新所有標的，回傳統計結果。
    - 每檔標的依 DB 狀態自動選擇起算日（5 天或 400 天）
    """
    db = NOCDatabase("noc_warroom.db")
    fetcher = NOCDataFetcher(token=FINMIND_TOKEN)
    total = len(symbols)
    success_count = 0
    fail_list = []
    total_records = 0

    logger.info(f"🚀 啟動多執行緒（{max_workers} 個 worker）每日增量補給 {total} 檔標的...")

    def worker(sym):
        # 加入隨機延遲，避免 Rate Limit
        time.sleep(random.uniform(0.05, 0.2))
        # 智慧決定起算日
        sym_start = get_stock_fetch_start_date(sym, db, initial_start_date, daily_start_date)
        ok, cnt = fetch_stock_robust(sym, sym_start, fetcher, db)
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
                    logger.info(f"[{idx}/{total}] ✅ {sym} 更新成功（總計 {cnt} 筆）")
                else:
                    fail_list.append(sym)
                    logger.warning(f"[{idx}/{total}] ❌ {sym} 更新失敗")
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
    logger.info("🚀 開始執行 NOC 盤後戰情資料庫每日補給作業 (增量 5 天 / 首次 400 天)...")

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
    logger.info(f"📊 鎖定 {len(symbols)} 檔目標（不重複），準備開始每日補給！")

    db = NOCDatabase("noc_warroom.db")
    fetcher = NOCDataFetcher(token=FINMIND_TOKEN)

    # ---- 三種起算日 ----
    market_start_date = (datetime.datetime.now() - datetime.timedelta(days=240)).strftime("%Y-%m-%d")
    # 首次建庫（DB 無資料）：抓 400 天
    initial_start_date = (datetime.datetime.now() - datetime.timedelta(days=400)).strftime("%Y-%m-%d")
    # 每日增量（DB 已有資料）：抓最近 5 天
    daily_start_date = (datetime.datetime.now() - datetime.timedelta(days=5)).strftime("%Y-%m-%d")

    # ---- 更新大盤指數 ----
    try:
        logger.info("📈 正在更新大盤指數與市場海象數據 (240 天)...")
        fetcher.fetch_market_health_data(market_start_date, db)
    except Exception as e:
        logger.error(f"大盤更新失敗: {e}")

    # ---- 更新個股（依 DB 狀態自動選擇 5 天或 400 天）----
    start_time = time.time()
    stats = update_all_stocks(symbols, initial_start_date, daily_start_date, max_workers=6)
    elapsed = time.time() - start_time

    # =========================================================================
    # 涵蓋率驗證（清單 vs 資料庫）
    # =========================================================================
    logger.info("=" * 60)
    logger.info("🔍 正在執行資料庫涵蓋率驗證...")

    with sqlite3.connect(db.db_path) as conn:
        db_symbols = {
            row[0] for row in conn.execute(
                "SELECT DISTINCT symbol FROM stock_prices"
            )
        }
        distinct_count = len(db_symbols)

    target_symbols = set(symbols)
    missing_symbols = sorted(target_symbols - db_symbols)
    extra_symbols = sorted(db_symbols - target_symbols)
    covered_count = len(target_symbols & db_symbols)

    # =========================================================================
    # 完整統計報告
    # =========================================================================
    logger.info("=" * 60)
    logger.info("🎉 每日戰情資料庫補給完畢！")
    logger.info(f"   ⏱️ 總耗時: {elapsed:.1f} 秒")
    logger.info(f"   📊 監控清單總數: {stats['total']} 檔")
    logger.info(f"   ✅ 成功更新: {stats['success']} 檔")
    logger.info(f"   ❌ 下載失敗: {stats['fail']} 檔")
    logger.info(f"   📦 本次總資料筆數: {stats['total_records']} 筆")
    logger.info("-" * 60)
    logger.info(f"   💾 資料庫實存標的：{distinct_count} 檔")
    logger.info(f"   🎯 清單涵蓋率：{covered_count}/{stats['total']} 檔 ({covered_count/stats['total']*100:.1f}%)")

    if stats['fail_list']:
        logger.error(f"   ❌ 下載失敗清單 ({len(stats['fail_list'])} 檔): {', '.join(stats['fail_list'])}")
    if missing_symbols:
        logger.error(f"   ❌ 資料庫仍缺少清單股票 ({len(missing_symbols)} 檔): {', '.join(missing_symbols)}")
    if extra_symbols:
        logger.warning(
            f"   ⚠️ 資料庫另含 {len(extra_symbols)} 檔不在 SCAN_LIST 的舊標的（保留歷史資料，未刪除）: "
            f"{', '.join(extra_symbols[:10])}{' ...' if len(extra_symbols) > 10 else ''}"
        )

    logger.info("=" * 60)

    # =========================================================================
    # CI/CD 品質判定（若有失敗或缺漏則標記失敗）
    # =========================================================================
    if stats["fail"] or missing_symbols:
        raise SystemExit(1)
