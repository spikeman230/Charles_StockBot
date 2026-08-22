# =============================================================================
# NOC 決策回顧引擎 (Decision Review Engine) v2.0 (SQLite 高速本地版)
# 用途：驗證過去決策的準確性，統計各訊號的勝率與報酬
# 執行方式：python noc_review_engine.py
# 輸出：decision_review.csv + 統計報告（終端機顯示）
# =============================================================================

import pandas as pd
import numpy as np
import sqlite3
from datetime import datetime, timedelta
import os
import sys
import logging
from pathlib import Path
from typing import Dict, List, Optional
import time

# === 組態 ===
DB_PATH = "noc_warroom.db"
LOG_FILE = "noc_review.log"
REVIEW_CSV = "decision_review.csv"
N_DAYS_LIST = [1, 5, 10, 20]           # 檢視多個時間框架
STOP_LOSS_PCT = 0.95                   # 假設停損為買入價的 -5%

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# ============================================================================
# 1. 讀取原始決策日誌
# ============================================================================
def load_decision_log(csv_path: str = "noc_trading_log.csv") -> pd.DataFrame:
    """讀取 CSV，並確保日期格式正確"""
    if not Path(csv_path).exists():
        logger.error(f"找不到 {csv_path}，請確認檔案路徑")
        sys.exit(1)
    df = pd.read_csv(csv_path, encoding="utf-8-sig")
    df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
    df = df.dropna(subset=["日期"])
    df = df.sort_values("日期").reset_index(drop=True)
    logger.info(f"成功讀取 {len(df)} 筆決策紀錄，日期範圍 {df['日期'].min().date()} ~ {df['日期'].max().date()}")
    return df

# ============================================================================
# 2. 本地預載所有歷史股價 (In-Memory 快速比對)
# ============================================================================
def preload_price_data(db_path: str = DB_PATH) -> pd.DataFrame:
    """一次性從本地 SQLite 載入所有歷史價格，建立記憶體索引"""
    if not Path(db_path).exists():
        logger.error(f"找不到本地資料庫 {db_path}，請先執行 update_db.py")
        sys.exit(1)

    logger.info("⚡ 正在從本地 SQLite 資料庫載入全市場歷史價格...")
    conn = sqlite3.connect(db_path)
    query = """
        SELECT symbol, date, open as Open, high as High, low as Low, close as Close, volume as Volume
        FROM stock_prices
        ORDER BY symbol, date ASC
    """
    df_prices = pd.read_sql_query(query, conn)
    conn.close()

    df_prices["date"] = pd.to_datetime(df_prices["date"])
    logger.info(f"✅ 成功載入 {len(df_prices)} 筆歷史 K 線數據！")
    return df_prices

# ============================================================================
# 3. 分析單一決策的績效 (純本地計算)
# ============================================================================
def analyze_decision_fast(row: pd.Series, df_prices: pd.DataFrame) -> Optional[Dict]:
    """利用預載的價格資料計算多個時間框架的績效"""
    symbol = row["代號"]
    decision_date = row["日期"]
    entry_price = row["收盤價"]

    if pd.isna(entry_price) or entry_price <= 0:
        return None

    # 抓取該標的在決策日之後的所有交易日
    stock_future = df_prices[
        (df_prices["symbol"] == symbol) & 
        (df_prices["date"] > decision_date)
    ].sort_values("date")

    if stock_future.empty:
        return None

    result = {
        "日期": decision_date.strftime("%Y-%m-%d"),
        "代號": symbol,
        "名稱": row.get("名稱", symbol),
        "收盤價": entry_price,
        "RSI": row.get("RSI", None),
        "戰場預判": row.get("戰場預判", "未知"),
        "籌碼訊號": row.get("籌碼訊號", "未知"),
        "行動指令": row.get("行動指令", "未知"),
    }

    # 計算各時間框架指標
    for n in N_DAYS_LIST:
        if len(stock_future) >= n:
            window = stock_future.iloc[:n]
            high = window["High"].max()
            low = window["Low"].min()
            final_price = window["Close"].iloc[-1]
            
            max_gain = (high - entry_price) / entry_price
            max_loss = (low - entry_price) / entry_price
            final_return = (final_price - entry_price) / entry_price
            hit_stop = low < (entry_price * STOP_LOSS_PCT)
        else:
            max_gain = max_loss = final_return = np.nan
            hit_stop = False

        result[f"{n}D_最高漲幅%"] = round(max_gain * 100, 2) if not np.isnan(max_gain) else None
        result[f"{n}D_最大跌幅%"] = round(max_loss * 100, 2) if not np.isnan(max_loss) else None
        result[f"{n}D_最終漲幅%"] = round(final_return * 100, 2) if not np.isnan(final_return) else None
        result[f"{n}D_觸及停損"] = hit_stop

    return result

# ============================================================================
# 4. 主流程
# ============================================================================
def main():
    start_time = time.time()
    logger.info("🚀 啟動 NOC 決策回顧引擎 (SQLite 本地極速版)...")
    
    df_decisions = load_decision_log()
    df_prices = preload_price_data()

    # 過濾具操作意義的決策
    interesting_actions = ["建倉", "試單", "波段", "佈局", "長線鎖籌", "加碼", "扣款", "獲利巡航", "浮虧防禦", "洗盤耐受", "戰術撤離"]
    df_target = df_decisions[df_decisions["行動指令"].astype(str).str.contains("|".join(interesting_actions), na=False)]
    logger.info(f"🎯 共篩選出 {len(df_target)} 筆具操作意義的決策，開始批次運算...")

    if len(df_target) == 0:
        logger.warning("沒有符合條件的決策，結束")
        return

    results = []
    for _, row in df_target.iterrows():
        res = analyze_decision_fast(row, df_prices)
        if res:
            results.append(res)

    if not results:
        logger.warning("無任何決策可分析，可能本地資料庫缺乏未來的價格數據")
        return

    df_results = pd.DataFrame(results)
    df_results.to_csv(REVIEW_CSV, index=False, encoding="utf-8-sig")
    logger.info(f"📁 已儲存詳細回顧結果至 {REVIEW_CSV}")

    # ============ 產生統計摘要 ============
    print("\n" + "="*60)
    print("NOC 戰情室決策績效統計摘要")
    print("="*60)

    # 依「戰場預判」分類統計
    categories = df_results["戰場預判"].unique()
    for cat in categories:
        if pd.isna(cat):
            continue
        subset = df_results[df_results["戰場預判"] == cat]
        print(f"\n📊 訊號類別：{cat} (共 {len(subset)} 筆)")
        for n in N_DAYS_LIST:
            col_final = f"{n}D_最終漲幅%"
            if col_final in subset:
                valid_subset = subset.dropna(subset=[col_final])
                if len(valid_subset) == 0:
                    continue
                win_count = (valid_subset[col_final] > 0).sum()
                avg_return = valid_subset[col_final].mean()
                hit_stop_pct = valid_subset[f"{n}D_觸及停損"].mean() * 100
                print(f"   {n}日後：勝率 {win_count/len(valid_subset)*100:.1f}% | 平均報酬 {avg_return:.2f}% | 停損觸及率 {hit_stop_pct:.1f}%")

    # 總體統計
    print("\n" + "─"*40)
    print("📈 總體績效（所有決策合併）")
    for n in N_DAYS_LIST:
        col_final = f"{n}D_最終漲幅%"
        if col_final in df_results:
            valid_df = df_results.dropna(subset=[col_final])
            if len(valid_df) > 0:
                avg = valid_df[col_final].mean()
                median = valid_df[col_final].median()
                win_rate = (valid_df[col_final] > 0).mean() * 100
                hit_stop = valid_df[f"{n}D_觸及停損"].mean() * 100
                print(f"   {n}日後：平均 {avg:.2f}% | 中位數 {median:.2f}% | 勝率 {win_rate:.1f}% | 停損觸及率 {hit_stop:.1f}%")

    logger.info(f"✨ 回顧完成！總耗時: {time.time() - start_time:.2f} 秒")

if __name__ == "__main__":
    main()
