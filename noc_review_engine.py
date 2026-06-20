# =============================================================================
# NOC 決策回顧引擎 (Decision Review Engine) v1.0
# 用途：驗證過去決策的準確性，統計各訊號的勝率與報酬
# 執行方式：python noc_review_engine.py
# 輸出：decision_review.csv + 統計報告（終端機顯示）
# =============================================================================

import pandas as pd
import numpy as np
import yfinance as yf
from datetime import datetime, timedelta
import os
import sys
import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import time

# === 組態 ===
LOG_FILE = "noc_review.log"
REVIEW_CSV = "decision_review.csv"
N_DAYS_LIST = [1, 5, 10, 20]           # 檢視多個時間框架
STOP_LOSS_PCT = 0.95                   # 假設停損為買入價的 -5%
SLEEP_BETWEEN_REQUESTS = 0.5           # 避免被 yfinance 限制

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
    logger.info(f"成功讀取 {len(df)} 筆決策紀錄，日期範圍 {df['日期'].min()} ~ {df['日期'].max()}")
    return df

# ============================================================================
# 2. 取得歷史股價 (包含未來 N 個交易日)
# ============================================================================
def fetch_future_prices(symbol: str, start_date: datetime, days_ahead: int = 30) -> pd.DataFrame:
    """
    利用 yfinance 取得從 start_date 之後的股價，最多 days_ahead 個交易日。
    回傳 DataFrame 包含 Open, High, Low, Close, Volume。
    """
    # 因 yfinance 需要美股格式，但台股代號如 3037.TW 可直接使用
    ticker = yf.Ticker(symbol)
    # 抓取從 start_date 到 (start_date + days_ahead*3) 的區間（避免假日太少）
    end_date = start_date + timedelta(days=days_ahead*3)
    hist = ticker.history(start=start_date.strftime("%Y-%m-%d"), end=end_date.strftime("%Y-%m-%d"))
    if hist.empty:
        logger.warning(f"無法取得 {symbol} 從 {start_date.date()} 之後的資料")
        return pd.DataFrame()
    # 排除決策日當天（只取未來）
    hist = hist[hist.index.date > start_date.date()]
    return hist

# ============================================================================
# 3. 分析單一決策的績效
# ============================================================================
def analyze_decision(row: pd.Series) -> Dict:
    """針對一筆決策，計算多個時間框架的績效指標"""
    symbol = row["代號"]
    decision_date = row["日期"]
    entry_price = row["收盤價"]
    # 抓取未來至少 20 個交易日的價格
    hist = fetch_future_prices(symbol, decision_date, days_ahead=max(N_DAYS_LIST)+2)
    if hist.empty:
        return None

    result = {
        "日期": decision_date,
        "代號": symbol,
        "名稱": row["名稱"],
        "收盤價": entry_price,
        "RSI": row["RSI"],
        "戰場預判": row["戰場預判"],
        "籌碼訊號": row["籌碼訊號"],
        "行動指令": row["行動指令"],
    }

    # 計算每個 N_DAYS 的指標
    for n in N_DAYS_LIST:
        if len(hist) >= n:
            prices = hist["Close"].iloc[:n]
            high = prices.max()
            low = prices.min()
            final_price = prices.iloc[-1]
            max_gain = (high - entry_price) / entry_price
            max_loss = (low - entry_price) / entry_price
            final_return = (final_price - entry_price) / entry_price
            # 是否觸及停損 (低於停損價)
            hit_stop = low < entry_price * STOP_LOSS_PCT
        else:
            max_gain = max_loss = final_return = np.nan
            hit_stop = False

        result[f"{n}D_最高漲幅%"] = max_gain * 100 if not np.isnan(max_gain) else None
        result[f"{n}D_最大跌幅%"] = max_loss * 100 if not np.isnan(max_loss) else None
        result[f"{n}D_最終漲幅%"] = final_return * 100 if not np.isnan(final_return) else None
        result[f"{n}D_觸及停損"] = hit_stop

    return result

# ============================================================================
# 4. 主流程
# ============================================================================
def main():
    logger.info("開始 NOC 決策回顧引擎...")
    df_decisions = load_decision_log()

    # 過濾：只分析有「行動指令」且不為「持股觀望」的決策（或你想全部分析）
    # 這裡我們只分析有明確買進或試單意圖的指令
    interesting_actions = ["建倉", "試單", "波段", "佈局", "長線鎖籌", "加碼", "扣款", "獲利巡航", "浮虧防禦", "洗盤耐受", "戰術撤離"]
    df_target = df_decisions[df_decisions["行動指令"].str.contains("|".join(interesting_actions), na=False)]
    logger.info(f"共篩選出 {len(df_target)} 筆具操作意義的決策")

    if len(df_target) == 0:
        logger.warning("沒有符合條件的決策，結束")
        return

    results = []
    for idx, row in df_target.iterrows():
        logger.info(f"分析 {row['代號']} 於 {row['日期'].date()}")
        res = analyze_decision(row)
        if res:
            results.append(res)
        time.sleep(SLEEP_BETWEEN_REQUESTS)  # 禮貌性延遲

    if not results:
        logger.warning("無任何決策可分析，可能缺乏未來價格數據")
        return

    df_results = pd.DataFrame(results)
    # 儲存詳細結果
    df_results.to_csv(REVIEW_CSV, index=False, encoding="utf-8-sig")
    logger.info(f"已儲存詳細回顧結果至 {REVIEW_CSV}")

    # ============ 產生統計摘要 ============
    print("\\n" + "="*60)
    print("NOC 戰情室決策績效統計摘要")
    print("="*60)

    # 依「戰場預判」分類統計
    categories = df_results["戰場預判"].unique()
    for cat in categories:
        if pd.isna(cat):
            continue
        subset = df_results[df_results["戰場預判"] == cat]
        print(f"\\n📊 訊號類別：{cat} (共 {len(subset)} 筆)")
        for n in N_DAYS_LIST:
            col_final = f"{n}D_最終漲幅%"
            col_win = f"{n}D_最終漲幅%"
            win_count = (subset[col_final] > 0).sum() if col_final in subset else 0
            avg_return = subset[col_final].mean() if col_final in subset else np.nan
            hit_stop_pct = subset[f"{n}D_觸及停損"].mean() * 100 if f"{n}D_觸及停損" in subset else 0
            print(f"   {n}日後：勝率 {win_count/len(subset)*100:.1f}% | 平均報酬 {avg_return:.2f}% | 停損觸及率 {hit_stop_pct:.1f}%")

    # 總體統計
    print("\\n" + "─"*40)
    print("📈 總體績效（所有決策合併）")
    for n in N_DAYS_LIST:
        col_final = f"{n}D_最終漲幅%"
        if col_final in df_results:
            avg = df_results[col_final].mean()
            median = df_results[col_final].median()
            win_rate = (df_results[col_final] > 0).mean() * 100
            hit_stop = df_results[f"{n}D_觸及停損"].mean() * 100 if f"{n}D_觸及停損" in df_results else 0
            print(f"   {n}日後：平均 {avg:.2f}% | 中位數 {median:.2f}% | 勝率 {win_rate:.1f}% | 停損觸及率 {hit_stop:.1f}%")

    print("\\n✅ 回顧完成。詳細數據請查閱", REVIEW_CSV)

if __name__ == "__main__":
    main()
