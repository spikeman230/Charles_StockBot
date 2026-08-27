# =============================================================================
# NOC 游擊隊雷達 (noc_radar.py) v18.1
# 整合：初升段突破、起漲攻擊區、旱地拔蔥、狙擊金叉、ABCX回踩
# 新增：市場廣度聯動（量價背離時提高門檻並過濾弱籌碼）
# 採用 Hybrid Data Fetcher (SQLite 歷史底庫 + yfinance 當日即時拼接)
# 獨立掃描清單：從 stock_scan_list.py 載入 SCAN_LIST
# 紅燈模式：無視紅燈，強制掃描（最後強制輸出空清單）
# =============================================================================

import yfinance as yf
import datetime
import pandas as pd
import numpy as np
import os
import json
import time
import logging
import re
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed, TimeoutError
from dotenv import load_dotenv
from typing import Optional, Dict, Any, Tuple, List, Union

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
logging.getLogger('yfinance').setLevel(logging.CRITICAL)

# ===== 現在才載入獨立掃描清單（logger 已可用） =====
try:
    from stock_scan_list import SCAN_LIST
except ImportError:
    logger.error("❌ 找不到 stock_scan_list.py，請確保該檔案存在於同目錄下。")
    SCAN_LIST = [] # 空清單避免崩潰

from noc_core import (
    NOCStrategy, NOCDatabase, NOCDataFetcher,
    assess_volume_turnover_signal,
    is_overheated,
    detect_initial_breakout,
    calculate_monster_breakout,
    calculate_sniper_signal,
    NOCChipMatrix,
    calculate_all_indicators,
    detect_abcx_pullback,
    analyze_volume_price_pattern
)

# 環境變數（與 stock_bot 共用）
FINMIND_TOKEN = os.getenv("FINMIND_TOKEN")

class RadarConfig:
    MAX_WORKERS : int = int(os.environ.get("MAX_WORKERS", "5"))
    TARGET_FILE : str = "radar_targets.json"

cfg = RadarConfig()

# ---------- 輔助函數：與 stock_bot 完全相同的數據獲取（含法人籌碼） ----------
def get_finmind_chip_data(symbol: str, start_date_str: str) -> pd.DataFrame:
    if not FINMIND_TOKEN:
        return pd.DataFrame()
    match = re.search(r"\d+", symbol)
    if not match:
        return pd.DataFrame()
    try:
        url = "https://api.finmindtrade.com/api/v4/data"
        params = {
            "dataset": "TaiwanStockInstitutionalInvestorsBuySell",
            "data_id": match.group(),
            "start_date": start_date_str,
            "token": FINMIND_TOKEN
        }
        r = requests.get(url, params=params, timeout=10)
        data = r.json()
        if data.get("msg") == "success" and data.get("data"):
            df = pd.DataFrame(data["data"])
            df["net_buy"] = df["buy"] - df["sell"]
            df["type"] = "Other"
            df.loc[df["name"].str.contains("外資"), "type"] = "Foreign_Inv"
            df.loc[df["name"].str.contains("投信"), "type"] = "Trust_Inv"
            df.loc[df["name"].str.contains("自營商"), "type"] = "Dealer_Inv"
            pivot_df = df.groupby(["date", "type"])["net_buy"].sum().unstack(fill_value=0).reset_index()
            for col in ["Foreign_Inv", "Trust_Inv", "Dealer_Inv"]:
                if col not in pivot_df.columns:
                    pivot_df[col] = 0
            pivot_df["Date"] = pd.to_datetime(pivot_df["date"]).dt.date
            pivot_df.set_index("Date", inplace=True)
            return pivot_df[["Foreign_Inv", "Trust_Inv", "Dealer_Inv"]]
    except:
        pass
    return pd.DataFrame()

def calculate_chip_signals(hist: pd.DataFrame) -> pd.DataFrame:
    hist["Chip_Status"] = "➖ 中性/偏空"
    hist["Trust_Streak"] = 0
    if not {"Foreign_Inv", "Trust_Inv", "Dealer_Inv"}.issubset(hist.columns):
        return hist
    hist["Total_Institutional"] = hist["Foreign_Inv"] + hist["Trust_Inv"] + hist["Dealer_Inv"]
    hist["Signal_CoBuy"] = (hist["Foreign_Inv"] > 0) & (hist["Trust_Inv"] > 0)
    hist["Signal_Trust_Trend"] = ((hist["Trust_Inv"] > 0).astype(int).rolling(5).sum() >= 4) & (hist["Trust_Inv"] > 0)
    trust_dir = np.sign(hist["Trust_Inv"])
    hist["Trust_Streak"] = trust_dir.groupby((trust_dir != trust_dir.shift()).cumsum()).cumsum()
    conds = [hist["Signal_CoBuy"], hist["Signal_Trust_Trend"], hist["Total_Institutional"] > 0]
    hist["Chip_Status"] = np.select(conds, ["🤝 土洋齊買", "🏦 投信作帳", "📈 法人偏多"], default="➖ 中性/偏空")
    return hist

# =============================================================================
# Hybrid Data Fetcher (SQLite 歷史底庫 + yfinance 當日即時拼接)
# =============================================================================
def get_hybrid_stock_data(symbol: str, db: NOCDatabase) -> Optional[pd.DataFrame]:
    """
    混合數據獲取：
    1. 從本地 SQLite 讀取歷史 K 線 (200天)
    2. 從 yfinance 獲取最新 2 天即時數據
    3. 按日期合併／覆蓋，確保最新
    4. 補充股本並計算全部技術指標
    """
    try:
        # ---- 1. 讀取本地歷史數據 ----
        hist_local = db.get_stock_dataframe(symbol, days=200)
        if hist_local is not None and len(hist_local) >= 60:
            # 本地數據充足，僅獲取最新 2 天即時數據
            try:
                ticker = yf.Ticker(symbol)
                hist_live = ticker.history(period="2d")
                if hist_live.empty:
                    # yfinance 無新數據，直接使用本地
                    logger.debug(f"⚠️ {symbol} yfinance 無即時數據，使用本地底庫。")
                    hist = hist_local.copy()
                else:
                    # 合併：以本地為基礎，用即時數據覆蓋/追加
                    hist_local = hist_local.tz_localize(None) if hist_local.index.tz is not None else hist_local
                    hist_live = hist_live.tz_localize(None) if hist_live.index.tz is not None else hist_live
                    # 過濾掉本地已有日期的即時數據（避免重複）
                    hist_live_new = hist_live[~hist_live.index.isin(hist_local.index)]
                    # 若有相同日期，則以即時數據覆蓋（保留最新）
                    hist_combined = hist_local.copy()
                    for idx in hist_live.index:
                        if idx in hist_combined.index:
                            hist_combined.loc[idx] = hist_live.loc[idx]
                        else:
                            hist_combined = pd.concat([hist_combined, hist_live.loc[[idx]]])
                    # 排序
                    hist = hist_combined.sort_index()
            except Exception as e:
                logger.warning(f"⚠️ {symbol} yfinance 即時資料獲取失敗 ({e})，使用本地底庫。")
                hist = hist_local.copy()
        else:
            # ---- 2. 本地數據不足，下載完整 8 個月數據 ----
            logger.debug(f"⏳ {symbol} 本地數據不足 ({len(hist_local) if hist_local is not None else 0} 天)，下載完整歷史...")
            ticker = yf.Ticker(symbol)
            hist = ticker.history(period="8mo").dropna(subset=["Close"])
            if hist.empty:
                logger.debug(f"❌ {symbol} yfinance 無任何數據")
                return None
            # 存入資料庫（非同步，但此處直接存）
            try:
                start_date = (datetime.datetime.now() - datetime.timedelta(days=240)).strftime("%Y-%m-%d")
                fetcher = NOCDataFetcher(token=FINMIND_TOKEN)
                fetcher.fetch_and_store_stock_data(symbol, start_date, db)
            except Exception as e:
                logger.warning(f"⚠️ {symbol} 存入資料庫失敗: {e}")

        # ---- 3. 數據有效性檢查 ----
        if hist is None or len(hist) < 60:
            logger.debug(f"❌ {symbol} 數據不足 60 天，無法分析")
            return None

        # ---- 4. 補齊股本資訊 ----
        shares_out = db.get_shares_out(symbol)
        if shares_out <= 0:
            try:
                ticker = yf.Ticker(symbol)
                info = ticker.info
                shares_out = info.get("sharesOutstanding") or info.get("impliedSharesOutstanding")
                if shares_out:
                    db.save_shares_out(symbol, shares_out)
                else:
                    shares_out = np.nan
            except:
                shares_out = np.nan
        hist['Shares_Out'] = shares_out if shares_out else np.nan

        # ---- 5. 法人籌碼合併 ----
        hist['Date_Key'] = hist.index.date
        if FINMIND_TOKEN and (".TW" in symbol or ".TWO" in symbol):
            chip_df = get_finmind_chip_data(symbol, (datetime.datetime.now() - datetime.timedelta(days=200)).strftime("%Y-%m-%d"))
            if not chip_df.empty:
                hist = hist.merge(chip_df, left_on="Date_Key", right_index=True, how="left").ffill().fillna(0)

        # ---- 6. 計算全套技術指標 ----
        hist = calculate_all_indicators(hist, symbol=symbol, token=FINMIND_TOKEN)
        hist = calculate_chip_signals(hist)

        # ---- 7. 計算狙擊金叉與旱地拔蔥 ----
        sniper_val = calculate_sniper_signal(hist)
        hist['Sniper_Signal'] = sniper_val
        td_temp = hist.iloc[-1]
        monster_val = calculate_monster_breakout(hist, td_temp)
        hist['Monster_Breakout'] = monster_val

        return hist
    except Exception as e:
        logger.debug(f"❌ 獲取 {symbol} 數據異常: {e}")
        return None
# ---------- 雷達掃描函數（整合市場廣度聯動） ----------
def scan_stock_for_wave(symbol: str, strategy: NOCStrategy, db: NOCDatabase,
                        breadth_status: str, divergence_ratio: float) -> dict:
    """
    掃描單一股票，回傳火種資訊。
    若市場處於「量價背離」，則提高量比門檻，並要求投信買超（Trust_Streak > 0）。
    """
    try:
        hist = get_hybrid_stock_data(symbol, db)
        if hist is None:
            return None

        td = hist.iloc[-1]
        close = td['Close']
        ma20 = td['20MA']
        ma60 = td['60MA']
        vol_ratio = td['Volume_Ratio']
        turnover = td['Turnover_Rate']
        price_position = td['Price_Position'] if not pd.isna(td['Price_Position']) else 0.5
        trust_streak = td.get('Trust_Streak', 0)

        # 趨勢與基本面
        trend_score = strategy.get_trend_score(hist)
        if trend_score < 0:
            return None
        raw_id = symbol.replace(".TW", "").replace(".TWO", "")
        fund_health = strategy.get_fundamental_health(raw_id)
        if "衰退" in fund_health or "警報" in fund_health:
            return None

        # 過熱攔截（傳入 gap_pct）
        gap_pct = td.get('Gap_Pct', 0.0)
        overheated, over_reason = is_overheated(
            close=close, ma20=ma20, ma60=ma60,
            recent_5d_return=td.get('Return_5D', 0),
            recent_10d_return=td.get('Return_10D', 0),
            price_position=price_position, vol_ratio=vol_ratio,
            gap_pct=gap_pct
        )
        if overheated:
            logger.debug(f"🔥 [過熱攔截] {symbol}: {over_reason}")
            return None

        # 四象限信號
        quadrant_signal = assess_volume_turnover_signal(
            vol_ratio=vol_ratio,
            turnover=turnover,
            shares_out=td.get('Shares_Out', 0),
            price_position=price_position,
            candle_ratio=td['Candle_Ratio'],
            is_red=td['Is_Red'],
            close_vs_high=td['Close_vs_High']
        )
        danger = ("🔴 主力出貨區", "⚠️ 量價背離陷阱", "🔴 爆量長上影 (假突破/出貨)", "⚠️ 黑K出量 (賣壓沉重)")
        if quadrant_signal in danger:
            return None

        # ========== 核心攻擊信號（含 ABCX 回踩） ==========
        initial_break, break_type, _ = detect_initial_breakout(hist, td, lookback=20)
        monster = td.get('Monster_Breakout', False)
        sniper = td.get('Sniper_Signal', False)

        # ABCX 回踩判定（需同時滿足站穩月季線）
        abcx = detect_abcx_pullback(hist, td)
        abcx_valid = abcx and (close > ma20) and (close > ma60)

        # ===== 市場廣度動態過濾 =====
        is_divergence = (breadth_status == "量價背離" or divergence_ratio >= 60.0)
        is_healthy = (breadth_status == "健康放量" or divergence_ratio <= 35.0)

        # 定義動態門檻
        if is_divergence:
            # 嚴格防誘多模式：量比提高 15%~20%，且要求投信買超
            min_vol_ratio = 1.5 # 原 1.3 提高約 15%
            min_turn_threshold = 1.2 # 換手率可略提高，但此處主要檢查量比
            require_trust_positive = True
            market_tip = "🔴 大盤量價背離(誘多盤)，嚴格限制作戰規模，禁止追高！"
        elif is_healthy:
            min_vol_ratio = 1.3 # 正常門檻
            require_trust_positive = False
            market_tip = "🟢 大盤健康放量順風，符合波段攻擊試單條件。"
        else:
            min_vol_ratio = 1.3
            require_trust_positive = False
            market_tip = "➖ 市場動能持平，正常篩選。"

        # 對 initial_break 進行額外過濾（若背離，則要求更嚴格的量比與投信買超）
        if initial_break:
            if is_divergence:
                # 重新檢查 good_volume 條件：我們可以在外部檢查 vol_ratio >= min_vol_ratio 且 trust_streak > 0
                # 由於 detect_initial_breakout 已經有 good_volume（基於 1.3 倍），若背離且 vol_ratio 不足，我們視為無效
                if vol_ratio < min_vol_ratio:
                    initial_break = False
                if require_trust_positive and trust_streak <= 0:
                    initial_break = False

        # 對 abcx_valid 進行額外過濾
        if abcx_valid and is_divergence:
            # 背離時要求更高的量比（表示回踩時仍有基本量能）且投信買超
            if vol_ratio < 1.2: # 額外檢查
                abcx_valid = False
            if require_trust_positive and trust_streak <= 0:
                abcx_valid = False

        # 對 monster 和 sniper 可考慮也加入過濾，但需求未指定，暫不處理

        # 任何一項成立即為有效火種（但已受動態過濾影響）
        is_valid = initial_break or monster or sniper or (quadrant_signal == "🟢 起漲攻擊區") or abcx_valid
        if not is_valid:
            return None

        # 戰術描述（優先級：旱地拔蔥 > 狙擊金叉 > ABCX > 初升段 > 起漲攻擊區）
        if monster:
            tactics_desc = f"🔥 旱地拔蔥 (爆量長紅突破季線)"
        elif sniper:
            tactics_desc = f"🌟 狙擊金叉 (底部扭轉)"
        elif abcx_valid:
            tactics_desc = f"🌀 ABCX回踩 (量縮不破且穩守月季線)"
        elif initial_break:
            tactics_desc = f"🔥 {break_type}"
        else:
            tactics_desc = f"🚀 中段加速 | {quadrant_signal}"

        # ===== 台股生存法則：計算量價口訣 =====
        vp_pattern = analyze_volume_price_pattern(hist, td)
        if vp_pattern != "➖ 價量結構平穩":
            extra_tip = f"量價: {vp_pattern}"
        else:
            extra_tip = ""

        # 計算 RSI 與乖離（可選，用於輸出）
        delta = hist["Close"].diff()
        rs = delta.clip(lower=0).ewm(com=13, adjust=False).mean() / (-delta.clip(upper=0)).ewm(com=13, adjust=False).mean().replace(0, 0.001)
        rsi = (100 - (100 / (1 + rs))).iloc[-1]
        bias_20 = ((close - ma20) / ma20) * 100 if ma20 else 0

        # 組裝 trello_tip，包含市場警語與量價口訣
        base_tip = f"系統雷達自動篩選，等待總司令確認建倉。{market_tip}"
        if extra_tip:
            full_tip = f"{base_tip} ({extra_tip})"
        else:
            full_tip = base_tip

        return {
            "symbol": symbol,
            "name": raw_id,
            "close": round(close, 2),
            "RSI": round(rsi, 2),
            "Bias20": round(bias_20, 2),
            "Volume_Ratio": round(vol_ratio, 2),
            "Turnover": round(turnover, 2),
            "Quadrant": quadrant_signal,
            "Signal": tactics_desc,
            "trello_tip": full_tip
        }
    except Exception as e:
        logger.debug(f"掃描 {symbol} 異常: {e}")
        return None
# ---------- 主程式 ----------
if __name__ == "__main__":
    logger.info("⚡ NOC 游擊隊雷達 v18.1 (市場廣度聯動版) 啟動...")
    start_time = time.time()

    # ---- 解析 SCAN_LIST ----
    if isinstance(SCAN_LIST, dict):
        symbols = list(SCAN_LIST.keys())
        logger.info(f"📋 載入 {len(symbols)} 檔標的 (來自 dict)")
    elif isinstance(SCAN_LIST, list):
        symbols = SCAN_LIST
        logger.info(f"📋 載入 {len(symbols)} 檔標的 (來自 list)")
    else:
        logger.error("❌ SCAN_LIST 格式錯誤，須為 dict 或 list")
        symbols = []

    if not symbols:
        logger.warning("⚠️ SCAN_LIST 為空，結束程式")
        exit(0)

    strategy = NOCStrategy()
    db = NOCDatabase()

    # 獲取大盤狀態（包含市場廣度）
    macro = strategy.get_macro_status()
    # 提取市場廣度指標（若無則使用預設）
    breadth_status = macro.get("breadth_status", "動能持平")
    divergence_ratio = macro.get("divergence_ratio", 0.0)
    breadth_summary = macro.get("market_breadth", "➖ 無廣度數據")

    logger.info(f"📊 市場廣度: {breadth_summary}")

    # ===== 紅燈處理：記錄旗標，清空檔案，最後強制輸出空清單 =====
    is_red_light = False
    if macro["status"] == "🔴 紅燈":
        logger.warning("🚨🚨🚨 警告：目前大盤處於 🔴 紅燈（空頭極端危險期）！已依據總司令作戰協議放寬防線限制，雷達將無視紅燈，強制執行掃描以提前捕捉上升股！")
        with open(cfg.TARGET_FILE, "w", encoding="utf-8") as f:
            json.dump({}, f)
        is_red_light = True
        # 不 exit，繼續往下執行掃描

    logger.info(f"📡 大盤{macro['status']}，開始掃描 {len(symbols)} 檔")
    found = []
    with ThreadPoolExecutor(max_workers=cfg.MAX_WORKERS) as ex:
        # 將廣度狀態傳入每個掃描任務
        futures = {
            ex.submit(scan_stock_for_wave, sym, strategy, db, breadth_status, divergence_ratio): sym
            for sym in symbols
        }
        for future in as_completed(futures, timeout=300):
            try:
                r = future.result()
                if r:
                    found.append(r)
                    logger.info(f"🎯 火種: {r['symbol']} 收{r['close']} | {r['Signal']} | {r['trello_tip']}")
            except Exception as e:
                logger.error(f"❌ 掃描 {futures[future]} 時發生例外: {e}")

    logger.info(f"掃描完成，耗時 {time.time()-start_time:.1f} 秒，共 {len(found)} 檔")

    # ===== 寫入 JSON，若紅燈則強制寫入空字典 =====
    if is_red_light:
        radar_dict = {} # 強制清空
        logger.info("🔴 紅燈模式：強制輸出空火種清單，確保無舊資料殘留。")
    else:
        radar_dict = {
            t["symbol"]: {
                "name": t["name"],
                "tactics": t["Signal"],
                "trello_tip": t["trello_tip"]
            } for t in found
        }

    with open(cfg.TARGET_FILE, "w", encoding="utf-8") as f:
        json.dump(radar_dict, f, ensure_ascii=False, indent=4)

    if not is_red_light:
        logger.info(f"✅ 火種已寫入 {cfg.TARGET_FILE}")
