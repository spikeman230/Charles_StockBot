# =============================================================================
# NOC 游擊隊雷達 (noc_radar.py) v17.2
# 整合：初升段突破、起漲攻擊區、旱地拔蔥、狙擊金叉、ABCX回踩
# 新增：台股生存法則（量價結構判讀）
# 採用與 stock_bot 完全相同的數據預處理（含動態量能、法人籌碼合併）
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
from typing import Optional, Dict, Any, Tuple

from noc_core import (
    NOCStrategy, NOCDatabase,
    assess_volume_turnover_signal,
    is_overheated,
    detect_initial_breakout,
    calculate_monster_breakout,
    calculate_sniper_signal,
    NOCChipMatrix,
    calculate_all_indicators,
    detect_abcx_pullback,
    analyze_volume_price_pattern # ✅ 新增：台股生存法則
)

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)
logging.getLogger('yfinance').setLevel(logging.CRITICAL)

# 環境變數（與 stock_bot 共用）
FINMIND_TOKEN = os.getenv("FINMIND_TOKEN")

class RadarConfig:
    MAX_WORKERS : int = int(os.environ.get("MAX_WORKERS", "5"))
    TARGET_FILE : str = "radar_targets.json"
    SCAN_LIST : list = [
        # 總司令專屬：200 檔旗艦級波段觀察池
        # [權值前 50 大]
        "0050.TW", "2330.TW", "2317.TW", "2454.TW", "2382.TW", "2308.TW", "3231.TW", "3037.TW",
        "2303.TW", "3008.TW", "3034.TW", "3711.TW", "2357.TW", "2395.TW", "2408.TW", "2353.TW",
        "2379.TW", "4938.TW", "2301.TW", "2345.TW", "2324.TW", "3661.TW", "6669.TW", "3714.TW",
        "2881.TW", "2882.TW", "2891.TW", "2886.TW", "2884.TW", "2892.TW", "2885.TW", "2880.TW",
        "2883.TW", "2887.TW", "5871.TW", "2890.TW", "5880.TW", "2002.TW", "1216.TW", "1301.TW",
        "1303.TW", "1326.TW", "2912.TW", "9904.TW", "2603.TW", "2609.TW", "2615.TW", "2207.TW",
        "1101.TW", "1102.TW",
        # [高動能科技 60 檔]
        "2356.TW", "3163.TWO", "5388.TW", "8299.TWO", "3260.TWO", "2377.TW", "2383.TW", "3017.TW",
        "2352.TW", "3443.TW", "3529.TWO", "3293.TWO", "6488.TWO", "8069.TWO", "6274.TWO", "6239.TW",
        "3044.TW", "2449.TW", "2344.TW", "2409.TW", "3481.TW", "6116.TW", "4958.TW", "6176.TW",
        "3532.TW", "2371.TW", "2404.TW", "3702.TW", "8046.TW", "5483.TWO", "3105.TWO", "5347.TWO",
        "6147.TWO", "6214.TW", "2313.TW", "2368.TW", "3013.TW", "3019.TW", "3042.TW", "3324.TWO",
        "3533.TW", "3583.TW", "3653.TW", "4966.TWO", "5269.TW", "6269.TW", "6415.TW", "6531.TW",
        "8016.TW", "8081.TW", "8150.TW", "3376.TW", "3035.TW", "3227.TWO", "3131.TWO", "2451.TW",
        "5469.TW", "3413.TW", "3450.TW", "4919.TW",
        # [區塊 3：重電/綠能/電纜與生技醫療 - 共 60 檔]
        # 重電綠能 (25檔)
        "1513.TW", "1514.TW", "1519.TW", "1605.TW", "1504.TW", "1503.TW", "1515.TW", "1520.TW",
        "3708.TW", "1609.TW", "1608.TW", "1611.TW", "1612.TW", "1618.TW", "9958.TW", "3712.TW",
        "6409.TW", "1582.TW", "1522.TW", "1532.TW", "4536.TW", "8926.TW", "6869.TW", "1537.TW",
        # [區塊 4：傳產塑化/汽車零組件/造船航太 - 共 30 檔]
        # 汽車零組件 (13檔)
        "1536.TW", "2231.TW", "1521.TW", "1525.TW", "2228.TW", "2115.TW", "2201.TW", "2204.TW",
        "3346.TW", "1339.TW", "6279.TW", "1524.TW", "1568.TW",
        # 傳產塑化化學 (12檔)
        "1314.TW", "1717.TW", "1304.TW", "1308.TW", "1309.TW", "1312.TW", "1305.TW", "1710.TW",
        "1704.TW", "4722.TW", "4739.TW", "1718.TW", "1319.TW", "6605.TW", "7736.TW", "1522.TW",
        # 造船與軍工航太 (5檔)
        "2208.TW", "2634.TW", "4541.TW", "8222.TW", "2646.TW"
    ]

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

def get_stock_data_for_radar(symbol: str) -> Optional[pd.DataFrame]:
    """
    與 stock_bot.get_stock_data 完全相同的實作，但無快取。
    使用 calculate_all_indicators 統一生成所有指標（含 Volume_Ratio_Act、真實 5VMA 等）
    """
    try:
        stock = yf.Ticker(symbol)
        info = stock.info
        shares_out = info.get("sharesOutstanding") or info.get("impliedSharesOutstanding")
        hist = stock.history(period="8mo").dropna(subset=["Close"])
        if len(hist) < 60:
            return None

        hist["Shares_Out"] = shares_out if shares_out else np.nan
        hist["Date_Key"] = hist.index.date

        # 合併法人籌碼
        if FINMIND_TOKEN and (".TW" in symbol or ".TWO" in symbol):
            chip_df = get_finmind_chip_data(symbol, (datetime.datetime.now() - datetime.timedelta(days=200)).strftime("%Y-%m-%d"))
            if not chip_df.empty:
                hist = hist.merge(chip_df, left_on="Date_Key", right_index=True, how="left").ffill().fillna(0)

        # ========== 使用核心引擎的統一指標計算 ==========
        hist = calculate_all_indicators(hist, symbol=symbol, token=FINMIND_TOKEN)

        # 補上籌碼信號
        hist = calculate_chip_signals(hist)

        # 計算狙擊金叉與旱地拔蔥（calculate_all_indicators 未包含）
        sniper_val = calculate_sniper_signal(hist)
        hist['Sniper_Signal'] = sniper_val

        td_temp = hist.iloc[-1]
        monster_val = calculate_monster_breakout(hist, td_temp)
        hist['Monster_Breakout'] = monster_val

        # 補充其他可能遺漏的指標（如 ATR、RSI 等，calculate_all_indicators 已包含 ATR）
        # 此處不再重複計算

        return hist
    except Exception as e:
        logger.debug(f"獲取 {symbol} 數據失敗: {e}")
        return None

# ---------- 雷達掃描函數 ----------
def scan_stock_for_wave(symbol: str, strategy: NOCStrategy) -> dict:
    try:
        hist = get_stock_data_for_radar(symbol)
        if hist is None:
            return None

        td = hist.iloc[-1]
        close = td['Close']
        ma20 = td['20MA']
        ma60 = td['60MA']
        vol_ratio = td['Volume_Ratio']
        turnover = td['Turnover_Rate']
        price_position = td['Price_Position'] if not pd.isna(td['Price_Position']) else 0.5

        # 趨勢與基本面（與 stock_bot 相同）
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

        # ✅ ABCX 回踩判定（需同時滿足站穩月季線）
        abcx = detect_abcx_pullback(hist, td)
        abcx_valid = abcx and (close > ma20) and (close > ma60)

        # 任何一項成立即為有效火種
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

        # 組裝 trello_tip（含量價口訣，若有）
        base_tip = "系統雷達自動篩選，等待總司令確認建倉。"
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
    logger.info("⚡ NOC 游擊隊雷達 v17.0 (含 ABCX 回踩 + 量價口訣) 啟動...")
    start_time = time.time()
    strategy = NOCStrategy()
    macro = strategy.get_macro_status()
    if macro["status"] == "🔴 紅燈":
        logger.warning("🚨 大盤跌破季線，停止掃描")
        with open(cfg.TARGET_FILE, "w", encoding="utf-8") as f:
            json.dump({}, f)
        exit(0)

    logger.info(f"📡 大盤{macro['status']}，開始掃描 {len(cfg.SCAN_LIST)} 檔")
    found = []
    with ThreadPoolExecutor(max_workers=cfg.MAX_WORKERS) as ex:
        futures = {ex.submit(scan_stock_for_wave, sym, strategy): sym for sym in cfg.SCAN_LIST}
        for future in as_completed(futures, timeout=300):
            r = future.result()
            if r:
                found.append(r)
                logger.info(f"🎯 火種: {r['symbol']} 收{r['close']} | {r['Signal']} | {r['trello_tip']}")

    logger.info(f"掃描完成，耗時 {time.time()-start_time:.1f} 秒，共 {len(found)} 檔")
    radar_dict = {t["symbol"]: {"name": t["name"], "tactics": t["Signal"], "trello_tip": t["trello_tip"]} for t in found}
    with open(cfg.TARGET_FILE, "w", encoding="utf-8") as f:
        json.dump(radar_dict, f, ensure_ascii=False, indent=4)
    logger.info(f"✅ 火種已寫入 {cfg.TARGET_FILE}")

