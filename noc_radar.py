# =============================================================================
# NOC 游擊隊雷達 (noc_radar.py) v16.7
# 整合：初升段突破、起漲攻擊區、旱地拔蔥、狙擊金叉
# 保留過熱攔截與基本面濾網
# =============================================================================

import yfinance as yf
import datetime
import pandas as pd
import numpy as np
import os
import json
import time
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed, TimeoutError
from dotenv import load_dotenv

from noc_core import (
    NOCStrategy, NOCDatabase,
    assess_volume_turnover_signal,
    is_overheated,
    detect_initial_breakout,
    calculate_monster_breakout,
    calculate_sniper_signal,
    NOCChipMatrix
)

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)
logging.getLogger('yfinance').setLevel(logging.CRITICAL)

class RadarConfig:
    MAX_WORKERS : int = int(os.environ.get("MAX_WORKERS", "5"))
    TARGET_FILE : str = "radar_targets.json"
    SCAN_LIST : list = [
        "2330.TW", "2317.TW", "2454.TW", "2382.TW", "2308.TW", "3231.TW", "3037.TW", "2303.TW",
        "3008.TW", "3034.TW", "3711.TW", "2357.TW", "2395.TW", "2408.TW", "2353.TW", "2379.TW",
        "4938.TW", "2301.TW", "2345.TW", "2324.TW", "3661.TW", "6669.TW", "3714.TW", "2881.TW",
        "2882.TW", "2891.TW", "2886.TW", "2884.TW", "2892.TW", "2885.TW", "2880.TW", "2883.TW",
        "2887.TW", "5871.TW", "2890.TW", "5880.TW", "2002.TW", "1216.TW", "1301.TW", "1303.TW",
        "1326.TW", "2912.TW", "9904.TW", "2603.TW", "2609.TW", "2615.TW", "2207.TW", "1101.TW",
        "1102.TW", "2412.TW",
        "2356.TW", "3163.TWO", "5388.TW", "8299.TWO", "3260.TWO", "2377.TW", "2383.TW", "3017.TW",
        "2352.TW", "3443.TW", "3529.TWO", "3293.TWO", "6488.TWO", "8069.TWO", "6274.TWO", "6239.TW",
        "3044.TW", "2449.TW", "2344.TW", "2409.TW", "3481.TW", "6116.TW", "4958.TW", "6176.TW",
        "3532.TW", "2371.TW", "2404.TW", "3702.TW", "8046.TW", "5483.TWO", "3105.TWO", "5347.TWO",
        "6147.TWO", "6214.TW", "2313.TW", "2368.TW", "3013.TW", "3019.TW", "3042.TW", "3324.TWO",
        "3533.TW", "3583.TW", "3653.TW", "4966.TWO", "5269.TW", "6269.TW", "6415.TW", "6531.TW",
        "8016.TW", "8081.TW", "8150.TW", "3376.TW", "3035.TW", "3227.TWO", "3131.TWO", "2451.TW",
        "5469.TW", "3413.TW", "3450.TW", "4919.TW",
        "1513.TW", "1514.TW", "1519.TW", "1605.TW", "1477.TW", "2049.TW", "2610.TW", "2618.TW",
        "1536.TW", "1795.TW", "2231.TW", "8464.TW", "9910.TW", "9914.TW", "9921.TW", "9941.TW",
        "1319.TW", "1707.TW", "1722.TW", "1504.TW", "1590.TW", "1609.TW", "1717.TW", "1802.TW",
        "2006.TW", "2014.TW", "2027.TW", "2101.TW", "2105.TW", "2106.TW", "2201.TW", "2204.TW",
        "2801.TW", "2834.TW", "2838.TW", "2845.TW", "2889.TW", "6005.TW", "2812.TW", "3045.TW",
        "4904.TW", "1229.TW", "1402.TW", "2606.TW", "5522.TW", "1108.TW", "1210.TW", "1314.TW",
        "1476.TW", "2633.TW", "4104.TW", "4105.TWO", "4142.TW", "4736.TW", "6446.TW", "6472.TW",
        "6550.TW", "6579.TW", "8422.TW", "6806.TW"
    ]

cfg = RadarConfig()

def scan_stock_for_wave(symbol: str, strategy: NOCStrategy) -> dict:
    try:
        stock = yf.Ticker(symbol)
        hist = stock.history(period="8mo").dropna(subset=["Close"])
        if len(hist) < 60:
            return None

        hist['20MA'] = hist['Close'].rolling(20).mean()
        hist['60MA'] = hist['Close'].rolling(60).mean()
        hist['5VMA'] = hist['Volume'].rolling(5).mean()

        info = stock.info
        shares_out = info.get("sharesOutstanding") or info.get("impliedSharesOutstanding")
        if shares_out is None or pd.isna(shares_out):
            shares_out = 0
        hist['Turnover_Rate'] = ((hist['Volume'] / shares_out) * 100).fillna(1.5) if shares_out > 0 else 1.5
        hist['Volume_Ratio'] = (hist['Volume'] / hist['5VMA'].shift(1)).fillna(1.0)

        hist['Candle_Ratio'] = (hist['High'] - hist[['Open','Close']].max(axis=1)) / (hist['High'] - hist['Low'] + 1e-9)
        hist['Close_vs_High'] = hist['Close'] / hist['High']
        hist['Is_Red'] = hist['Close'] >= hist['Open']

        hist['High_60'] = hist['High'].rolling(60).max()
        hist['Low_60'] = hist['Low'].rolling(60).min()
        hist['Price_Position'] = (hist['Close'] - hist['Low_60']) / (hist['High_60'] - hist['Low_60'] + 1e-9)

        hist['Return_5D'] = hist['Close'].pct_change(5) * 100
        hist['Return_10D'] = hist['Close'].pct_change(10) * 100

        td = hist.iloc[-1]
        close = td['Close']
        ma20 = td['20MA']
        ma60 = td['60MA']
        vol_ratio = td['Volume_Ratio']
        turnover = td['Turnover_Rate']
        price_position = td['Price_Position']

        # 趨勢與基本面
        trend_score = strategy.get_trend_score(hist)
        if trend_score < 0:
            return None
        raw_id = symbol.replace(".TW", "").replace(".TWO", "")
        fund_health = strategy.get_fundamental_health(raw_id)
        if "衰退" in fund_health or "警報" in fund_health:
            return None

        # 過熱攔截
        overheated, over_reason = is_overheated(
            close=close, ma20=ma20, ma60=ma60,
            recent_5d_return=td.get('Return_5D', 0),
            recent_10d_return=td.get('Return_10D', 0),
            price_position=price_position, vol_ratio=vol_ratio
        )
        if overheated:
            logger.debug(f"🔥 [過熱攔截] {symbol}: {over_reason}")
            return None

        # 四象限
        quadrant_signal = assess_volume_turnover_signal(
            vol_ratio=vol_ratio, turnover=turnover, shares_out=shares_out,
            price_position=price_position, candle_ratio=td['Candle_Ratio'],
            is_red=td['Is_Red'], close_vs_high=td['Close_vs_High']
        )
        danger = ("🔴 主力出貨區", "⚠️ 量價背離陷阱", "🔴 爆量長上影 (假突破/出貨)", "⚠️ 黑K出量 (賣壓沉重)")
        if quadrant_signal in danger:
            return None

        # 核心攻擊信號
        initial_break, break_type, _ = detect_initial_breakout(hist, td, lookback=20)
        monster = calculate_monster_breakout(hist, td)
        sniper = calculate_sniper_signal(hist)

        is_valid = initial_break or monster or sniper or (quadrant_signal == "🟢 起漲攻擊區")
        if not is_valid:
            return None

        # 戰術描述
        if monster:
            tactics_desc = f"🔥 旱地拔蔥 (爆量長紅突破季線)"
        elif sniper:
            tactics_desc = f"🌟 狙擊金叉 (底部扭轉)"
        elif initial_break:
            tactics_desc = f"🔥 {break_type}"
        else:
            tactics_desc = f"🚀 中段加速 | {quadrant_signal}"

        # RSI, 乖離
        delta = hist["Close"].diff()
        rs = delta.clip(lower=0).ewm(com=13, adjust=False).mean() / (-delta.clip(upper=0)).ewm(com=13, adjust=False).mean().replace(0, 0.001)
        rsi = (100 - (100 / (1 + rs))).iloc[-1]
        bias_20 = ((close - ma20) / ma20) * 100 if ma20 else 0

        return {
            "symbol": symbol, "name": raw_id, "close": round(close, 2),
            "RSI": round(rsi, 2), "Bias20": round(bias_20, 2),
            "Volume_Ratio": round(vol_ratio, 2), "Turnover": round(turnover, 2),
            "Quadrant": quadrant_signal, "Signal": tactics_desc,
            "trello_tip": "系統雷達自動篩選，等待總司令確認建倉。"
        }
    except Exception as e:
        logger.debug(f"掃描 {symbol} 異常: {e}")
        return None

if __name__ == "__main__":
    logger.info("⚡ NOC 游擊隊雷達 v16.7 (整合初升段/旱地拔蔥/狙擊金叉) 啟動...")
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
                logger.info(f"🎯 火種: {r['symbol']} 收{r['close']} | {r['Signal']}")

    logger.info(f"掃描完成，耗時 {time.time()-start_time:.1f} 秒，共 {len(found)} 檔")
    radar_dict = {t["symbol"]: {"name": t["name"], "tactics": t["Signal"], "trello_tip": t["trello_tip"]} for t in found}
    with open(cfg.TARGET_FILE, "w", encoding="utf-8") as f:
        json.dump(radar_dict, f, ensure_ascii=False, indent=4)
    logger.info(f"✅ 火種已寫入 {cfg.TARGET_FILE}")
