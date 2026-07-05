# =============================================================================
# NOC 短線飆股搜尋器 (noc_momentum.py) v1.0
# 核心功能：
# 1. 每日掃描台股，找出具備短線飆漲潛力的標的
# 2. 篩選條件：量比≥1.5、漲幅≥3%、換手率達股本門檻、突破20MA或20日高點
# 3. 排除爆量長上影、過熱乖離>20%
# 4. 產出 JSON 供 stock_bot 整合追蹤
# =============================================================================

import yfinance as yf
import pandas as pd
import numpy as np
import os
import json
import logging
import time
import re
import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed, TimeoutError
from dotenv import load_dotenv

# 可選：從 noc_core 導入部分函數（若存在且穩定）
try:
    from noc_core import calculate_all_indicators, get_finmind_chip_data
    USE_NOC_CORE = True
except ImportError:
    USE_NOC_CORE = False
    logging.warning("noc_core 未找到，將使用內建指標計算")

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)
logging.getLogger('yfinance').setLevel(logging.CRITICAL)

# =============================================================================
# 掃描池設定（沿用 noc_radar 清單，可自行增減）
# =============================================================================
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
        "1513.TW", "1514.TW", "1519.TW", "1605.TW", "1504.TW", "1503.TW", "1515.TW", "3708.TW", 
        "1609.TW", "1608.TW", "1611.TW", "1612.TW", "1618.TW", "9958.TW", "3712.TW", "1520.TW",
        "6409.TW", "1582.TW", "1532.TW", "4536.TW", "8926.TW", "6869.TW", "1537.TW",
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

OUTPUT_FILE = "momentum_targets.json"
MAX_WORKERS = int(os.environ.get("MAX_WORKERS", "5"))

# =============================================================================
# 短線飆股篩選參數
# =============================================================================
VOL_RATIO_MIN = 1.5          # 量比最低門檻
PRICE_CHANGE_MIN = 3.0       # 漲幅最低百分比
BIAS_20_MAX = 20.0           # 20MA乖離最高百分比（避免過熱）
TURNOVER_LARGE = 1.0         # 大型股（股本≥30億股）換手率門檻
TURNOVER_MID = 2.5           # 中型股（10~30億股）換手率門檻
TURNOVER_SMALL = 5.0         # 小型股（<10億股）換手率門檻

# =============================================================================
# 單一股票掃描函數
# =============================================================================
def scan_momentum_stock(symbol: str) -> dict:
    try:
        stock = yf.Ticker(symbol)
        # 下載 8 個月歷史資料（確保足夠計算 20MA、60MA）
        hist = stock.history(period="8mo").dropna(subset=["Close", "Volume"])
        if len(hist) < 60:
            return None

        # 取得股本資訊（用於換手率）
        info = stock.info
        shares_out = info.get("sharesOutstanding") or info.get("impliedSharesOutstanding")
        if not shares_out or shares_out == 0:
            # 若無股本，跳過該標的
            return None

        # 計算技術指標
        hist['5VMA'] = hist['Volume'].rolling(5).mean()
        hist['20MA'] = hist['Close'].rolling(20).mean()
        hist['60MA'] = hist['Close'].rolling(60).mean()
        hist['Turnover'] = (hist['Volume'] / shares_out) * 100   # 換手率%
        # 上影線比例
        hist['Candle_Upper'] = hist['High'] - hist[['Open','Close']].max(axis=1)
        hist['Candle_Range'] = hist['High'] - hist['Low']
        hist['Upper_Shadow_Ratio'] = hist['Candle_Upper'] / hist['Candle_Range'].replace(0, 1e-9)

        # 今日與昨日資料
        curr = hist.iloc[-1]
        prev = hist.iloc[-2]
        curr_vol = curr['Volume']
        prev_vol_ma5 = hist['5VMA'].iloc[-2]

        if pd.isna(prev_vol_ma5) or prev_vol_ma5 == 0:
            return None

        vol_ratio = curr_vol / prev_vol_ma5
        price_change = (curr['Close'] - prev['Close']) / prev['Close'] * 100

        # ----- 條件 1：量比與漲幅 -----
        if vol_ratio < VOL_RATIO_MIN or price_change < PRICE_CHANGE_MIN:
            return None

        # ----- 條件 2：換手率門檻（依股本分級）-----
        if shares_out >= 3_000_000_000:
            turn_th = TURNOVER_LARGE
        elif shares_out >= 1_000_000_000:
            turn_th = TURNOVER_MID
        else:
            turn_th = TURNOVER_SMALL
        if curr['Turnover'] < turn_th:
            return None

        # ----- 條件 3：突破20MA 或 突破20日高點 -----
        # 20日高點（不含今日）
        high_20 = hist['High'].rolling(20).max().shift(1).iloc[-1]
        # 是否首次站上20MA（昨日收盤低於昨日20MA）
        break_20ma = (curr['Close'] > curr['20MA']) and (prev['Close'] <= hist['20MA'].iloc[-2])
        # 是否首次突破20日高點（昨日收盤低於前20日高點）
        prev_high_20 = hist['High'].rolling(20).max().shift(2).iloc[-1] if len(hist) >= 2 else high_20
        break_20high = (curr['Close'] > high_20) and (prev['Close'] <= prev_high_20)
        if not (break_20ma or break_20high):
            return None

        # ----- 條件 4：避免爆量長上影（上影線比例 > 0.5 且量比 > 2.0）-----
        if curr['Upper_Shadow_Ratio'] > 0.5 and vol_ratio > 2.0:
            return None

        # ----- 條件 5：20MA乖離不可過大（避免追高）-----
        bias_20 = (curr['Close'] - curr['20MA']) / curr['20MA'] * 100
        if bias_20 > BIAS_20_MAX:
            return None

        # ----- 通過所有篩選，回傳結果 -----
        raw_id = symbol.replace(".TW", "").replace(".TWO", "")
        return {
            "symbol": symbol,
            "name": raw_id,
            "close": round(curr['Close'], 2),
            "vol_ratio": round(vol_ratio, 1),
            "change_pct": round(price_change, 1),
            "turnover": round(curr['Turnover'], 2),
            "bias20": round(bias_20, 1),
            "break_type": "突破20MA" if break_20ma else "突破20日高點",
            "tactics": "⚡ 短線飆股候選 (量價突破+換手達標)",
            "trello_tip": f"漲幅{price_change:.1f}% 量比{vol_ratio:.1f} 換手{curr['Turnover']:.2f}% 乖離{bias_20:.1f}% | { '突破20MA' if break_20ma else '突破20日高點' }"
        }

    except Exception as e:
        logger.debug(f"掃描 {symbol} 異常: {e}")
        return None


# =============================================================================
# 主程式
# =============================================================================
if __name__ == "__main__":
    logger.info("⚡ NOC 短線飆股搜尋器 v1.0 啟動...")
    start_time = time.time()

    # 可選：檢查大盤是否極度空頭（簡易版，可選用）
    try:
        twii = yf.Ticker("^TWII").history(period="2mo")
        if not twii.empty:
            twii['20MA'] = twii['Close'].rolling(20).mean()
            if twii['Close'].iloc[-1] < twii['20MA'].iloc[-1]:
                logger.warning("大盤跌破月線，空頭風險升高，短線搜尋結果僅供參考")
    except:
        pass

    found_targets = []
    logger.info(f"📡 開始掃描 {len(SCAN_LIST)} 檔標的，篩選短線飆股...")

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        future_to_symbol = {executor.submit(scan_momentum_stock, sym): sym for sym in SCAN_LIST}
        try:
            for future in as_completed(future_to_symbol, timeout=300):
                sym = future_to_symbol[future]
                try:
                    result = future.result()
                    if result:
                        found_targets.append(result)
                        logger.info(f"🎯 發現短線飆股: {sym} | 漲幅 +{result['change_pct']}% | 量比 {result['vol_ratio']} | {result['break_type']}")
                except Exception:
                    pass
        except TimeoutError:
            logger.error("掃描超時，強制中止")
        finally:
            executor.shutdown(wait=False, cancel_futures=True)

    elapsed = time.time() - start_time
    logger.info(f"掃描完成，耗時 {elapsed:.1f} 秒，共篩選出 {len(found_targets)} 檔")

    # 寫入 JSON
    output_dict = {t["symbol"]: {
        "name": t["name"],
        "tactics": t["tactics"],
        "trello_tip": t["trello_tip"]
    } for t in found_targets}

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output_dict, f, ensure_ascii=False, indent=4)

    logger.info(f"✅ 結果已寫入 {OUTPUT_FILE}，可匯入 stock_bot 進行後續追蹤。")
