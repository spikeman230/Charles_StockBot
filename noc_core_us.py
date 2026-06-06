# =============================================================================
# NOC 美股戰情室核心引擎 v1.0
# 適用市場：美股 (NYSE/NASDAQ)
# 核心邏輯：承襲 v16.12 三層防禦架構，參數微調以符合美股特性
# =============================================================================

import yfinance as yf
import pandas as pd
import numpy as np
import datetime
import logging
from typing import Dict, Any, Optional, Tuple

# 靜音防護
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)
logging.getLogger('yfinance').setLevel(logging.CRITICAL)

# =============================================================================
# 美股專用參數（相較台股，門檻略降）
# =============================================================================
VOL_RATIO_THRESHOLD = 1.3 # 台股1.5
TURNOVER_THRESHOLD_LARGE = 1.0 # 大型股(市值>100億美元)
TURNOVER_THRESHOLD_MID = 2.0 # 中型股(市值20-100億美元)
TURNOVER_THRESHOLD_SMALL = 4.0 # 小型股(市值<20億美元)

OVERHEAT_5D_RETURN = 20 # 台股30
OVERHEAT_10D_RETURN = 35 # 台股50
OVERHEAT_BIAS20 = 20 # 台股30
OVERHEAT_BIAS60 = 35 # 台股50
OVERHEAT_PRICE_POSITION = 0.85 # 台股0.9
OVERHEAT_VOL_RATIO = 2.5 # 與台股相同

# =============================================================================
# 1. 量價四象限戰術矩陣 (美股版)
# =============================================================================
def assess_volume_turnover_signal(vol_ratio: float, turnover: float, market_cap: float,
                                  price_position: float, candle_ratio: float = 0.0,
                                  is_red: bool = True, close_vs_high: float = 1.0) -> str:
    """
    美股版量價四象限矩陣
    - market_cap: 市值（美元），用於判斷股本分級門檻
    """
    # 市值分級門檻
    if market_cap >= 100_000_000_000: # 1000億美元以上
        threshold = TURNOVER_THRESHOLD_LARGE
    elif market_cap >= 20_000_000_000: # 200億美元以上
        threshold = TURNOVER_THRESHOLD_MID
    else:
        threshold = TURNOVER_THRESHOLD_SMALL

    if vol_ratio >= VOL_RATIO_THRESHOLD and turnover >= threshold:
        if (close_vs_high < 0.96 and not is_red) or candle_ratio > 0.5:
            return "🔴 爆量長上影 (假突破/出貨)"
        return "🟢 起漲攻擊區"

    if vol_ratio >= 2.0 and turnover >= threshold * 1.6 and price_position > OVERHEAT_PRICE_POSITION:
        return "🔴 主力出貨區"

    if vol_ratio >= 1.8 and turnover < threshold * 0.5:
        return "⚠️ 量價背離陷阱"

    if vol_ratio < 0.8 and turnover < threshold:
        return "➖ 量縮低換手 (洗盤/人氣退潮)"

    if not is_red and vol_ratio > 1.2 and turnover < threshold * 1.2:
        return "⚠️ 黑K出量 (賣壓沉重)"

    return "➖ 中性觀望"

# =============================================================================
# 2. 過熱攔截函數 (美股版)
# =============================================================================
def is_overheated(close: float, ma20: float, ma60: float,
                  recent_5d_return: float, recent_10d_return: float,
                  price_position: float, vol_ratio: float) -> Tuple[bool, str]:
    reasons = []
    if ma20 > 0:
        bias20 = (close - ma20) / ma20 * 100
        if bias20 > OVERHEAT_BIAS20:
            reasons.append(f"20MA乖離{bias20:.1f}%")
    if ma60 > 0:
        bias60 = (close - ma60) / ma60 * 100
        if bias60 > OVERHEAT_BIAS60:
            reasons.append(f"60MA乖離{bias60:.1f}%")
    if recent_5d_return > OVERHEAT_5D_RETURN:
        reasons.append(f"5日漲幅{recent_5d_return:.1f}%")
    if recent_10d_return > OVERHEAT_10D_RETURN:
        reasons.append(f"10日漲幅{recent_10d_return:.1f}%")
    if price_position > OVERHEAT_PRICE_POSITION and vol_ratio > OVERHEAT_VOL_RATIO:
        reasons.append(f"高檔爆量(位置{price_position:.2f},量比{vol_ratio:.1f})")
    if reasons:
        return True, " | ".join(reasons)
    return False, ""

# =============================================================================
# 3. 初升段突破偵測 (美股版，邏輯相同，門檻沿用)
# =============================================================================
def detect_initial_breakout(hist: pd.DataFrame, td: pd.Series, lookback: int = 20) -> Tuple[bool, str, int]:
    close = td['Close']
    ma20 = td['20MA']
    if pd.isna(ma20):
        return False, "", 0

    hist_slice = hist.iloc[-lookback-1:-1]
    was_above_ma20 = (hist_slice['Close'] > hist_slice['20MA']).any()
    first_above_ma20 = (close > ma20) and not was_above_ma20

    high_20 = hist['High'].rolling(20).max().shift(1).iloc[-1]
    hist_high_20 = hist['High'].rolling(20).max().shift(1)
    was_break_high = (hist_slice['Close'] > hist_high_20.iloc[-lookback-1:-1]).any()
    first_break_high = (close > high_20) and not was_break_high

    vol_ratio = td.get('Volume_Ratio', 1.0)
    turnover = td.get('Turnover_Rate', 0.0)
    # 美股版市值門檻
    market_cap = td.get('Market_Cap', 0)
    if market_cap >= 100_000_000_000:
        turn_th = TURNOVER_THRESHOLD_LARGE
    elif market_cap >= 20_000_000_000:
        turn_th = TURNOVER_THRESHOLD_MID
    else:
        turn_th = TURNOVER_THRESHOLD_SMALL
    good_volume = vol_ratio >= VOL_RATIO_THRESHOLD and turnover >= turn_th

    bias = (close - ma20) / ma20 * 100 if ma20 > 0 else 0
    if bias > 20:
        return False, "", 0

    if (first_above_ma20 or first_break_high) and good_volume:
        if first_break_high:
            return True, "🚀 首次突破20日高點", 3
        else:
            return True, "🔥 放量站上20MA", 2
    return False, "", 0

# =============================================================================
# 4. 旱地拔蔥偵測 (邏輯相同，門檻沿用)
# =============================================================================
def calculate_monster_breakout(hist: pd.DataFrame, td: pd.Series) -> bool:
    close = td['Close']
    ma60 = td['60MA']
    if pd.isna(ma60):
        return False
    prev_close = hist['Close'].iloc[-2]
    prev_ma60 = hist['60MA'].iloc[-2] if len(hist) >= 2 else ma60
    just_crossed = (close > ma60) and (prev_close <= prev_ma60)
    if not just_crossed:
        return False
    vol_ratio = td.get('Volume_Ratio', 1.0)
    if vol_ratio < 3.0:
        return False
    solid_green = (close >= hist['Close'].iloc[-2] * 1.04)
    return solid_green

# =============================================================================
# 5. 狙擊金叉偵測 (邏輯相同)
# =============================================================================
def calculate_sniper_signal(hist: pd.DataFrame) -> bool:
    if len(hist) < 10:
        return False
    hist['MACD'] = hist['Close'].ewm(span=12, adjust=False).mean() - hist['Close'].ewm(span=26, adjust=False).mean()
    hist['MACD_Hist'] = hist['MACD'] - hist['MACD'].ewm(span=9, adjust=False).mean()
    hist['5MA'] = hist['Close'].rolling(5).mean()
    hist['Is_Bottoming'] = (
        (hist['Close'] < hist['5MA']) &
        (hist['MACD_Hist'].shift(2) < hist['MACD_Hist'].shift(1)) &
        (hist['MACD_Hist'].shift(1) < hist['MACD_Hist']) &
        (hist['MACD_Hist'] < 0)
    ).astype(int)
    hist['5VMA'] = hist['Volume'].rolling(5).mean()
    hist['Is_Breakout'] = (
        (hist['Close'].shift(1) < hist['5MA'].shift(1)) &
        (hist['Close'] > hist['5MA']) &
        (hist['Volume'] > hist['5VMA'] * 1.2)
    )
    bottom_3d = hist['Is_Bottoming'].rolling(3).max().fillna(0).astype(bool)
    sniper = bottom_3d.iloc[-1] and hist['Is_Breakout'].iloc[-1]
    return bool(sniper)

# =============================================================================
# 6. 高品質訊號三重確認濾網
# =============================================================================
def is_high_quality_signal(hist: pd.DataFrame, td: pd.Series, matrix_signal: str, market_mode: str) -> bool:
    recent_20_high = hist['High'].rolling(20).max().shift(1).iloc[-1]
    if pd.isna(recent_20_high):
        recent_20_high = hist['High'].iloc[-2]
    price_break = td['Close'] > recent_20_high
    vol_ratio = td.get('Volume_Ratio', 1.0)
    strong_volume = vol_ratio >= 2.0
    strong_chip = any(key in matrix_signal for key in ["極速發動", "加速起漲"])
    trend_score = td.get('Trend_Score', -1.0)
    good_trend = trend_score > 0
    return price_break and strong_volume and (strong_chip or good_trend)

# =============================================================================
# 7. 完整技術指標計算函數 (美股版)
# =============================================================================
def calculate_all_indicators(hist: pd.DataFrame, market_cap: float = 0) -> pd.DataFrame:
    """給定基礎 OHLCV 與市值，計算所有技術指標"""
    if hist is None or hist.empty:
        return hist

    # 動態量能預估（盤中即時修正，收盤後不影響）
    now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=-4))) # 美東時間
    market_open = now.replace(hour=9, minute=30, second=0, microsecond=0)
    market_close = now.replace(hour=16, minute=0, second=0, microsecond=0)
    total_trading_minutes = (market_close - market_open).total_seconds() / 60.0
    if market_open < now < market_close:
        elapsed_mins = max(1.0, (now - market_open).total_seconds() / 60.0)
        vol_mult = total_trading_minutes / elapsed_mins
    else:
        vol_mult = 1.0
    hist["Est_Volume"] = hist["Volume"].copy()
    if len(hist) > 0:
        hist.iloc[-1, hist.columns.get_loc("Est_Volume")] = int(hist["Volume"].iloc[-1] * vol_mult)

    # 均線
    hist["5MA"] = hist["Close"].rolling(5).mean()
    hist["20MA"] = hist["Close"].rolling(20).mean()
    hist["25MA"] = hist["Close"].rolling(25).mean()
    hist["60MA"] = hist["Close"].rolling(60).mean()
    hist["5VMA"] = hist["Est_Volume"].rolling(5).mean()
    hist["60VMA"] = hist["Volume"].rolling(60).mean()

    # 換手率（需市值，若無則用預設）
    if market_cap > 0:
        hist["Turnover_Rate"] = ((hist["Est_Volume"] * hist["Close"]) / market_cap) * 100
    else:
        hist["Turnover_Rate"] = 1.5
    hist["Volume_Ratio"] = (hist["Est_Volume"] / hist["5VMA"].shift(1)).fillna(1.0)

    # K線特徵
    hist['Candle_Ratio'] = (hist['High'] - hist[['Open','Close']].max(axis=1)) / (hist['High'] - hist['Low'] + 1e-9)
    hist['Close_vs_High'] = hist['Close'] / hist['High']
    hist['Is_Red'] = hist['Close'] >= hist['Open']

    # 乖離與漲幅
    hist['Bias_20MA'] = (hist['Close'] - hist['20MA']) / hist['20MA'] * 100
    hist['Bias_60MA'] = (hist['Close'] - hist['60MA']) / hist['60MA'] * 100
    hist['Return_5D'] = hist['Close'].pct_change(5) * 100
    hist['Return_10D'] = hist['Close'].pct_change(10) * 100

    # 價格位置 (60日區間)
    hist['High_60'] = hist['High'].rolling(60).max()
    hist['Low_60'] = hist['Low'].rolling(60).min()
    hist['Price_Position'] = (hist['Close'] - hist['Low_60']) / (hist['High_60'] - hist['Low_60'] + 1e-9)

    # KD
    l9 = hist["Low"].rolling(9).min()
    h9 = hist["High"].rolling(9).max()
    hist["K"] = ((hist["Close"] - l9) / (h9 - l9).replace(0, np.nan) * 100).ewm(com=2, adjust=False).mean()
    hist["D"] = hist["K"].ewm(com=2, adjust=False).mean()

    # RSI
    delta = hist["Close"].diff()
    rs = delta.clip(lower=0).ewm(com=13, adjust=False).mean() / (-delta.clip(upper=0)).ewm(com=13, adjust=False).mean().replace(0, np.nan)
    hist["RSI"] = (100 - (100 / (1 + rs))).fillna(50)

    # ATR
    tr = pd.concat([hist["High"] - hist["Low"], (hist["High"] - hist["Close"].shift(1)).abs(), (hist["Low"] - hist["Close"].shift(1)).abs()], axis=1).max(axis=1)
    hist["ATR"] = tr.rolling(14).mean()

    # MACD
    hist["MACD"] = hist["Close"].ewm(span=12, adjust=False).mean() - hist["Close"].ewm(span=26, adjust=False).mean()
    hist["MACD_Hist"] = hist["MACD"] - hist["MACD"].ewm(span=9, adjust=False).mean()
    hist["STD20"] = hist["Close"].rolling(20).std()
    hist["BB_Width"] = (4 * hist["STD20"]) / hist["20MA"].replace(0, np.nan)

    # 狙擊金叉
    hist["Is_Bottoming"] = ((hist["Close"] < hist["5MA"]) & (hist["MACD_Hist"].shift(2) < hist["MACD_Hist"].shift(1)) & (hist["MACD_Hist"].shift(1) < hist["MACD_Hist"]) & (hist["MACD_Hist"] < 0)).astype(int)
    hist["Is_Breakout"] = ((hist["Close"].shift(1) < hist["5MA"].shift(1)) & (hist["Close"] > hist["5MA"]) & (hist["Est_Volume"] > hist["5VMA"] * 1.2))
    hist["Sniper_Signal"] = (hist["Is_Bottoming"].rolling(3).max().fillna(0).astype(bool) & hist["Is_Breakout"])

    # 旱地拔蔥
    just_crossed_60ma = (hist["Close"] > hist["60MA"]) & (hist["Close"].shift(1) <= hist["60MA"].shift(1))
    extreme_volume = hist["Volume_Ratio"] >= 3.0
    solid_green = (hist["Close"] >= hist["Close"].shift(1) * 1.04)
    hist["Monster_Breakout"] = (just_crossed_60ma & extreme_volume & solid_green)

    # 其他
    hist["20_High"] = hist["High"].rolling(20).max().shift(1)
    hist["Shadow_Ratio"] = (hist["High"] - hist[["Open", "Close"]].max(axis=1)) / (hist["High"] - hist["Low"]).replace(0, 0.001)

    return hist

# =============================================================================
# 8. 統一數據獲取函數 (美股版，直接從 yfinance 讀取)
# =============================================================================
def get_stock_data(symbol: str) -> Optional[pd.DataFrame]:
    try:
        stock = yf.Ticker(symbol)
        hist = stock.history(period="8mo").dropna(subset=["Close"])
        if len(hist) < 60:
            return None
        # 取得市值（用於計算換手率）
        info = stock.info
        market_cap = info.get("marketCap", 0)
        # 計算技術指標（傳入市值）
        hist = calculate_all_indicators(hist, market_cap)
        hist['Market_Cap'] = market_cap
        # 補充基本面資訊
        hist['PE'] = info.get("trailingPE", "N/A")
        hist['YoY'] = (info.get("revenueGrowth", 0) * 100) if info.get("revenueGrowth") else "N/A"
        return hist
    except Exception as e:
        logger.error(f"❌ 美股標的 [{symbol}] 執行技術分析失敗: {e}")
        return None

# =============================================================================
# 9. 大盤風向儀 (美股版，使用 SPY 或 QQQ)
# =============================================================================
class NOCStrategy_US:
    def __init__(self):
        self.logger = logging.getLogger(__name__)

    def get_macro_status(self, index: str = "SPY") -> dict:
        try:
            etf = yf.Ticker(index)
            hist = etf.history(period="6mo")
            if hist.empty:
                return {"status": "🟡 黃燈", "desc": "無法取得美股大盤資料"}
            hist['20MA'] = hist['Close'].rolling(20).mean()
            hist['60MA'] = hist['Close'].rolling(60).mean()
            td = hist.iloc[-1]
            y_td = hist.iloc[-2]
            if td['Close'] > td['20MA'] and td['20MA'] >= y_td['20MA']:
                return {"status": "🟢 綠燈", "desc": "美股多頭格局 (SPY站上月線且向上)"}
            elif td['Close'] < td['60MA']:
                return {"status": "🔴 紅燈", "desc": "美股空頭警戒 (跌破季線)"}
            else:
                return {"status": "🟡 黃燈", "desc": "美股震盪洗盤期"}
        except Exception as e:
            self.logger.error(f"美股大盤判讀異常: {e}")
            return {"status": "🟡 黃燈", "desc": "大盤引擎異常"}

    def get_fundamental_health(self, symbol: str) -> str:
        """簡易基本面（僅供參考，因 yfinance 不一定有完整資料）"""
        try:
            info = yf.Ticker(symbol).info
            rev = info.get("revenueGrowth")
            if rev is not None:
                pct = rev * 100
                if rev < -0.15:
                    return f"❌ 營收衰退 ({pct:.2f}%)"
                elif rev < 0:
                    return f"⚠️ 營收偏弱 ({pct:.2f}%)"
                else:
                    return f"✅ 營收成長 ({pct:.2f}%)"
            else:
                return "⚠️ 無營收數據"
        except:
            return "✅ 數據寬容"
