# =============================================================================
# NOC 戰情室核心引擎 v18.1 (ABCX 3.0 真主力洗盤伏擊 + 市場廣度價量背離判讀)
# 功能：籌碼矩陣、四象限量價、精準買點引擎（ABCX 3.0 + 起漲第一棒）
# 嚴格過熱過濾、階梯式動態停損/停利、大盤海象判讀 (市場廣度背離防禦)
# 本地 SQLite 資料庫支援
# =============================================================================

import yfinance as yf
import pandas as pd
import numpy as np
import threading
import json
import logging
import math
import datetime
import sqlite3
import os
import re
import requests
from typing import Dict, Any, Optional, Tuple, List

# 靜音防護
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)
logging.getLogger('yfinance').setLevel(logging.CRITICAL)

# =============================================================================
# 1. 籌碼矩陣判定引擎（基本版）
# =============================================================================
def analyze_chip_tactics(turnover: float, volume_ratio: float, market_mode: str = "BEAR") -> str:
    t_val = turnover * 1.3 if market_mode == "BULL" else turnover
    v_val = volume_ratio * 1.3 if market_mode == "BULL" else volume_ratio

    if t_val > 10.0 and v_val > 5.0:
        return "🚀【極速發動】換手與量能極致爆發！主力籌碼全軍突擊，低檔可佈局，高檔提防動能竭盡！"
    elif 5.0 <= t_val <= 10.0 and 3.0 <= v_val <= 5.0:
        return "🔥【加速起漲】法人大單瘋狂鎖籌！多頭動能確認，波段安全加碼點。"
    elif t_val > 5.0 and 2.0 <= v_val < 3.0:
        return "✅【啟動訊號】主力實彈溫和點火換手！籌碼結構洗淨，波段最佳右側核心底倉建倉點。"
    elif t_val > 5.0 and v_val < 2.0:
        return "⚠️【陷阱警報】換手率極高但量比完全衰退！量價嚴重背離，假突破真倒貨，指令：撤離！"
    else:
        return "➖ 籌碼動態平穩"

# =============================================================================
# 2. 進階籌碼矩陣 (NOCChipMatrix)
# =============================================================================
class NOCChipMatrix:
    def analyze(self, df: pd.DataFrame, market_mode: str = "BEAR") -> str:
        try:
            latest = df.iloc[-1]
            volume_ratio = latest.get('Volume_Ratio')
            turnover_rate = latest.get('Turnover_Rate', 0.0)
            shares_out = latest.get('Shares_Out', 0.0)

            if volume_ratio is None or pd.isna(volume_ratio):
                vma5_yesterday = df['Volume'].rolling(5).mean().shift(1).iloc[-1]
                volume_ratio = latest['Volume'] / vma5_yesterday if vma5_yesterday else 0.0

            if pd.isna(shares_out) or shares_out == 0:
                turnover_threshold = 3.0
            elif shares_out >= 3_000_000_000:
                turnover_threshold = 1.0
            elif shares_out >= 1_000_000_000:
                turnover_threshold = 2.5
            else:
                turnover_threshold = 5.0

            if market_mode == "BULL":
                volume_threshold = 1.3
                high_lookback = 10
            else:
                volume_threshold = 1.5
                high_lookback = 20

            recent_high = df['High'].rolling(high_lookback).max().iloc[-1]

            if (volume_ratio >= volume_threshold) and (latest['High'] >= recent_high) and (turnover_rate >= turnover_threshold):
                return "🔥 主力點火 (籌碼突破)"
            else:
                return "➖ 無點火訊號"
        except Exception as e:
            logger.error(f"籌碼矩陣分析異常: {e}")
            return "⚠️ 籌碼分析失敗"

# =============================================================================
# 3. 量價四象限戰術矩陣 (含K線形態防禦)
# =============================================================================
def assess_volume_turnover_signal(vol_ratio: float, turnover: float, shares_out: float,
                                  price_position: float, candle_ratio: float = 0.0,
                                  is_red: bool = True, close_vs_high: float = 1.0) -> str:
    if shares_out >= 3_000_000_000:
        threshold = 1.0
    elif shares_out >= 1_000_000_000:
        threshold = 2.5
    else:
        threshold = 5.0

    if vol_ratio >= 1.5 and turnover >= threshold:
        if (close_vs_high < 0.96 and not is_red) or candle_ratio > 0.5:
            return "🔴 爆量長上影 (假突破/出貨)"
        return "🟢 起漲攻擊區"

    if vol_ratio >= 2.0 and turnover >= threshold * 1.6 and price_position > 0.8:
        return "🔴 主力出貨區"

    if vol_ratio >= 1.8 and turnover < threshold * 0.5:
        return "⚠️ 量價背離陷阱"

    if vol_ratio < 0.8 and turnover < threshold:
        return "➖ 量縮低換手 (洗盤/人氣退潮)"

    if not is_red and vol_ratio > 1.2 and turnover < threshold * 1.2:
        return "⚠️ 黑K出量 (賣壓沉重)"

    return "➖ 中性觀望"

# =============================================================================
# 4. 過熱攔截函數
# =============================================================================
def is_overheated(close: float, ma20: float, ma60: float,
                  recent_5d_return: float, recent_10d_return: float,
                  price_position: float, vol_ratio: float, gap_pct: float = 0.0) -> Tuple[bool, str]:
    reasons = []
    if ma20 > 0:
        bias20 = (close - ma20) / ma20 * 100
        if bias20 > 30:
            reasons.append(f"20MA乖離{bias20:.1f}%")
    if ma60 > 0:
        bias60 = (close - ma60) / ma60 * 100
        if bias60 > 50:
            reasons.append(f"60MA乖離{bias60:.1f}%")
    if recent_5d_return > 30:
        reasons.append(f"5日漲幅{recent_5d_return:.1f}%")
    if recent_10d_return > 50:
        reasons.append(f"10日漲幅{recent_10d_return:.1f}%")
    if price_position > 0.9 and vol_ratio > 2.5:
        reasons.append(f"高檔爆量(位置{price_position:.2f},量比{vol_ratio:.1f})")
    if gap_pct > 4.0 and price_position > 0.8:
        reasons.append(f"高檔竭盡缺口(跳空{gap_pct:.1f}%)")
    if reasons:
        return True, " | ".join(reasons)
    return False, ""

def is_entry_overheated(td: pd.Series) -> Tuple[bool, str]:
    reasons = []
    bias_20ma = td.get('Bias_20MA', 0.0)
    bias_5ma = td.get('Bias_5MA', 0.0)
    gap_pct = td.get('Gap_Pct', 0.0)
    upper_shadow = td.get('Upper_Shadow_Ratio', 0.0)

    if bias_20ma > 7.5:
        reasons.append(f"20MA乖離{bias_20ma:.1f}% > 7.5%")
    if bias_5ma > 4.5:
        reasons.append(f"5MA乖離{bias_5ma:.1f}% > 4.5%")
    if gap_pct > 1.8:
        reasons.append(f"跳空{gap_pct:.1f}% > 1.8%")
    if upper_shadow > 0.40:
        reasons.append(f"上影線比例{upper_shadow:.2f} > 0.40")

    if reasons:
        return True, " | ".join(reasons)
    return False, ""

# =============================================================================
# 5. 初升段突破偵測
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
    shares_out = td.get('Shares_Out', 0)
    if shares_out >= 3_000_000_000:
        turn_th = 1.0
    elif shares_out >= 1_000_000_000:
        turn_th = 1.5
    else:
        turn_th = 3.0
    good_volume = vol_ratio >= 1.3 and turnover >= turn_th

    bias = (close - ma20) / ma20 * 100 if ma20 > 0 else 0
    if bias > 20:
        return False, "", 0

    gap_pct = td.get('Gap_Pct', 0.0)
    gap_str = f" [帶突破缺口 {gap_pct:.1f}%]" if gap_pct > 1.0 else ""

    if (first_above_ma20 or first_break_high) and good_volume:
        if first_break_high:
            return True, f"🚀 首次突破20日高點{gap_str}", 3
        else:
            return True, f"🔥 放量站上20MA{gap_str}", 2
    return False, "", 0

# =============================================================================
# 5-1. 精準買點判斷引擎 (ABCX 3.0 真主力洗盤伏擊 + 起漲第一棒)
# =============================================================================
def detect_precision_buy_point(hist: pd.DataFrame, td: pd.Series) -> Tuple[bool, str, float]:
    close = td['Close']
    ma20 = td.get('20MA', 0)
    ma60 = td.get('60MA', 0)
    if pd.isna(ma20) or pd.isna(ma60):
        return False, "", 0.0
    if close <= ma20 or close <= ma60:
        return False, "", 0.0

    overheated, reason = is_entry_overheated(td)
    if overheated:
        logger.debug(f"買點過熱過濾: {reason}")
        return False, "", 0.0

    vol_ratio = td.get('Volume_Ratio', 1.0)
    bias_20ma = td.get('Bias_20MA', 0.0)

    # ---------- 戰術A：ABCX 3.0 真主力洗盤伏擊 ----------
    lookback_start = max(0, len(hist) - 10)
    lookback_end = max(0, len(hist) - 3)
    recent_hist = hist.iloc[lookback_start:lookback_end] if lookback_end > lookback_start else pd.DataFrame()
    if not recent_hist.empty:
        a_point = None
        for idx in recent_hist.index:
            row = recent_hist.loc[idx]
            if row['Volume'] > row['5VMA'] * 1.5 and row['Close'] >= row['Open']:
                a_point = row
                break
        if a_point is not None:
            vma5_yest = hist['5VMA'].shift(1).iloc[-1]
            if not pd.isna(vma5_yest) and td['Volume'] < vma5_yest * 0.7:
                if 0 <= bias_20ma <= 3.5:
                    # 濾網1：右側止跌確認
                    is_red = td['Close'] >= td['Open']
                    low = td['Low']
                    high = td['High']
                    open_price = td['Open']
                    close_price = td['Close']
                    lower_shadow = (min(open_price, close_price) - low) / (high - low + 1e-6)
                    filter1_pass = (is_red or lower_shadow >= 0.25)

                    # 濾網2：主力籌碼未渙散
                    trust_streak = td.get('Trust_Streak', 0)
                    chip_status = td.get('Chip_Status', '')
                    is_bearish = "偏空" in chip_status or "空頭" in chip_status
                    filter2_pass = not (trust_streak < -1 or is_bearish)

                    # 濾網3：實質基本面/題材動能 (相容字串與數值)
                    yoy_raw = td.get('YoY', None)
                    yoy_num = None
                    if yoy_raw is not None:
                        if isinstance(yoy_raw, (int, float)):
                            yoy_num = yoy_raw
                        elif isinstance(yoy_raw, str):
                            cleaned = yoy_raw.strip()
                            if cleaned.upper() != "N/A" and cleaned != "":
                                match = re.search(r'([-+]?\d*\.?\d+)', cleaned)
                                if match:
                                    yoy_num = float(match.group(1))
                    
                    if yoy_num is not None and yoy_num <= 0:
                        filter3_pass = False
                    else:
                        filter3_pass = True

                    if filter1_pass and filter2_pass and filter3_pass:
                        stop_loss = max(a_point['Low'], ma20 * 0.985)
                        return True, "🌀 ABCX真主力洗盤 (籌碼鎖定+量縮止跌)", round(stop_loss, 2)

    # ---------- 戰術B：起漲第一棒 (旱地拔蔥) ----------
    if len(hist) >= 2:
        prev_close = hist['Close'].iloc[-2]
        prev_ma20 = hist['20MA'].iloc[-2]
        if prev_close < prev_ma20 and close > ma20 and td['Close'] > td['Open'] and vol_ratio >= 1.8 and bias_20ma <= 4.0:
            stop_loss = td['Open'] * 0.98
            return True, "🔥 起漲第一棒 (旱地拔蔥)", round(stop_loss, 2)

    return False, "", 0.0

def detect_abcx_pullback(hist: pd.DataFrame, td: pd.Series) -> bool:
    if len(hist) < 20:
        return False
    required_cols = ['5VMA', '20MA', '60MA']
    for col in required_cols:
        if col not in hist.columns:
            return False
    recent_hist = hist.iloc[-11:-1]
    breakout_days = recent_hist[
        (recent_hist['Volume'] > recent_hist['5VMA'] * 1.5) & 
        (recent_hist['Close'] >= recent_hist['Open'])
    ]
    if breakout_days.empty:
        return False
    actual_vol = td.get('Volume', 0)
    vma5_yest = hist['5VMA'].shift(1).iloc[-1]
    if vma5_yest <= 0:
        return False
    is_volume_shrunk = (actual_vol < vma5_yest * 0.7)
    close = td['Close']
    ma20 = td['20MA']
    ma60 = td['60MA']
    return bool(is_volume_shrunk and (close >= ma20) and (close >= ma60))

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
# 基本面輔助函數
# =============================================================================
def get_revenue_yoy(symbol: str, token: str = "") -> str:
    if not token:
        return "N/A"
    match = re.search(r"\d+", symbol)
    if not match:
        return "N/A"
    try:
        url = "https://api.finmindtrade.com/api/v4/data"
        params = {
            "dataset": "TaiwanStockMonthRevenue",
            "data_id": match.group(),
            "start_date": (datetime.datetime.now() - datetime.timedelta(days=400)).strftime("%Y-%m-%d"),
            "token": token
        }
        r = requests.get(url, params=params, timeout=10)
        data = r.json()
        if data.get("msg") == "success" and data.get("data"):
            df = pd.DataFrame(data["data"])
            latest = df.iloc[-1]
            prev = df[(df["revenue_year"] == latest["revenue_year"] - 1) & (df["revenue_month"] == latest["revenue_month"])]
            if not prev.empty and prev.iloc[-1]["revenue"] > 0:
                return str(round((latest["revenue"] - prev.iloc[-1]["revenue"]) / prev.iloc[-1]["revenue"] * 100, 2)) + "%"
    except:
        pass
    return "N/A"

def get_pe_ratio(symbol: str) -> str:
    try:
        ticker = yf.Ticker(symbol)
        info = ticker.info
        pe = info.get("trailingPE") or info.get("forwardPE")
        return str(round(pe, 2)) if pe else "N/A"
    except:
        return "N/A"

def get_finmind_chip_data(symbol: str, start_date_str: str, token: str = "") -> pd.DataFrame:
    if not token:
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
            "token": token
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
# 台股生存法則：量價結構判讀
# =============================================================================
def analyze_volume_price_pattern(hist: pd.DataFrame, td: pd.Series) -> str:
    if len(hist) < 2:
        return "➖ 價量結構平穩"
    
    ret = (td['Close'] - hist['Close'].iloc[-2]) / hist['Close'].iloc[-2] * 100
    vol_ratio_act = td.get('Volume_Ratio_Act', 1.0)
    bias20 = td.get('Bias_20MA', 0.0)
    
    is_spike = vol_ratio_act > 1.5
    is_shrink = vol_ratio_act < 0.7
    is_rise = ret > 0.5
    is_fall = ret < -0.5
    is_flat = -0.5 <= ret <= 0.5
    
    lookback = 10
    fast_up_slow_down = False
    fast_down_slow_up = False
    
    if len(hist) >= lookback + 1:
        recent = hist.iloc[-lookback-1:-1]
        recent_ret = (recent['Close'] - recent['Close'].shift(1)) / recent['Close'].shift(1) * 100
        up_ret = recent_ret[recent_ret > 0]
        down_ret = recent_ret[recent_ret < 0]
        up_days = len(up_ret)
        down_days = len(down_ret)
        avg_gain = up_ret.mean() if up_days > 0 else 0.0
        avg_loss = abs(down_ret.mean()) if down_days > 0 else 0.0
        
        if avg_gain > avg_loss * 1.5 and up_days < down_days:
            fast_up_slow_down = True
        if avg_loss > avg_gain * 1.5 and down_days < up_days:
            fast_down_slow_up = True
    
    if fast_up_slow_down and bias20 > 5:
        return "🔥 快漲慢跌 (主力出貨中)"
    if fast_down_slow_up and bias20 < -5:
        return "📥 快跌慢漲 (主力吸籌中)"
    if is_spike and is_rise:
        return "⚠️ 放量上漲 (提防短線回落)"
    if is_spike and is_fall:
        return "🔄 放量下跌 (恐慌盤，準備反彈)"
    if is_shrink and is_flat and bias20 > 10:
        return "🔴 縮量不漲/不跌 (高檔頭部已成)"
    if is_shrink and is_rise:
        return "📈 縮量上漲 (籌碼鎖定，繼續上漲)"
    if is_shrink and is_fall:
        return "📉 縮量下跌 (無量下殺，繼續下跌)"
    return "➖ 價量結構平穩"

# =============================================================================
# 7. 本地 SQLite 資料庫支援
# =============================================================================
class NOCDatabase:
    def __init__(self, db_path: str = "noc_warroom.db"):
        self.db_path = db_path
        self._lock = threading.Lock()
        self._init_tables()

    def _init_tables(self):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute('''
                CREATE TABLE IF NOT EXISTS market_health (
                    date TEXT PRIMARY KEY,
                    twii_close REAL,
                    twii_20ma REAL,
                    twii_60ma REAL,
                    foreign_futures_net INTEGER
                )
            ''')
            conn.execute('''
                CREATE TABLE IF NOT EXISTS stock_prices (
                    symbol TEXT,
                    date TEXT,
                    open REAL,
                    high REAL,
                    low REAL,
                    close REAL,
                    volume INTEGER,
                    adj_close REAL,
                    PRIMARY KEY (symbol, date)
                )
            ''')
            try:
                conn.execute("ALTER TABLE stock_prices ADD COLUMN adj_close REAL")
            except sqlite3.OperationalError:
                pass
            conn.execute('''
                CREATE TABLE IF NOT EXISTS stock_info (
                    symbol TEXT PRIMARY KEY,
                    shares_out REAL,
                    last_update TEXT
                )
            ''')
            conn.execute('''
                CREATE TABLE IF NOT EXISTS fundamental_state (
                    symbol TEXT PRIMARY KEY,
                    status TEXT,
                    entry REAL,
                    trailing_stop REAL,
                    last_fetch TEXT
                )
            ''')

    def load_state(self) -> dict:
        try:
            with sqlite3.connect(self.db_path) as conn:
                df = pd.read_sql_query("SELECT symbol, status, entry, trailing_stop, last_fetch FROM fundamental_state", conn)
                if df.empty:
                    return {}
                return {row['symbol']: {"status": row['status'], "entry": row['entry'], "trailing_stop": row['trailing_stop'], "last_fetch": row['last_fetch']} for _, row in df.iterrows()}
        except:
            return {}

    def save_state(self, data: dict) -> bool:
        try:
            with sqlite3.connect(self.db_path) as conn:
                for sym, info in data.items():
                    conn.execute('''
                        INSERT OR REPLACE INTO fundamental_state (symbol, status, entry, trailing_stop, last_fetch)
                        VALUES (?, ?, ?, ?, ?)
                    ''', (sym, info.get('status', 'NONE'), info.get('entry', 0.0), info.get('trailing_stop', 0.0), info.get('last_fetch', '')))
            return True
        except:
            return False

    def get_stock_dataframe(self, symbol: str, days: int = 200) -> Optional[pd.DataFrame]:
        try:
            with sqlite3.connect(self.db_path) as conn:
                df = pd.read_sql_query('''
                    SELECT date, open, high, low, close, volume
                    FROM stock_prices
                    WHERE symbol = ?
                    ORDER BY date DESC
                    LIMIT ?
                ''', conn, params=(symbol, days))
            if df.empty:
                return None
            df['date'] = pd.to_datetime(df['date'])
            df = df.sort_values('date')
            df.set_index('date', inplace=True)
            df.rename(columns={
                'open': 'Open', 'high': 'High', 'low': 'Low', 'close': 'Close', 'volume': 'Volume'
            }, inplace=True)
            return df
        except Exception as e:
            logger.error(f"從資料庫讀取 {symbol} 失敗: {e}")
            return None

    def get_shares_out(self, symbol: str) -> float:
        try:
            with sqlite3.connect(self.db_path) as conn:
                cur = conn.execute("SELECT shares_out FROM stock_info WHERE symbol = ?", (symbol,))
                row = cur.fetchone()
                return row[0] if row else 0.0
        except:
            return 0.0

    def save_shares_out(self, symbol: str, shares_out: float):
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.execute("INSERT OR REPLACE INTO stock_info (symbol, shares_out, last_update) VALUES (?, ?, ?)",
                             (symbol, shares_out, datetime.datetime.now().isoformat()))
        except:
            pass

# =============================================================================
# 8. 數據獲取器 (NOCDataFetcher)
# =============================================================================
class NOCDataFetcher:
    def __init__(self, token: str = ""):
        self.token = token
        self.logger = logging.getLogger(__name__)

    def fetch_market_health_data(self, start_date: str, db: NOCDatabase):
        try:
            twii = yf.Ticker("^TWII").history(start=start_date)
            if twii.empty:
                self.logger.warning("無法下載加權指數資料")
                return
            twii['20MA'] = twii['Close'].rolling(20).mean()
            twii['60MA'] = twii['Close'].rolling(60).mean()
            with sqlite3.connect(db.db_path) as conn:
                for idx, row in twii.iterrows():
                    date_str = idx.strftime("%Y-%m-%d")
                    conn.execute('''
                        INSERT OR REPLACE INTO market_health (date, twii_close, twii_20ma, twii_60ma, foreign_futures_net)
                        VALUES (?, ?, ?, ?, ?)
                    ''', (date_str, row['Close'], row['20MA'], row['60MA'], 0))
            self.logger.info(f"大盤歷史資料已更新至 {db.db_path}")
        except Exception as e:
            self.logger.error(f"大盤資料下載失敗: {e}")

    def fetch_and_store_stock_data(self, symbol: str, start_date: str, db: NOCDatabase):
        try:
            ticker = yf.Ticker(symbol)
            hist = ticker.history(start=start_date)
            if hist.empty:
                return

            info = ticker.info
            shares_out = info.get("sharesOutstanding") or info.get("impliedSharesOutstanding")
            if shares_out:
                db.save_shares_out(symbol, shares_out)

            with sqlite3.connect(db.db_path) as conn:
                for idx, row in hist.iterrows():
                    date_str = idx.strftime("%Y-%m-%d")
                    conn.execute('''
                        INSERT OR REPLACE INTO stock_prices (symbol, date, open, high, low, close, volume, adj_close)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (symbol, date_str, row['Open'], row['High'], row['Low'], row['Close'], int(row['Volume']), row['Close']))
        except Exception as e:
            logger.error(f"❌ {symbol} 儲存失敗: {e}")

    def fetch_financial_statements(self, symbol: str, db: NOCDatabase) -> None:
        pass

# =============================================================================
# 9. 策略與風險管理類別
# =============================================================================
class NOCStrategy:
    def __init__(self, db: Optional[NOCDatabase] = None):
        self.logger = logging.getLogger(__name__)
        self.db = db

    def analyze_market_breadth(self, days: int = 2) -> Dict[str, Any]:
        """從本地資料庫計算市場價量廣度 (價漲量縮背離佔比)"""
        default = {
            "divergence_ratio": 0.0,
            "breadth_status": "動能持平",
            "summary": "➖ 市場數據不足，動能持平",
            "total_up": 0,
            "divergence_count": 0
        }
        if self.db is None:
            return default

        try:
            conn = sqlite3.connect(self.db.db_path)
            query = """
                SELECT symbol, date, close, volume
                FROM stock_prices
                WHERE date IN (
                    SELECT DISTINCT date
                    FROM stock_prices
                    ORDER BY date DESC
                    LIMIT ?
                )
                ORDER BY date DESC
            """
            df = pd.read_sql_query(query, conn, params=(days,))
            conn.close()
            if df.empty or len(df['date'].unique()) < 2:
                return default

            df['date'] = pd.to_datetime(df['date']).dt.date
            dates = sorted(df['date'].unique(), reverse=True)
            latest = dates[0]
            prev = dates[1]

            latest_df = df[df['date'] == latest].set_index('symbol')
            prev_df = df[df['date'] == prev].set_index('symbol')

            common_symbols = latest_df.index.intersection(prev_df.index)
            if len(common_symbols) == 0:
                return default

            merged = pd.DataFrame({
                'close_latest': latest_df.loc[common_symbols, 'close'],
                'close_prev': prev_df.loc[common_symbols, 'close'],
                'vol_latest': latest_df.loc[common_symbols, 'volume'],
                'vol_prev': prev_df.loc[common_symbols, 'volume']
            })
            merged['price_up'] = merged['close_latest'] > merged['close_prev']
            merged['vol_up'] = merged['vol_latest'] > merged['vol_prev']

            total_up = merged['price_up'].sum()
            if total_up == 0:
                divergence_ratio = 0.0
                divergence_count = 0
            else:
                divergence = merged[merged['price_up'] & ~merged['vol_up']]
                divergence_count = len(divergence)
                divergence_ratio = (divergence_count / total_up) * 100

            total_vol_latest = merged['vol_latest'].sum()
            total_vol_prev = merged['vol_prev'].sum()
            vol_change = total_vol_latest - total_vol_prev

            if divergence_ratio >= 60.0:
                breadth_status = "量價背離"
                summary = f"⚠️ 量價背離 (價漲量縮佔比 {divergence_ratio:.1f}%，市場追價動能衰退，誘多風險高)"
            elif divergence_ratio <= 35.0 and vol_change > 0:
                breadth_status = "健康放量"
                summary = f"🔥 健康放量 (價漲量縮佔比 {divergence_ratio:.1f}%，多方結構扎實)"
            else:
                breadth_status = "動能持平"
                summary = f"➖ 動能持平 (價漲量縮佔比 {divergence_ratio:.1f}%)"

            return {
                "divergence_ratio": round(divergence_ratio, 1),
                "breadth_status": breadth_status,
                "summary": summary,
                "total_up": int(total_up),
                "divergence_count": int(divergence_count)
            }
        except Exception as e:
            self.logger.error(f"市場廣度計算失敗: {e}")
            return default

    def get_macro_status(self) -> dict:
        try:
            twii = yf.Ticker("^TWII").history(period="6mo")
            if twii.empty:
                return {"status": "🟡 黃燈", "desc": "無法取得台股加權指數資料，啟動震盪保護機制，請嚴控資金。"}
            twii['20MA'] = twii['Close'].rolling(20).mean()
            twii['60MA'] = twii['Close'].rolling(60).mean()
            td = twii.iloc[-1]
            y_td = twii.iloc[-2]
            twii_close = td['Close']

            above_20ma = td['Close'] > td['20MA']
            above_20ma_yest = y_td['Close'] > y_td['20MA']
            ma20_rising = td['20MA'] >= y_td['20MA']
            if above_20ma and above_20ma_yest and ma20_rising:
                status = "🟢 綠燈"
                desc = "大盤多頭格局順風..."
            elif td['Close'] < td['60MA']:
                status = "🔴 紅燈"
                desc = "大盤崩盤警告..."
            else:
                status = "🟡 黃燈"
                desc = "大盤進入高密度震盪洗盤期..."

            # 計算市場廣度與動態降級
            breadth_data = self.analyze_market_breadth(days=2)
            divergence_ratio = breadth_data.get("divergence_ratio", 0.0)
            breadth_status = breadth_data.get("breadth_status", "動能持平")
            breadth_summary = breadth_data.get("summary", "")

            if status == "🟢 綠燈" and breadth_status == "量價背離" and divergence_ratio >= 65.0:
                status = "🟡 黃燈"
                desc = f"⚠️ 量價背離防禦！原綠燈因市場廣度嚴重背離 (Divergence {divergence_ratio:.1f}%) 強制降級，建議減半倉位、嚴禁追高！"

            # 剔除台積電權重
            try:
                tsmc = yf.Ticker("2330.TW")
                tsmc_info = tsmc.info
                default_weight = float(os.getenv("TSMC_WEIGHT", "0.40"))
                weight = max(0.35, min(0.45, default_weight))
            except Exception as e:
                weight = 0.40
            index_without_tsmc = twii_close * (1 - weight)

            return {
                "status": status,
                "desc": desc,
                "twii_close": round(twii_close, 0),
                "tsmc_weight": round(weight * 100, 1),
                "index_without_tsmc": round(index_without_tsmc, 0),
                "market_breadth": breadth_summary,
                "divergence_ratio": divergence_ratio,
                "breadth_status": breadth_status
            }
        except Exception as e:
            self.logger.error(f"❌ 大盤風向儀運算異常: {e}")
            return {"status": "🟡 黃燈", "desc": "總體經濟風向引擎異常，強制啟動系統震盪保護機制。"}

    def get_trend_score(self, hist_df: pd.DataFrame, market_mode: str = "BEAR") -> float:
        if len(hist_df) < 60:
            return -1.0
        if market_mode == "BULL":
            hist_df['10MA'] = hist_df['Close'].rolling(10).mean()
            hist_df['20MA'] = hist_df['Close'].rolling(20).mean()
            cp = hist_df['Close'].iloc[-1]
            ma10 = hist_df['10MA'].iloc[-1]
            ma20 = hist_df['20MA'].iloc[-1]
            return 1.0 if (cp > ma10 and ma10 > ma20) else -1.0
        else:
            hist_df['20MA'] = hist_df['Close'].rolling(20).mean()
            hist_df['60MA'] = hist_df['Close'].rolling(60).mean()
            cp = hist_df['Close'].iloc[-1]
            ma20 = hist_df['20MA'].iloc[-1]
            ma60 = hist_df['60MA'].iloc[-1]
            if cp > ma20 and ma20 > ma60:
                return 1.0
            elif cp > ma20 and cp > ma60:
                return 0.5
            else:
                return -1.0

    def get_fundamental_health(self, symbol: str) -> str:
        try:
            clean = symbol.replace(".TW", "").replace(".TWO", "")
            info = yf.Ticker(f"{clean}.TW").info
            rev = info.get("revenueGrowth")
            if rev is not None:
                pct = rev * 100
                if rev < -0.15:
                    return f"❌ 【基本面衰退】營收 YoY 呈現嚴重衰退 ({pct:.2f}%)，不符合長線波段體質！"
                elif rev < 0:
                    return f"⚠️ 【營收谷底轉機】營收 YoY 偏弱 ({pct:.2f}%)，但允許技術面與籌碼面突圍！"
                else:
                    return f"✅ 【基本面優良】營收 YoY 成長 ({pct:.2f}%)，符合龍蝦養殖標準"
            else:
                return "⚠️ 【數據寬容】外部 API 暫無 YoY 數據，交由技術與籌碼面判定。"
        except Exception:
            return "✅【營收健康】符合波段持有條件"

    def check_defcon_1_status(self) -> bool:
        try:
            twii = yf.Ticker("^TWII").history(period="3mo")
            if twii.empty:
                return False
            twii['60MA'] = twii['Close'].rolling(60).mean()
            return twii['Close'].iloc[-1] < twii['60MA'].iloc[-1]
        except Exception as e:
            self.logger.error(f"❌ DEFCON 1 協議監測器異常: {e}")
            return False

class NOCRiskManager:
    def __init__(self, total_capital: float = 130000.0):
        self.total_capital = total_capital

    def calculate_atr(self, hist_df: pd.DataFrame, period: int = 14) -> float:
        if len(hist_df) < period + 1:
            return hist_df['Close'].iloc[-1] * 0.025
        hl = hist_df['High'] - hist_df['Low']
        hc = np.abs(hist_df['High'] - hist_df['Close'].shift())
        lc = np.abs(hist_df['Low'] - hist_df['Close'].shift())
        tr = pd.concat([hl, hc, lc], axis=1).max(axis=1)
        return tr.rolling(period).mean().iloc[-1]

    def get_position_and_defense(self, symbol: str, current_price: float, hist_df: pd.DataFrame = None,
                                 market_mode: str = "BEAR", is_yellow_light: bool = False) -> dict:
        try:
            if hist_df is None or hist_df.empty:
                hist_df = yf.Ticker(symbol).history(period="6mo")
            atr = self.calculate_atr(hist_df, 14)
            mult = 2.0 if is_yellow_light else (1.8 if market_mode == "BULL" else 3.0)
            stop = current_price - (atr * mult)
            if not hist_df.empty and len(hist_df) >= 20:
                ma20 = hist_df['Close'].rolling(20).mean().iloc[-1]
                if not pd.isna(ma20):
                    stop = min(stop, ma20)
            risk_per_share = current_price - stop
            if risk_per_share <= 0:
                risk_per_share = current_price * 0.10
            max_shares = math.floor((self.total_capital * 0.02) / risk_per_share)
            max_alloc = math.floor((self.total_capital * 0.15) / current_price)
            total = min(max_shares, max_alloc)
            core = total // 2
            tactical = total - core
            return {
                "current_price": round(current_price, 2),
                "defense_line": round(stop, 2),
                "core_shares": int(core),
                "tactical_shares": int(tactical),
                "total_shares": int(total),
                "risk_per_share": round(risk_per_share, 2)
            }
        except Exception as e:
            logger.error(f"❌ 執行部位與移動防禦精算時發生異常: {e}")
            fallback_stop = current_price * 0.90
            fallback_shares = math.floor((self.total_capital * 0.075) / current_price)
            return {
                "current_price": round(current_price, 2),
                "defense_line": round(fallback_stop, 2),
                "core_shares": int(fallback_shares),
                "tactical_shares": int(fallback_shares),
                "total_shares": int(fallback_shares * 2),
                "risk_per_share": round(current_price - fallback_stop, 2)
            }

# =============================================================================
# 9-1. 階梯式動態出場與防守引擎
# =============================================================================
def evaluate_exit_and_stop(td: pd.Series, entry_price: float, highest_price: float, stop_price: float) -> Tuple[str, str]:
    close = td['Close']
    vol_ratio = td.get('Volume_Ratio', 1.0)
    ma5 = td.get('5MA', 0)
    
    highest_profit_pct = (highest_price - entry_price) / entry_price * 100 if entry_price > 0 else 0

    if close < stop_price:
        return "STOP_LOSS", f"🩸【戰術撤離】跌破防守價 {stop_price:.2f}"

    if vol_ratio > 2.2 and (td['Close'] < td['Open']) and (td['Close'] - td['Open']) / td['Open'] * 100 < -3.0:
        return "DISTRIBUTION", "🔴【高檔倒貨撤離】爆量長黑，主力出貨跡象"

    if highest_profit_pct >= 15.0:
        if close < ma5:
            return "TRAILING_STOP", "📉【主升段獲利了結】跌破5MA，移動停利觸發"
        else:
            return "HOLD", "📈【波段巡航】獲利>15%，沿5MA持有"

    if highest_profit_pct >= 8.0:
        if close < entry_price:
            return "BREAK_EVEN", "🛡️【保本離場】回測跌破成本價，零虧損出場"
        else:
            return "HOLD", "🔒【保本防禦】獲利>8%，防守線提升至成本價"

    return "HOLD", "⏳【洗盤耐受中】持有觀察"

# =============================================================================
# 10. 完整的技術指標計算函數
# =============================================================================
def calculate_all_indicators(hist: pd.DataFrame, symbol: str = "", token: str = "") -> pd.DataFrame:
    if hist is None or hist.empty:
        return hist

    if 'Shares_Out' not in hist.columns:
        hist['Shares_Out'] = np.nan

    now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8)))
    market_open = now.replace(hour=9, minute=0, second=0, microsecond=0)
    market_close = now.replace(hour=13, minute=30, second=0, microsecond=0)
    total_trading_minutes = (market_close - market_open).total_seconds() / 60.0
    if market_open < now < market_close:
        elapsed_mins = max(1.0, (now - market_open).total_seconds() / 60.0)
        vol_mult = total_trading_minutes / elapsed_mins
    else:
        vol_mult = 1.0
    hist["Est_Volume"] = hist["Volume"].copy()
    if len(hist) > 0:
        hist.iloc[-1, hist.columns.get_loc("Est_Volume")] = int(hist["Volume"].iloc[-1] * vol_mult)

    hist["5MA"] = hist["Close"].rolling(5).mean()
    hist["10MA"] = hist["Close"].rolling(10).mean()
    hist["20MA"] = hist["Close"].rolling(20).mean()
    hist["25MA"] = hist["Close"].rolling(25).mean()
    hist["60MA"] = hist["Close"].rolling(60).mean()

    hist["5VMA"] = hist["Volume"].rolling(5).mean()
    hist["20VMA"] = hist["Volume"].rolling(20).mean()
    hist["60VMA"] = hist["Volume"].rolling(60).mean()

    hist["Turnover_Rate"] = ((hist["Est_Volume"] / hist["Shares_Out"]) * 100).fillna(1.5)

    hist["Volume_Ratio_Act"] = (hist["Volume"] / hist["5VMA"].shift(1)).fillna(1.0)
    hist["Volume_Ratio_Est"] = (hist["Est_Volume"] / hist["5VMA"].shift(1)).fillna(1.0)
    hist["Volume_Ratio"] = hist["Volume_Ratio_Est"]

    hist['Candle_Ratio'] = (hist['High'] - hist[['Open','Close']].max(axis=1)) / (hist['High'] - hist['Low'] + 1e-9)
    hist['Close_vs_High'] = hist['Close'] / hist['High']
    hist['Is_Red'] = hist['Close'] >= hist['Open']

    hist['Upper_Shadow_Ratio'] = (hist['High'] - hist[['Open', 'Close']].max(axis=1)) / (hist['High'] - hist['Low'] + 1e-5)

    hist['Bias_5MA'] = (hist['Close'] - hist['5MA']) / hist['5MA'] * 100
    hist['Bias_20MA'] = (hist['Close'] - hist['20MA']) / hist['20MA'] * 100
    hist['Bias_60MA'] = (hist['Close'] - hist['60MA']) / hist['60MA'] * 100
    hist['Return_5D'] = hist['Close'].pct_change(5) * 100
    hist['Return_10D'] = hist['Close'].pct_change(10) * 100

    hist['Gap_Pct'] = ((hist['Open'] - hist['Close'].shift(1)) / hist['Close'].shift(1)).fillna(0) * 100

    hist["25MA_Rising"] = hist["25MA"] > hist["25MA"].shift(1)
    hist["Is_Red_Candle"] = hist["Close"] > hist["Open"]
    hist["Lower_Shadow_Ratio"] = (hist[["Open", "Close"]].min(axis=1) - hist["Low"]) / (hist["High"] - hist["Low"]).replace(0, 0.001)

    hist["Signal_2560"] = (hist["25MA"] > hist["25MA"].shift(3)) & (hist["5VMA"] > hist["60VMA"]) & (hist["Low"] <= hist["25MA"] * 1.015) & (hist["Close"] >= hist["25MA"] * 0.985) & (hist["Est_Volume"] < hist["5VMA"])
    hist["High_60"] = hist["High"].rolling(window=60, min_periods=20).max()
    hist["Low_60"] = hist["Low"].rolling(window=60, min_periods=20).min()
    hist["Price_Position"] = (hist["Close"] - hist["Low_60"]) / (hist["High_60"] - hist["Low_60"]).replace(0, np.nan)

    l9 = hist["Low"].rolling(9).min()
    h9 = hist["High"].rolling(9).max()
    hist["K"] = ((hist["Close"] - l9) / (h9 - l9).replace(0, np.nan) * 100).ewm(com=2, adjust=False).mean()
    hist["D"] = hist["K"].ewm(com=2, adjust=False).mean()

    delta = hist["Close"].diff()
    rs = delta.clip(lower=0).ewm(com=13, adjust=False).mean() / (-delta.clip(upper=0)).ewm(com=13, adjust=False).mean().replace(0, np.nan)
    hist["RSI"] = (100 - (100 / (1 + rs))).fillna(50)

    tr = pd.concat([hist["High"] - hist["Low"], (hist["High"] - hist["Close"].shift(1)).abs(), (hist["Low"] - hist["Close"].shift(1)).abs()], axis=1).max(axis=1)
    hist["ATR"] = tr.rolling(14).mean()

    hist["MACD"] = hist["Close"].ewm(span=12, adjust=False).mean() - hist["Close"].ewm(span=26, adjust=False).mean()
    hist["MACD_Hist"] = hist["MACD"] - hist["MACD"].ewm(span=9, adjust=False).mean()
    hist["STD20"] = hist["Close"].rolling(20).std()
    hist["BB_Width"] = (4 * hist["STD20"]) / hist["20MA"].replace(0, np.nan)

    hist["Is_Bottoming"] = ((hist["Close"] < hist["5MA"]) & (hist["MACD_Hist"].shift(2) < hist["MACD_Hist"].shift(1)) & (hist["MACD_Hist"].shift(1) < hist["MACD_Hist"]) & (hist["MACD_Hist"] < 0)).astype(int)
    hist["Is_Breakout"] = ((hist["Close"].shift(1) < hist["5MA"].shift(1)) & (hist["Close"] > hist["5MA"]) & (hist["Est_Volume"] > hist["5VMA"] * 1.2))
    hist["Sniper_Signal"] = (hist["Is_Bottoming"].rolling(3).max().fillna(0).astype(bool) & hist["Is_Breakout"])

    just_crossed_60ma = (hist["Close"] > hist["60MA"]) & (hist["Close"].shift(1) <= hist["60MA"].shift(1))
    extreme_volume = hist["Volume_Ratio"] >= 3.0
    solid_green = (hist["Close"] >= hist["Close"].shift(1) * 1.04)
    hist["Monster_Breakout"] = (just_crossed_60ma & extreme_volume & solid_green)

    hist["20_High"] = hist["High"].rolling(20).max().shift(1)
    hist["Shadow_Ratio"] = (hist["High"] - hist[["Open", "Close"]].max(axis=1)) / (hist["High"] - hist["Low"]).replace(0, 0.001)

    if symbol:
        hist['PE'] = get_pe_ratio(symbol)
        hist['YoY'] = get_revenue_yoy(symbol, token)
    else:
        hist['PE'] = 'N/A'
        hist['YoY'] = 'N/A'

    return hist

# =============================================================================
# 11. 統一的 get_stock_data 函數
# =============================================================================
def get_stock_data(symbol: str, db: Optional[NOCDatabase] = None, name: str = "") -> Optional[pd.DataFrame]:
    if db is None:
        db = NOCDatabase()

    hist = db.get_stock_dataframe(symbol, days=200)
    if hist is None or hist.empty:
        try:
            stock = yf.Ticker(symbol)
            hist = stock.history(period="8mo").dropna(subset=["Close"])
            if len(hist) < 60:
                return None
            start_date = (datetime.datetime.now() - datetime.timedelta(days=240)).strftime("%Y-%m-%d")
            fetcher = NOCDataFetcher()
            fetcher.fetch_and_store_stock_data(symbol, start_date, db)
        except Exception as e:
            logger.error(f"❌ 標的 [{symbol}] 即時下載失敗: {e}")
            return None
    else:
        if len(hist) < 60:
            return None

    shares_out = db.get_shares_out(symbol)
    if shares_out > 0:
        hist['Shares_Out'] = shares_out
    else:
        try:
            ticker = yf.Ticker(symbol)
            info = ticker.info
            shares_out = info.get("sharesOutstanding") or info.get("impliedSharesOutstanding")
            if shares_out:
                hist['Shares_Out'] = shares_out
                db.save_shares_out(symbol, shares_out)
            else:
                hist['Shares_Out'] = np.nan
        except:
            hist['Shares_Out'] = np.nan

    hist = calculate_all_indicators(hist)
    token = os.getenv("FINMIND_TOKEN", "")
    if token and 'Foreign_Inv' not in hist.columns:
        try:
            start_date = (datetime.datetime.now() - datetime.timedelta(days=200)).strftime("%Y-%m-%d")
            chip_df = get_finmind_chip_data(symbol, start_date, token)
            if not chip_df.empty:
                hist = hist.merge(chip_df, left_index=True, right_index=True, how='left').ffill().fillna(0)
        except:
            pass
    hist = calculate_chip_signals(hist)

    if 'PE' not in hist.columns:
        hist['PE'] = get_pe_ratio(symbol)
    if 'YoY' not in hist.columns:
        hist['YoY'] = get_revenue_yoy(symbol, token)

    return hist

# =============================================================================
# 12. 輔助函數：從資料庫取得大盤狀態 (可選)
# =============================================================================
def get_macro_status_from_db(db: NOCDatabase) -> dict:
    try:
        with sqlite3.connect(db.db_path) as conn:
            df = pd.read_sql_query("SELECT date, twii_close, twii_20ma, twii_60ma FROM market_health ORDER BY date DESC LIMIT 1", conn)
            if df.empty:
                return {"status": "🟡 黃燈", "desc": "無大盤資料"}
            row = df.iloc[-1]
            if row['twii_close'] > row['twii_20ma']:
                return {"status": "🟢 綠燈", "desc": "多頭格局"}
            elif row['twii_close'] < row['twii_60ma']:
                return {"status": "🔴 紅燈", "desc": "空頭格局"}
            else:
                return {"status": "🟡 黃燈", "desc": "震盪盤整"}
    except:
        return {"status": "🟡 黃燈", "desc": "資料庫讀取失敗"}
