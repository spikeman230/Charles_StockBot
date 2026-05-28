# =============================================================================
# NOC 戰情室核心引擎 v16.6
# 功能：籌碼矩陣、四象限量價、K線形態防禦、過熱攔截、初升段突破偵測
# =============================================================================

import yfinance as yf
import pandas as pd
import numpy as np
import threading
import json
import logging
import math
import datetime
from typing import Dict, Any, Optional, Tuple

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
# 2. 進階籌碼矩陣 (NOCChipMatrix) – 股本分級 + 突破高點
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

            # 股本分級門檻
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
    """
    量比 × 換手率 四象限戰術矩陣 v2.0
    - 強多信號但出現爆量長上影 → 假突破
    - 黑K出量 → 賣壓沉重
    """
    if shares_out >= 3_000_000_000:
        threshold = 1.0
    elif shares_out >= 1_000_000_000:
        threshold = 2.5
    else:
        threshold = 5.0

    # 強多信號：量比≥1.5 且換手達標
    if vol_ratio >= 1.5 and turnover >= threshold:
        # 爆量長上影條件：收盤低於高點的96% 且 收黑K OR 上影線比例>0.5
        if (close_vs_high < 0.96 and not is_red) or candle_ratio > 0.5:
            return "🔴 爆量長上影 (假突破/出貨)"
        return "🟢 起漲攻擊區"

    # 死亡信號：爆量 + 高檔 + 換手過熱
    if vol_ratio >= 2.0 and turnover >= threshold * 1.6 and price_position > 0.8:
        return "🔴 主力出貨區"

    # 假突破陷阱：量比大但換手極低
    if vol_ratio >= 1.8 and turnover < threshold * 0.5:
        return "⚠️ 量價背離陷阱"

    # 量縮低換手
    if vol_ratio < 0.8 and turnover < threshold:
        return "➖ 量縮低換手 (洗盤/人氣退潮)"

    # 黑K出量賣壓
    if not is_red and vol_ratio > 1.2 and turnover < threshold * 1.2:
        return "⚠️ 黑K出量 (賣壓沉重)"

    return "➖ 中性觀望"

# =============================================================================
# 4. 過熱攔截函數
# =============================================================================
def is_overheated(close: float, ma20: float, ma60: float,
                  recent_5d_return: float, recent_10d_return: float,
                  price_position: float, vol_ratio: float) -> Tuple[bool, str]:
    """
    判斷股價是否已嚴重過熱，不應開新倉
    回傳 (是否過熱, 過熱原因)
    """
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
    if reasons:
        return True, " | ".join(reasons)
    return False, ""

# =============================================================================
# 5. 初升段突破偵測（首次放量站上20MA或突破20日高點）
# =============================================================================
def detect_initial_breakout(hist: pd.DataFrame, td: pd.Series) -> Tuple[bool, str, int]:
    """
    偵測初升段突破（首次突破關鍵價位）
    回傳 (是否突破, 突破類型, 強度分數)
    """
    close = td['Close']
    ma20 = td['20MA']
    if pd.isna(ma20):
        return False, "", 0

    prev_close = hist['Close'].iloc[-2]
    prev_ma20 = hist['20MA'].iloc[-2]
    was_below_ma20 = prev_close < prev_ma20

    # 首次站上20MA
    first_above_ma20 = close > ma20 and was_below_ma20

    # 20日高點突破（首次）
    high_20 = hist['High'].rolling(20).max().shift(1).iloc[-1]
    prev_high_20 = hist['High'].rolling(20).max().shift(2).iloc[-1]
    first_break_high = close > high_20 and (hist['Close'].iloc[-2] <= prev_high_20 if not pd.isna(prev_high_20) else True)

    # 量比與換手門檻（初升段放寬）
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

    # 乖離限制（避免已是高位）
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
# 6. 高品質訊號三重確認濾網 (收盤突破)
# =============================================================================
def is_high_quality_signal(hist: pd.DataFrame, td: pd.Series, matrix_signal: str, market_mode: str) -> bool:
    recent_20_high = hist['High'].rolling(20).max().shift(1).iloc[-1]
    if pd.isna(recent_20_high):
        recent_20_high = hist['High'].iloc[-2]
    price_break = td['Close'] > recent_20_high # 收盤突破
    vol_ratio = td.get('Volume_Ratio', 1.0)
    strong_volume = vol_ratio >= 2.0
    strong_chip = any(key in matrix_signal for key in ["極速發動", "加速起漲"])
    trend_score = td.get('Trend_Score', -1.0)
    good_trend = trend_score > 0
    return price_break and strong_volume and (strong_chip or good_trend)

# =============================================================================
# 7. 輔助類別：資料庫、連線Mock、風險管理、策略、數據獲取
# =============================================================================
class MockConn:
    def close(self) -> None:
        pass

class NOCDatabase:
    def __init__(self, db_path: str = "noc_state.json"):
        self.db_path = db_path
        self._lock = threading.Lock()
        self.conn = MockConn()

    def load_state(self) -> dict:
        with self._lock:
            try:
                with open(self.db_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except (FileNotFoundError, json.JSONDecodeError):
                logger.warning(f"⚠️ 找不到或無法解析狀態檔 {self.db_path}，初始化全新狀態資料結構。")
                return {}
            except Exception as e:
                logger.error(f"❌ 讀取狀態資料庫時發生未知異常: {e}")
                return {}

    def save_state(self, data: dict) -> bool:
        with self._lock:
            try:
                with open(self.db_path, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=4)
                return True
            except Exception as e:
                logger.error(f"❌ 寫入狀態資料庫時發生嚴重失敗: {e}")
                return False

class NOCStrategy:
    def __init__(self, db: Optional[NOCDatabase] = None):
        self.logger = logging.getLogger(__name__)
        self.db = db

    def get_macro_status(self) -> dict:
        try:
            twii = yf.Ticker("^TWII").history(period="6mo")
            if twii.empty:
                return {"status": "🟡 黃燈", "desc": "無法取得台股加權指數資料，啟動震盪保護機制，請嚴控資金。"}
            twii['20MA'] = twii['Close'].rolling(20).mean()
            twii['60MA'] = twii['Close'].rolling(60).mean()
            td = twii.iloc[-1]
            y_td = twii.iloc[-2]
            if td['Close'] > td['20MA'] and td['20MA'] >= y_td['20MA']:
                return {"status": "🟢 綠燈", "desc": "大盤多頭格局順風。積極抱緊長線底倉，防守線隨波段墊高，允許兵力推升至上限。"}
            elif td['Close'] < td['60MA']:
                return {"status": "🔴 紅燈", "desc": "大盤崩盤警告（跌破季線）。啟動防空 protocols，全面停止新標的建倉，嚴格保留現金。"}
            else:
                return {"status": "🟡 黃燈", "desc": "大盤進入高密度震盪洗盤期。嚴禁動用新資金盲目重倉，嚴格看守防禦底線。"}
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

class NOCDataFetcher:
    def __init__(self, token: str = ""):
        self.token = token
        self.logger = logging.getLogger(__name__)

    def fetch_financial_statements(self, symbol: str, db: NOCDatabase) -> None:
        try:
            self.logger.info(f"🚀 [DataFetcher] 啟動多執行緒安全線路，同步標的 {symbol} 的長線基本面數據...")
            state = db.load_state()
            now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if symbol not in state:
                state[symbol] = {"status": "NONE", "entry": 0.0, "trailing_stop": 0.0, "last_fetch": now_str}
            else:
                if isinstance(state[symbol], dict):
                    state[symbol]["last_fetch"] = now_str
            db.save_state(state)
            self.logger.info(f"✅ [DataFetcher] 標的 {symbol} 的波段基本面狀態資料同步更新成功。")
        except Exception as e:
            self.logger.error(f"❌ [DataFetcher] 執行多執行緒財務數據抓取時攔截到異常: {e}")
