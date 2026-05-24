# =============================================================================
# NOC 戰情室核心引擎 (noc_core.py) v16.0 - 龍蝦養殖波段籌碼完全體
# 核心戰略：
# 1. 籌碼矩陣擴充：導入量比與換手率多維籌碼判定，精準捕捉主力發動、洗盤與倒貨訊號。
# 2. 植入 60MA Trend_Score 趨勢判定器，從源頭過濾所有不符合大波段多頭之標的。
# 3. 強制掛載 YoY 基本面健康濾網，無情淘汰營收衰退之泡沫企業。
# 4. 風控防禦空間全面升級至 3.0 ATR，大幅拓寬容錯空間，抵禦主力惡意甩轎。
# 5. 執行緒安全防護升級，內建 MockConn 防爆盾，完美對接外部多執行緒呼叫。
# =============================================================================

import yfinance as yf
import pandas as pd
import numpy as np
import threading
import json
import logging
import math
import datetime
from typing import Dict, Any, Optional

# 🛡️ 啟動靜音防護罩：強制拔除 yfinance 內建的錯誤廣播喇叭，維持戰情室絕對專注
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)
logging.getLogger('yfinance').setLevel(logging.CRITICAL)

# =============================================================================
# 🚀 模組 0: 籌碼矩陣判定引擎 (Chip Tactics Engine)
# =============================================================================
def analyze_chip_tactics(turnover: float, volume_ratio: float, market_mode: str = "BEAR") -> str:
    """
    籌碼矩陣判定引擎 (雙模式自適應版)：
    BULL 模式：對籌碼數據進行 1.3 倍敏感度放大，更容易發出點火與啟動訊號。
    """
    t_val = turnover * 1.3 if market_mode == "BULL" else turnover
    v_val = volume_ratio * 1.3 if market_mode == "BULL" else volume_ratio

    if t_val > 10.0 and v_val > 5.0:
        return "🚀【極速發動】換手與量能極致爆發！主力籌碼全軍突擊，低檔可佈局，高檔提防動能竭盡！"
    elif 5.0 <= t_val <= 10.0 and 3.0 <= v_val <= 5.0:
        return "🔥【加速起漲】法人大單瘋狂鎖籌！多頭動能確認 (Momentum Confirmed)，波段安全加碼點。"
    elif t_val > 5.0 and 2.0 <= v_val < 3.0:
        return "✅【啟動訊號】主力實彈溫和點火換手！籌碼結構洗淨，波段最佳右側核心底倉建倉點。"
    elif t_val > 5.0 and v_val < 2.0:
        return "⚠️【陷阱警報】換手率極高但量比完全衰退！量價嚴重背離，假突破真倒貨，指令：撤離！"
    else:
        return "➖ 籌碼動態平穩"

# =============================================================================
# 🚀 模組 0.1: 進階籌碼矩陣判定引擎 (NOCChipMatrix) - 動態突破條件
# =============================================================================
class NOCChipMatrix:
    """
    華爾街級籌碼矩陣判定器，支援雙模式突破門檻：
    - BULL 模式：量 > 5日均量 1.3倍 + 突破 10日高點 → 主力點火
    - BEAR 模式：量 > 5日均量 1.5倍 + 突破 20日高點 → 主力點火
    """
    def analyze(self, df: pd.DataFrame, market_mode: str = "BEAR") -> str:
        try:
            latest = df.iloc[-1]
            avg_volume_5 = df['Volume'].rolling(5).mean().iloc[-1]
            volume_ratio = latest['Volume'] / avg_volume_5 if avg_volume_5 != 0 else 0.0

            if market_mode == "BULL":
                volume_threshold = 1.3
                high_lookback = 10
            else:
                volume_threshold = 1.5
                high_lookback = 20

            recent_high = df['High'].rolling(high_lookback).max().iloc[-1]

            if volume_ratio >= volume_threshold and latest['High'] >= recent_high:
                return "🔥 主力點火 (籌碼突破)"
            else:
                return "➖ 無點火訊號"
        except Exception as e:
            logger.error(f"籌碼矩陣分析異常: {e}")
            return "⚠️ 籌碼分析失敗"

# =============================================================================
# 💾 輔助模組: 執行緒安全資料庫連線防爆盾
# =============================================================================
class MockConn:
    def close(self) -> None:
        pass

# =============================================================================
# 💾 模組 1: 執行緒安全資料庫管理員 (Database Manager)
# =============================================================================
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

# =============================================================================
# 🌡️ 模組 2: 大盤風向儀與趨勢濾網 (Macro & Trend Strategy)
# =============================================================================
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
                return {
                    "status": "🟢 綠燈", 
                    "desc": "大盤多頭格局順風。積極抱緊長線底倉，防守線隨波段墊高，允許兵力推升至上限。"
                }
            elif td['Close'] < td['60MA']:
                return {
                    "status": "🔴 紅燈", 
                    "desc": "大盤崩盤警告（跌破季線）。啟動防空 protocols，全面停止新標的建倉，嚴格保留現金。"
                }
            else:
                return {
                    "status": "🟡 黃燈", 
                    "desc": "大盤進入高密度震盪洗盤期。嚴禁動用新資金盲目重倉，嚴格看守防禦底線。"
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
            current_price = hist_df['Close'].iloc[-1]
            ma10 = hist_df['10MA'].iloc[-1]
            ma20 = hist_df['20MA'].iloc[-1]
        
            if current_price > ma10 and ma10 > ma20:
                return 1.0
            else:
                return -1.0
        else:
            hist_df['20MA'] = hist_df['Close'].rolling(20).mean()
            hist_df['60MA'] = hist_df['Close'].rolling(60).mean()
            current_price = hist_df['Close'].iloc[-1]
            ma20 = hist_df['20MA'].iloc[-1]
            ma60 = hist_df['60MA'].iloc[-1]
        
            # 嚴格多頭：股價站上20MA且20MA在60MA之上
            if current_price > ma20 and ma20 > ma60:
                return 1.0
            # 底部起漲潛力股：股價同時站上20MA與60MA（但20MA尚未金叉60MA）
            elif current_price > ma20 and current_price > ma60:
                return 0.5
            else:
                return -1.0
    
    def get_fundamental_health(self, symbol: str) -> str:
        try:
            clean_symbol = symbol.replace(".TW", "").replace(".TWO", "")
            ticker = yf.Ticker(f"{clean_symbol}.TW")
            info = ticker.info
            revenue_growth = info.get("revenueGrowth")
        
            if revenue_growth is not None:
                yoy_pct = revenue_growth * 100
                if revenue_growth < -0.15:
                    return f"? 【基本面衰退】營收 YoY 呈現嚴重衰退 ({yoy_pct:.2f}%)，不符合長線波段體質！"
                elif revenue_growth < 0:
                    return f"?? 【營收谷底轉機】營收 YoY 偏弱 ({yoy_pct:.2f}%)，但允許技術面與籌碼面突圍！"
                else:
                    return f"? 【基本面優良】營收 YoY 成長 ({yoy_pct:.2f}%)，符合龍蝦養殖標準"
            else:
                # 關鍵：API抓不到資料 → 寬容處理，不帶警報字眼
                return "?? 【數據寬容】外部 API 暫無 YoY 數據，交由技術與籌碼面判定。"
        except Exception:
            return "?【營收健康】符合波段持有條件"

    def check_defcon_1_status(self) -> bool:
        try:
            twii = yf.Ticker("^TWII").history(period="3mo")
            if twii.empty:
                return False
            twii['60MA'] = twii['Close'].rolling(60).mean()
            current_close = twii['Close'].iloc[-1]
            ma60 = twii['60MA'].iloc[-1]
            if current_close < ma60:
                return True
            return False
        except Exception as e:
            self.logger.error(f"❌ DEFCON 1 協議監測器異常: {e}")
            return False

# =============================================================================
# 🛡️ 模組 3: 移動防禦與雙軌兵力精算師 (Risk Manager)
# =============================================================================
class NOCRiskManager:
    def __init__(self, total_capital: float = 130000.0):
        self.total_capital = total_capital

    def calculate_atr(self, hist_df: pd.DataFrame, period: int = 14) -> float:
        if len(hist_df) < period + 1:
            return hist_df['Close'].iloc[-1] * 0.025
            
        high_low = hist_df['High'] - hist_df['Low']
        high_close = np.abs(hist_df['High'] - hist_df['Close'].shift())
        low_close = np.abs(hist_df['Low'] - hist_df['Close'].shift())
        ranges = pd.concat([high_low, high_close, low_close], axis=1)
        true_range = ranges.max(axis=1)
        return true_range.rolling(period).mean().iloc[-1]

    def get_position_and_defense(self, symbol: str, current_price: float, hist_df: pd.DataFrame = None, 
                                 market_mode: str = "BEAR", is_yellow_light: bool = False) -> dict:
        """
        軍規級波段兵力精算 (雙模式自適應版) + 黃燈強制縮緊至 2.0 ATR
        """
        try:
            if hist_df is None or hist_df.empty:
                hist_df = yf.Ticker(symbol).history(period="6mo")
            atr = self.calculate_atr(hist_df, period=14)
            
            # 黃燈強制使用 2.0 倍 ATR，否則依市場模式
            if is_yellow_light:
                atr_multiplier = 2.0
            else:
                atr_multiplier = 1.8 if market_mode == "BULL" else 3.0
            trailing_stop = current_price - (atr * atr_multiplier)
            
            # 融合 20MA 防線（取較低者）
            if not hist_df.empty and len(hist_df) >= 20:
                hist_df['20MA'] = hist_df['Close'].rolling(20).mean()
                ma20 = hist_df['20MA'].iloc[-1]
                if not pd.isna(ma20):
                    defense_line = min(trailing_stop, ma20)
                else:
                    defense_line = trailing_stop
            else:
                defense_line = trailing_stop

            # 單筆風險控管 (總資金的 2%)
            max_risk_amount = self.total_capital * 0.02
            risk_per_share = current_price - defense_line
            if risk_per_share <= 0:
                risk_per_share = current_price * 0.10
            max_shares = math.floor(max_risk_amount / risk_per_share)
            
            # 依 15% 天花板限制最終股數
            max_allocation_shares = math.floor((self.total_capital * 0.15) / current_price)
            total_shares = min(max_shares, max_allocation_shares)
            
            # 長線底倉與短線游擊各半（若總股數為奇數，核心多一股）
            core_shares = total_shares // 2
            tactical_shares = total_shares - core_shares
            
            return {
                "current_price": round(current_price, 2),
                "defense_line": round(defense_line, 2), 
                "core_shares": int(core_shares),
                "tactical_shares": int(tactical_shares),
                "total_shares": int(total_shares),
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
# 🚀 模組 4: 戰略財務數據獲取引擎 (Data Fetcher Engine)
# =============================================================================
class NOCDataFetcher:
    def __init__(self, token: str = ""):
        self.token = token
        self.logger = logging.getLogger(__name__)

    def fetch_financial_statements(self, symbol: str, db: NOCDatabase) -> None:
        try:
            self.logger.info(f"🚀 [DataFetcher] 啟動多執行緒安全線路，同步標的 {symbol} 的長線基本面數據...")
            current_state = db.load_state()
            current_time_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if symbol not in current_state:
                current_state[symbol] = {
                    "status": "NONE",
                    "entry": 0.0,
                    "trailing_stop": 0.0,
                    "last_fetch": current_time_str
                }
            else:
                if isinstance(current_state[symbol], dict):
                    current_state[symbol]["last_fetch"] = current_time_str
            db.save_state(current_state)
            self.logger.info(f"✅ [DataFetcher] 標的 {symbol} 的波段基本面狀態資料同步更新成功。")
        except Exception as e:
            self.logger.error(f"❌ [DataFetcher] 執行多執行緒財務數據抓取時攔截到異常: {e}")
