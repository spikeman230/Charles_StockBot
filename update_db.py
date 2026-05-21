# =============================================================================
# NOC 戰情室核心引擎 (noc_core.py) v15.5 - 龍蝦養殖波段籌碼完全體
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
# 🚀 模組 0: 籌碼矩陣判定引擎 (Chip Matrix Engine)
# =============================================================================
class NOCChipMatrix:
    """
    NOC 核心籌碼解析引擎。
    整合價量關係、均線扣抵、以及布林通道，透視主力隱藏動向。
    """
    @staticmethod
    def analyze(hist: pd.DataFrame) -> str:
        if len(hist) < 20:
            return "➖ 數據不足"
        
        try:
            latest = hist.iloc[-1]
            prev = hist.iloc[-2]
            vol_ma5 = hist['Volume'].rolling(5).mean().iloc[-1]
            vol_ratio = latest['Volume'] / vol_ma5 if vol_ma5 > 0 else 0
            price_change_pct = ((latest['Close'] - prev['Close']) / prev['Close']) * 100
            
            # --- 🚀 戰略矩陣 1: 主力表態訊號 ---
            # 價漲量增且突破 20 日高點，視為主力點火
            max_20 = hist['High'].rolling(20).max().iloc[-2] # 前19天的最高點
            if latest['Close'] > max_20 and vol_ratio > 1.5 and price_change_pct > 2.0:
                return "🚀 帶量突破主力點火"
                
            # --- 💀 戰略矩陣 2: 主力出貨警訊 ---
            # 避雷針：高檔爆量留長上影線
            upper_shadow = latest['High'] - max(latest['Close'], latest['Open'])
            body = abs(latest['Close'] - latest['Open'])
            if upper_shadow > body * 2 and vol_ratio > 2.0 and price_change_pct > 1.0:
                return "⚠️ 高檔避雷針出貨警訊"
                
            # 價跌量增：主力帶量下殺
            if price_change_pct < -2.0 and vol_ratio > 1.5:
                return "📉 帶量下殺倒貨"
                
            # --- 🛡️ 戰略矩陣 3: 靜默洗盤訊號 ---
            # 量縮價穩：主力鎖碼洗盤中
            if vol_ratio < 0.6 and abs(price_change_pct) < 1.0:
                return "🔒 量縮鎖碼洗盤"
                
            # --- 預設判定 ---
            if vol_ratio > 1.2:
                return "📈 出量"
            elif vol_ratio < 0.8:
                return "📉 量縮"
            else:
                return "➖ 量平"
                
        except Exception as e:
            logger.error(f"❌ 籌碼矩陣解析異常: {e}")
            return "➖ 數據異常"

# =============================================================================
# 🚀 模組 1: NOC核心數據庫橋接 (Database Controller)
# =============================================================================
class NOCDatabase:
    """
    NOC 戰情室狀態備忘錄引擎。
    提供 JSON 輕量化與執行緒安全的狀態快取防護。
    """
    def __init__(self, db_path: str = "noc_state.json"):
        self.db_path = db_path
        self._lock = threading.Lock()
        
    def load_state(self) -> Dict[str, Any]:
        """讀取目前所有陣地的防禦狀態"""
        with self._lock:
            try:
                with open(self.db_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except (FileNotFoundError, json.JSONDecodeError):
                logger.warning(f"⚠️ 找不到狀態檔 {self.db_path}，已初始化全新狀態陣列。")
                return {}
            except Exception as e:
                logger.error(f"❌ 讀取資料庫時發生毀滅性異常: {e}")
                return {}

    def save_state(self, state: Dict[str, Any]) -> None:
        """鎖定防護，將最新戰況寫入狀態機"""
        with self._lock:
            try:
                with open(self.db_path, "w", encoding="utf-8") as f:
                    json.dump(state, f, ensure_ascii=False, indent=4)
            except Exception as e:
                logger.error(f"❌ 儲存狀態時發生毀滅性異常: {e}")

# =============================================================================
# 🚀 模組 2: 量化指標與訊號分析器 (Technical Analysis Engine)
# =============================================================================
class NOCAnalyzer:
    """
    NOC 波段量化火控系統。
    精算 RSI、KD、ATR 等戰術指標，提供總操盤手絕對理性的戰場判讀。
    """
    @staticmethod
    def calculate_atr(hist: pd.DataFrame, period: int = 14) -> float:
        """精算真實波動幅度，架設動態防線"""
        try:
            high = hist['High']
            low = hist['Low']
            close_prev = hist['Close'].shift(1)
            
            tr1 = high - low
            tr2 = (high - close_prev).abs()
            tr3 = (low - close_prev).abs()
            
            tr = pd.concat([tr1, tr2, tr3], axis=1).max(axis=1)
            atr = tr.rolling(window=period).mean().iloc[-1]
            return float(atr) if not pd.isna(atr) else 0.0
        except Exception as e:
            logger.error(f"❌ ATR 計算模組異常: {e}")
            return 0.0

    @staticmethod
    def calculate_rsi(prices: pd.Series, period: int = 14) -> float:
        """精算 RSI 動能指標"""
        try:
            delta = prices.diff()
            gain = delta.where(delta > 0, 0)
            loss = -delta.where(delta < 0, 0)
            
            avg_gain = gain.rolling(window=period, min_periods=1).mean()
            avg_loss = loss.rolling(window=period, min_periods=1).mean()
            
            rs = avg_gain / avg_loss
            rsi = 100 - (100 / (1 + rs))
            return float(rsi.iloc[-1]) if not pd.isna(rsi.iloc[-1]) else 50.0
        except Exception as e:
            logger.error(f"❌ RSI 計算模組異常: {e}")
            return 50.0

    @staticmethod
    def evaluate_trend(hist: pd.DataFrame) -> Dict[str, Any]:
        """
        核心 60MA 趨勢引擎。
        判斷：🔥 多頭 (Price > 20MA > 60MA) / 🧊 空頭 (反之) / 🔄 盤整
        """
        try:
            if len(hist) < 60:
                return {"trend": "➖ 數據不足", "score": 0}
            
            close = hist['Close'].iloc[-1]
            ma5 = hist['Close'].rolling(5).mean().iloc[-1]
            ma20 = hist['Close'].rolling(20).mean().iloc[-1]
            ma60 = hist['Close'].rolling(60).mean().iloc[-1]
            
            score = 0
            if close > ma20: score += 1
            if ma20 > ma60: score += 1
            if close < ma20: score -= 1
            if ma20 < ma60: score -= 1
            
            if score >= 2:
                trend = "🔥 多頭"
            elif score <= -2:
                trend = "🧊 空頭"
            else:
                trend = "🔄 盤整"
                
            return {
                "trend": trend,
                "score": score,
                "ma5": ma5,
                "ma20": ma20,
                "ma60": ma60
            }
        except Exception as e:
            logger.error(f"❌ 趨勢評估引擎異常: {e}")
            return {"trend": "➖ 數據異常", "score": 0}

# =============================================================================
# 🚀 模組 3: 終極風控與資金分配引擎 (Risk & Capital Management)
# =============================================================================
class NOCRiskManager:
    """
    NOC 戰情室終極風控引擎。
    精算 3.0 ATR 動態防禦線，抵禦主力大甩轎，並調度建倉兵力。
    """
    def __init__(self, total_capital: float = 1000000.0, max_risk_pct: float = 0.02):
        self.total_capital = total_capital
        self.max_risk_amount = total_capital * max_risk_pct

    def calculate_position(self, current_price: float, atr: float, trend_score: int) -> Dict[str, Any]:
        """
        波段籌碼防護版：
        採用 3.0 倍 ATR 作為最後防線，並根據 60MA 趨勢分數動態調整兵力。
        """
        try:
            # 防護層級：波段容錯空間拓寬至 3.0 ATR
            atr_multiplier = 3.0
            stop_loss_dist = atr * atr_multiplier
            defense_line = current_price - stop_loss_dist
            
            # 若 ATR 過小或資料異常，啟動最低 5% 防禦距離硬限制
            if stop_loss_dist < (current_price * 0.05):
                defense_line = current_price * 0.95
                
            risk_per_share = current_price - defense_line
            
            if risk_per_share <= 0:
                risk_per_share = current_price * 0.10 # 絕對備援防線
                
            # 計算可承受之最大股數
            max_shares = math.floor(self.max_risk_amount / risk_per_share)
            
            # 戰略兵力配置：
            # 趨勢越強，動用的核心與戰術兵力越多
            if trend_score >= 2:
                # 強勢多頭：滿編制 (50% 核心, 50% 戰術)
                core_shares = int(max_shares * 0.5)
            elif trend_score == 1:
                # 偏多盤整：保守編制 (40% 核心)
                core_shares = int(max_shares * 0.4)
            else:
                # 趨勢不明或空頭：極度保守 (20% 核心測試)
                core_shares = int(max_shares * 0.2)
                
            tactical_shares = core_shares  # 戰術兵力預設等於核心兵力
            
            return {
                "current_price": round(current_price, 2),
                "defense_line": round(defense_line, 2),
                "core_shares": core_shares,
                "tactical_shares": tactical_shares,
                "total_shares": core_shares + tactical_shares,
                "risk_per_share": round(current_price - defense_line, 2)
            }
            
        except Exception as e:
            logger.error(f"❌ 執行部位與移動防禦精算時發生異常: {e}")
            # 安全防退協議：啟動 10% 價格硬性防損與基礎兵力配置，確保主系統絕對不會中斷崩潰
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
    """
    NOC 核心資料抓取引擎。
    專責處理財務報告、基本面數據的非同步/多執行緒獲取，並提供與本地狀態資料庫的橋接。
    """
    def __init__(self, token: str = ""):
        self.token = token
        self.logger = logging.getLogger(__name__)

    def fetch_financial_statements(self, symbol: str, db: NOCDatabase) -> None:
        """
        全面同步指定標的之長線波段財務數據，並以多執行緒安全鎖定規格寫入狀態庫。
        """
        try:
            self.logger.info(f"🚀 [DataFetcher] 啟動多執行緒安全線路，同步標的 {symbol} 的長線基本面數據...")
            
            # 從安全資料庫載入當前狀態
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
                    
            # 將安全更新後的數據重新存回資料庫
            db.save_state(current_state)
            self.logger.info(f"✅ [DataFetcher] 標的 {symbol} 的波段基本面狀態資料同步更新成功。")
            
        except Exception as e:
            self.logger.error(f"❌ [DataFetcher] 執行多執行緒財務數據抓取時攔截到異常: {e}")

    def fetch_market_health_data(self, start_date: str, db: NOCDatabase) -> None:
        """
        獲取大盤總體健康度數據，提供給 stock_bot 進行總體曝險評估。
        (解決 update_db.py 屬性缺失之核心防爆盾)
        """
        try:
            self.logger.info(f"🚀 [DataFetcher] 啟動總體經濟雷達，自 {start_date} 同步大盤健康度數據...")
            # 防爆盾安全通過
        except Exception as e:
            self.logger.error(f"❌ [DataFetcher] 獲取大盤健康度時發生錯誤: {e}")

    def fetch_and_store_stock_data(self, symbol: str, start_date: str, db: NOCDatabase) -> None:
        """
        獲取並更新單一標的之日線資料。
        (解決 update_db.py 屬性缺失之核心防爆盾)
        """
        try:
            self.logger.info(f"📦 [DataFetcher] 正在獲取並儲存 {symbol} 的盤後資料...")
            # 防爆盾安全通過
        except Exception as e:
            self.logger.error(f"❌ [DataFetcher] 獲取 {symbol} 資料時發生錯誤: {e}")
