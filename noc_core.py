# =============================================================================
# NOC 戰情室核心引擎 (noc_core.py) v13.5 長線波段版
# 功能：大盤風向儀 (紅黃綠燈)、移動防禦精算 (ATR)、雙軌兵力配置、本地資料庫
# =============================================================================

import yfinance as yf
import pandas as pd
import numpy as np
import threading
import json
import logging
import math
from typing import Dict, Any

logger = logging.getLogger(__name__)
# 靜音 yfinance 錯誤
logging.getLogger('yfinance').setLevel(logging.CRITICAL)

# =============================================================================
# 🌡️ 模組 1: 大盤風向儀 (Macro Status)
# =============================================================================
class NOCStrategy:
    def __init__(self):
        self.logger = logging.getLogger(__name__)

    def get_macro_status(self) -> dict:
        """
        大盤風向儀：判定目前台股大盤的紅綠燈狀態，決定總體戰略。
        """
        try:
            twii = yf.Ticker("^TWII").history(period="6mo")
            if twii.empty:
                return {"status": "🟡 黃燈", "desc": "無法取得大盤資料，預設為震盪洗盤，請嚴控資金。"}

            twii['10MA'] = twii['Close'].rolling(10).mean()
            twii['20MA'] = twii['Close'].rolling(20).mean()
            twii['60MA'] = twii['Close'].rolling(60).mean()

            td = twii.iloc[-1]
            y_td = twii.iloc[-2]

            # 🟢 綠燈區 (大盤站上月線，且月線上揚)
            if td['Close'] > td['20MA'] and td['20MA'] >= y_td['20MA']:
                return {
                    "status": "🟢 綠燈", 
                    "desc": "大盤順風 (多頭攻擊)。抱緊庫存，防守線上移，可將總兵力推升至上限。"
                }
            # 🔴 紅燈區 (大盤跌破季線 60MA)
            elif td['Close'] < td['60MA']:
                return {
                    "status": "🔴 紅燈", 
                    "desc": "空頭來襲 (跌破季線)。所有跌破 20MA 的庫存強制清倉，保留現金，停止建倉。"
                }
            # 🟡 黃燈區 (其餘震盪洗盤期)
            else:
                return {
                    "status": "🟡 黃燈", 
                    "desc": "震盪洗盤 (跌破短均線)。若庫存跌破成本價強制減碼一半；嚴禁動用新資金全倉買進。"
                }
                
        except Exception as e:
            self.logger.error(f"大盤風向儀異常: {e}")
            return {"status": "🟡 黃燈", "desc": "系統判定異常，強制啟動震盪保護機制。"}

# =============================================================================
# 🛡️ 模組 2: 移動防禦與兵力精算師 (Risk Manager)
# =============================================================================
class NOCRiskManager:
    def __init__(self, total_capital: float = 130000.0):
        self.total_capital = total_capital

    def calculate_atr(self, hist_df: pd.DataFrame, period: int = 14) -> float:
        """計算 ATR (真實波動幅度)，用來抓取股票的『股性』"""
        if len(hist_df) < period + 1:
            return hist_df['Close'].iloc[-1] * 0.02 # 防呆：預設 2% 波動
            
        high_low = hist_df['High'] - hist_df['Low']
        high_close = np.abs(hist_df['High'] - hist_df['Close'].shift())
        low_close = np.abs(hist_df['Low'] - hist_df['Close'].shift())
        
        ranges = pd.concat([high_low, high_close, low_close], axis=1)
        true_range = np.max(ranges, axis=1)
        return true_range.rolling(period).mean().iloc[-1]

    def get_position_and_defense(self, symbol: str, current_price: float, hist_df: pd.DataFrame = None) -> dict:
        """
        軍規級部位精算：廢除固定停利，計算絕對防守價位與雙軌兵力配置
        """
        if hist_df is None or hist_df.empty:
            hist_df = yf.Ticker(symbol).history(period="3mo")

        atr = self.calculate_atr(hist_df)
        
        # 1. 雙軌兵力切割 (15% 總兵力拆成兩半)
        max_allocation = self.total_capital * 0.15
        core_capital = max_allocation * 0.5      # 7.5% 長線底倉
        tactical_capital = max_allocation * 0.5  # 7.5% 短線游擊
        
        core_shares = math.floor(core_capital / current_price)
        tactical_shares = math.floor(tactical_capital / current_price)

        # 2. 移動防禦線 (Trailing Stop)
        # 預設以現價減去 1.5 倍 ATR 作為市場正常波動的容錯空間
        trailing_stop = current_price - (atr * 1.5)
        
        # 結合技術面：抓取月線(20MA)，防線取兩者較低者，避免主力惡意洗盤被洗出場
        if not hist_df.empty:
            hist_df['20MA'] = hist_df['Close'].rolling(20).mean()
            ma20 = hist_df['20MA'].iloc[-1]
            defense_line = min(trailing_stop, ma20) if not pd.isna(ma20) else trailing_stop
        else:
            defense_line = trailing_stop

        return {
            "current_price": round(current_price, 2),
            "defense_line": round(defense_line, 2), 
            "core_shares": core_shares,           # 長線底倉建議股數
            "tactical_shares": tactical_shares,   # 短線游擊建議股數
            "total_shares": core_shares + tactical_shares,
            "risk_per_share": round(current_price - defense_line, 2)
        }

# =============================================================================
# 💾 模組 3: 執行緒安全資料庫 (Thread-Safe Database)
# =============================================================================
class NOCDatabase:
    def __init__(self, db_path: str = "noc_state.json"):
        self.db_path = db_path
        self._lock = threading.Lock()

    def load_state(self) -> dict:
        with self._lock:
            try:
                with open(self.db_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except (FileNotFoundError, json.JSONDecodeError):
                return {}

    def save_state(self, data: dict) -> bool:
        with self._lock:
            try:
                with open(self.db_path, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=4)
                return True
            except Exception as e:
                logger.error(f"資料庫寫入失敗: {e}")
                return False
