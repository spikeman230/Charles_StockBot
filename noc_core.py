# noc_core.py (v13.5 長線波段轉型重構版)
import yfinance as yf
import pandas as pd
import numpy as np
import threading
import json
import logging
import math
from typing import Dict, Any

logger = logging.getLogger(__name__)
logging.getLogger('yfinance').setLevel(logging.CRITICAL)

class NOCStrategy:
    def __init__(self):
        self.logger = logging.getLogger(__name__)

    def get_macro_status(self) -> dict:
        """大盤風向儀：判定台股大盤狀態"""
        try:
            twii = yf.Ticker("^TWII").history(period="6mo")
            if twii.empty: return {"status": "🟡 黃燈", "desc": "無法取得資料"}
            
            twii['20MA'] = twii['Close'].rolling(20).mean()
            twii['60MA'] = twii['Close'].rolling(60).mean()
            td, y_td = twii.iloc[-1], twii.iloc[-2]

            if td['Close'] > td['20MA'] and td['20MA'] >= y_td['20MA']:
                return {"status": "🟢 綠燈", "desc": "多頭攻擊，抱緊長線底倉。"}
            elif td['Close'] < td['60MA']:
                return {"status": "🔴 紅燈", "desc": "跌破季線，執行拔插頭協議。"}
            else:
                return {"status": "🟡 黃燈", "desc": "震盪洗盤，嚴禁重倉。"}
        except Exception as e:
            return {"status": "🟡 黃燈", "desc": "判定異常，啟動防禦。"}

    def get_trend_score(self, hist_df: pd.DataFrame) -> float:
        """長線波段趨勢評分：計算 60MA 斜率與乖離率 [cite: 9]"""
        if len(hist_df) < 60: return -1.0
        ma60 = hist_df['Close'].rolling(60).mean()
        slope = (ma60.iloc[-1] - ma60.iloc[-5]) / 5
        bias = (hist_df['Close'].iloc[-1] - ma60.iloc[-1]) / ma60.iloc[-1]
        # 僅在均線向上且乖離不至於過大時，視為健康長線趨勢
        return 1.0 if (slope > 0 and bias < 0.15) else -1.0

    def get_fundamental_health(self, symbol: str) -> str:
        """強制基本面濾網：檢查 YoY 成長率 """
        from stock_bot import get_revenue_yoy # 假設跨模組呼叫
        yoy = get_revenue_yoy(symbol)
        if isinstance(yoy, float) and yoy < 0:
            return "⚠️【基本面警報】營收 YoY 衰退，禁止長線佈局 [cite: 7]"
        return "✅【營收健康】符合波段持有條件 [cite: 15]"

class NOCRiskManager:
    def __init__(self, total_capital: float = 130000.0):
        self.total_capital = total_capital

    def calculate_atr(self, hist_df: pd.DataFrame, period: int = 14) -> float:
        high_low = hist_df['High'] - hist_df['Low']
        high_close = np.abs(hist_df['High'] - hist_df['Close'].shift())
        low_close = np.abs(hist_df['Low'] - hist_df['Close'].shift())
        ranges = pd.concat([high_low, high_close, low_close], axis=1)
        return ranges.max(axis=1).rolling(period).mean().iloc[-1]

    def get_position_and_defense(self, symbol: str, current_price: float, hist_df: pd.DataFrame) -> dict:
        """軍規部位精算：7.5% 長線底倉 + 7.5% 短線游擊"""
        atr = self.calculate_atr(hist_df)
        max_allocation = self.total_capital * 0.15
        
        return {
            "core_shares": math.floor((max_allocation * 0.5) / current_price),
            "tactical_shares": math.floor((max_allocation * 0.5) / current_price),
            "defense_line": round(current_price - (atr * 1.5), 2)
        }

class NOCDatabase:
    def __init__(self, db_path: str = "noc_state.json"):
        self.db_path = db_path
        self._lock = threading.Lock()

    def load_state(self) -> dict:
        with self._lock:
            try:
                with open(self.db_path, "r", encoding="utf-8") as f: return json.load(f)
            except: return {}

    def save_state(self, data: dict) -> bool:
        with self._lock:
            try:
                with open(self.db_path, "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False, indent=4)
                return True
            except: return False
