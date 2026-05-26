# =============================================================================
# NOC 戰情室核心引擎 v16.1 - 支援 SQLite 盤後補給 (龍蝦養殖波段籌碼完全體)
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
from typing import Dict, Any, Optional, List

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)
logging.getLogger('yfinance').setLevel(logging.CRITICAL)


# =============================================================================
# 籌碼矩陣判定引擎 (保留原有功能)
# =============================================================================
def analyze_chip_tactics(turnover: float, volume_ratio: float, market_mode: str = "BEAR") -> str:
    t_val = turnover * 1.3 if market_mode == "BULL" else turnover
    v_val = volume_ratio * 1.3 if market_mode == "BULL" else volume_ratio

    if t_val > 10.0 and v_val > 5.0:
        return "🚀【極速發動】換手與量能極致爆發！"
    elif 5.0 <= t_val <= 10.0 and 3.0 <= v_val <= 5.0:
        return "🔥【加速起漲】法人大單瘋狂鎖籌！"
    elif t_val > 5.0 and 2.0 <= v_val < 3.0:
        return "✅【啟動訊號】主力實彈溫和點火！"
    elif t_val > 5.0 and v_val < 2.0:
        return "⚠️【陷阱警報】換手率極高但量比衰退！"
    else:
        return "➖ 籌碼動態平穩"


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
                logger.warning(f"找不到狀態檔 {self.db_path}，初始化新狀態。")
                return {}
            except Exception as e:
                logger.error(f"讀取狀態資料庫異常: {e}")
                return {}

    def save_state(self, data: dict) -> bool:
        with self._lock:
            try:
                with open(self.db_path, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=4)
                return True
            except Exception as e:
                logger.error(f"寫入狀態資料庫失敗: {e}")
                return False


class NOCStrategy:
    def __init__(self, db: Optional[NOCDatabase] = None):
        self.logger = logging.getLogger(__name__)
        self.db = db

    def get_macro_status(self) -> dict:
        try:
            twii = yf.Ticker("^TWII").history(period="6mo")
            if twii.empty:
                return {"status": "🟡 黃燈", "desc": "無法取得大盤資料，請嚴控資金。"}
            twii['20MA'] = twii['Close'].rolling(20).mean()
            twii['60MA'] = twii['Close'].rolling(60).mean()
            td = twii.iloc[-1]
            y_td = twii.iloc[-2]

            if td['Close'] > td['20MA'] and td['20MA'] >= y_td['20MA']:
                return {"status": "🟢 綠燈", "desc": "大盤多頭順風。"}
            elif td['Close'] < td['60MA']:
                return {"status": "🔴 紅燈", "desc": "大盤跌破季線，全面停止建倉。"}
            else:
                return {"status": "🟡 黃燈", "desc": "震盪洗盤期，嚴控資金。"}
        except Exception as e:
            self.logger.error(f"大盤風向儀異常: {e}")
            return {"status": "🟡 黃燈", "desc": "系統判定異常，啟動保護。"}

    def get_trend_score(self, hist_df: pd.DataFrame, market_mode: str = "BEAR") -> float:
        if len(hist_df) < 60:
            return -1.0
        if market_mode == "BULL":
            hist_df['10MA'] = hist_df['Close'].rolling(10).mean()
            hist_df['20MA'] = hist_df['Close'].rolling(20).mean()
            current = hist_df['Close'].iloc[-1]
            ma10 = hist_df['10MA'].iloc[-1]
            ma20 = hist_df['20MA'].iloc[-1]
            return 1.0 if (current > ma10 and ma10 > ma20) else -1.0
        else:
            hist_df['20MA'] = hist_df['Close'].rolling(20).mean()
            hist_df['60MA'] = hist_df['Close'].rolling(60).mean()
            current = hist_df['Close'].iloc[-1]
            ma20 = hist_df['20MA'].iloc[-1]
            ma60 = hist_df['60MA'].iloc[-1]
            if current > ma20 and ma20 > ma60:
                return 1.0
            elif current > ma20 and current > ma60:
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
                    return f"❌ 【基本面衰退】營收 YoY {yoy_pct:.2f}%"
                elif revenue_growth < 0:
                    return f"⚠️ 【營收谷底轉機】YoY {yoy_pct:.2f}%"
                else:
                    return f"✅ 【基本面優良】YoY {yoy_pct:.2f}%"
            else:
                return "⚠️ 【數據寬容】API 暫無 YoY 數據"
        except Exception:
            return "✅【營收健康】"

    def check_defcon_1_status(self) -> bool:
        try:
            twii = yf.Ticker("^TWII").history(period="3mo")
            if twii.empty:
                return False
            twii['60MA'] = twii['Close'].rolling(60).mean()
            current = twii['Close'].iloc[-1]
            ma60 = twii['60MA'].iloc[-1]
            return current < ma60
        except Exception as e:
            self.logger.error(f"DEFCON 1 監測異常: {e}")
            return False


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
        try:
            if hist_df is None or hist_df.empty:
                hist_df = yf.Ticker(symbol).history(period="6mo")
            atr = self.calculate_atr(hist_df, period=14)
            if is_yellow_light:
                atr_multiplier = 2.0
            else:
                atr_multiplier = 1.8 if market_mode == "BULL" else 3.0
            trailing_stop = current_price - (atr * atr_multiplier)

            if not hist_df.empty and len(hist_df) >= 20:
                hist_df['20MA'] = hist_df['Close'].rolling(20).mean()
                ma20 = hist_df['20MA'].iloc[-1]
                defense_line = min(trailing_stop, ma20) if not pd.isna(ma20) else trailing_stop
            else:
                defense_line = trailing_stop

            max_risk_amount = self.total_capital * 0.02
            risk_per_share = current_price - defense_line
            if risk_per_share <= 0:
                risk_per_share = current_price * 0.10
            max_shares = math.floor(max_risk_amount / risk_per_share)
            max_allocation_shares = math.floor((self.total_capital * 0.15) / current_price)
            total_shares = min(max_shares, max_allocation_shares)
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
            logger.error(f"部位精算異常: {e}")
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
# 升級版 NOCDataFetcher - 支援 SQLite 盤後補給
# =============================================================================
class NOCDataFetcher:
    def __init__(self, token: str = "", db_path: str = "noc_warroom.db"):
        self.token = token
        self.logger = logging.getLogger(__name__)
        self.db_path = db_path
        self._init_db()

    def _init_db(self):
        """初始化 SQLite 資料庫與表格"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS stock_prices (
                symbol TEXT,
                date TEXT,
                open REAL,
                high REAL,
                low REAL,
                close REAL,
                volume INTEGER,
                PRIMARY KEY (symbol, date)
            )
        ''')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_symbol_date ON stock_prices (symbol, date)')
        conn.commit()
        conn.close()
        self.logger.info(f"SQLite 資料庫初始化完成: {self.db_path}")

    def fetch_and_store_stock_data(self, symbol: str, start_date: str, db: NOCDatabase = None) -> None:
        """抓取個股歷史日K並存入 SQLite（用於 init_db / update_db）"""
        try:
            self.logger.info(f"📦 正在抓取 {symbol} 歷史資料 ({start_date} ~ 至今)...")
            stock = yf.Ticker(symbol)
            hist = stock.history(start=start_date)
            if hist.empty:
                self.logger.warning(f"{symbol} 無歷史資料")
                return
            conn = sqlite3.connect(self.db_path)
            for date, row in hist.iterrows():
                date_str = date.strftime("%Y-%m-%d")
                conn.execute('''
                    INSERT OR REPLACE INTO stock_prices (symbol, date, open, high, low, close, volume)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (symbol, date_str, row['Open'], row['High'], row['Low'], row['Close'], int(row['Volume'])))
            conn.commit()
            conn.close()
            self.logger.info(f"✅ {symbol} 歷史資料已存入資料庫")
        except Exception as e:
            self.logger.error(f"❌ 抓取 {symbol} 歷史資料失敗: {e}")

    def fetch_market_health_data(self, start_date: str, db: NOCDatabase = None) -> None:
        """抓取大盤指數 (^TWII) 歷史資料存入 SQLite"""
        try:
            self.logger.info(f"📊 正在抓取大盤指數歷史資料 ({start_date} ~ 至今)...")
            twii = yf.Ticker("^TWII")
            hist = twii.history(start=start_date)
            if hist.empty:
                self.logger.warning("大盤指數無歷史資料")
                return
            conn = sqlite3.connect(self.db_path)
            for date, row in hist.iterrows():
                date_str = date.strftime("%Y-%m-%d")
                conn.execute('''
                    INSERT OR REPLACE INTO stock_prices (symbol, date, open, high, low, close, volume)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', ("^TWII", date_str, row['Open'], row['High'], row['Low'], row['Close'], int(row['Volume'])))
            conn.commit()
            conn.close()
            self.logger.info("✅ 大盤指數歷史資料已存入資料庫")
        except Exception as e:
            self.logger.error(f"❌ 抓取大盤歷史資料失敗: {e}")

    def fetch_financial_statements(self, symbol: str, db: NOCDatabase) -> None:
        """保留原有功能：僅更新時間戳（不實際抓取財報）"""
        try:
            self.logger.info(f"🚀 [DataFetcher] 同步標的 {symbol} 的長線基本面數據...")
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
            self.logger.info(f"✅ 標的 {symbol} 的波段基本面狀態資料同步更新成功。")
        except Exception as e:
            self.logger.error(f"❌ [DataFetcher] 執行多執行緒財務數據抓取時攔截到異常: {e}")
