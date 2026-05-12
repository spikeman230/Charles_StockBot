import sqlite3
import pandas as pd
from FinMind.data import DataLoader
import datetime
import os

class NOCDatabase:
    def __init__(self, db_name="noc_warroom.db"):
        self.conn = sqlite3.connect(db_name)
        self.cursor = self.conn.cursor()
        self._create_tables()

    def _create_tables(self):
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS daily_kline (
            date TEXT, stock_id TEXT, open REAL, max REAL, min REAL, close REAL, Trading_Volume INTEGER,
            PRIMARY KEY (date, stock_id)
        )''')
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS margin_trades (
            date TEXT, stock_id TEXT, MarginPurchaseBuy INTEGER, MarginPurchaseSell INTEGER, MarginPurchaseTodayBalance INTEGER,
            PRIMARY KEY (date, stock_id)
        )''')
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS large_holders (
            date TEXT, stock_id TEXT, HoldingSharesLevel TEXT, percent REAL,
            PRIMARY KEY (date, stock_id, HoldingSharesLevel)
        )''')
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS futures_institutional (
            date TEXT, name TEXT, OpenInterestNetLot INTEGER,
            PRIMARY KEY (date, name)
        )''')
        self.conn.commit()

class NOCStrategy:
    def __init__(self, db_obj):
        self.db = db_obj

    def check_defcon_1_status(self):
        df_taiex = pd.read_sql("SELECT * FROM daily_kline WHERE stock_id='TAIEX' ORDER BY date DESC LIMIT 60", self.db.conn)
        if len(df_taiex) < 60: return False
        current_close = df_taiex.iloc[0]['close']
        ma60 = df_taiex['close'].mean()
        df_fii = pd.read_sql("SELECT * FROM futures_institutional WHERE name='外資及陸資' ORDER BY date DESC LIMIT 1", self.db.conn)
        fii_net_oi = df_fii.iloc[0]['OpenInterestNetLot'] if not df_fii.empty else 0
        return current_close < ma60 and fii_net_oi < -30000

    def analyze_stock_opportunity(self, stock_id):
        """🌟 v12.5 強化版：新增散戶退場偵測"""
        try:
            df_margin = pd.read_sql(f"SELECT * FROM margin_trades WHERE stock_id='{stock_id}' ORDER BY date DESC LIMIT 5", self.db.conn)
            if len(df_margin) < 5: return "資料不足"
            
            latest_bal = df_margin.iloc[0]['MarginPurchaseTodayBalance']
            prev_bal = df_margin.iloc[4]['MarginPurchaseTodayBalance']
            margin_change_pct = (latest_bal - prev_bal) / prev_bal if prev_bal != 0 else 0
            
            # 🌟 邏輯優化：散戶退場視為轉強訊號
            if margin_change_pct < -0.02: return "🟢 散戶退場 (籌碼洗淨)"
            if margin_change_pct > 0.10: return "⚠️ 融資暴增 (散戶上車)"
            return "➖ 融資穩定"
        except:
            return "分析異常"
