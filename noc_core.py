import sqlite3
import pandas as pd
from FinMind.data import DataLoader
import datetime
import os

# ==========================================
# ⚙️ 模組 1：SQLite 戰情資料庫建置 (v2.1 裝甲防禦版)
# ==========================================
class NOCDatabase:
    def __init__(self, db_name="noc_warroom.db"):
        self.conn = sqlite3.connect(db_name)
        self.cursor = self.conn.cursor()
        self._create_tables()

    def _create_tables(self):
        """建立本地數據倉儲的資料表 (Schema)"""
        # 1. K線報價表
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS daily_kline (
                   date TEXT, stock_id TEXT, open REAL, max REAL, min REAL, close REAL, Trading_Volume INTEGER,
                   PRIMARY KEY (date, stock_id)
               )''')
        # 2. 融資融券表 (修正對應 FinMind 的 MarginPurchaseTodayBalance 欄位)
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS margin_trades (
                   date TEXT, stock_id TEXT, MarginPurchaseBuy INTEGER, MarginPurchaseSell INTEGER, MarginPurchaseTodayBalance INTEGER,
                   PRIMARY KEY (date, stock_id)
               )''')
        # 3. 大戶持股比例表
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS large_holders (
                   date TEXT, stock_id TEXT, HoldingSharesLevel TEXT, percent REAL,
                   PRIMARY KEY (date, stock_id, HoldingSharesLevel)
               )''')
        # 4. 外資期貨未平倉表
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS futures_institutional (
                   date TEXT, name TEXT, item TEXT, OpenInterestNetLot INTEGER,
                   PRIMARY KEY (date, name, item)
               )''')
        self.conn.commit()

    def save_df_to_db(self, df, table_name):
        """將 Pandas DataFrame 高速寫入 SQLite (具備自動過濾與防重複寫入機制)"""
        if df is None or df.empty:
            return

        # 1. 取得資料庫中該資料表的真實欄位名稱
        self.cursor.execute(f"PRAGMA table_info({table_name})")
        db_columns = [col[1] for col in self.cursor.fetchall()]

        # 2. 自動過濾 DataFrame，只保留資料庫有定義的欄位，丟棄 FinMind 塞進來的額外欄位
        valid_cols = [c for c in df.columns if c in db_columns]
        df_filtered = df[valid_cols]

        # 3. 寫入資料庫 (防衝突處理)
        try:
            # 先嘗試整批高速寫入
            df_filtered.to_sql(table_name, self.conn, if_exists='append', index=False, method='multi')
        except sqlite3.IntegrityError:
            # 若發生重複 (Primary Key 衝突)，改為逐筆寫入並忽略重複項
            for _, row in df_filtered.iterrows():
                try:
                    row.to_frame().T.to_sql(table_name, self.conn, if_exists='append', index=False)
                except sqlite3.IntegrityError:
                    pass # 忽略已經存在的歷史數據

# ==========================================
# 📡 模組 2：FinMind 數據抓取與更新
# ==========================================
class NOCDataFetcher:
    def __init__(self, token=None):
        self.api = DataLoader()
        if token:
            self.api.login_by_token(api_token=token)

    def fetch_and_store_stock_data(self, stock_id, start_date, db_instance):
        """抓取個股三大金礦並存入本地庫"""
        print(f"📡 正在抓取 {stock_id} 的最新戰情數據...")

        # 1. 基本 K 線
        df_kline = self.api.taiwan_stock_daily(stock_id=stock_id, start_date=start_date)
        db_instance.save_df_to_db(df_kline, 'daily_kline')

        # 2. 金礦一：融資融券餘額
        df_margin = self.api.taiwan_stock_margin_purchase_short_sale(stock_id=stock_id, start_date=start_date)
        db_instance.save_df_to_db(df_margin, 'margin_trades')

        # 3. 金礦二：股權分散表 (大戶持股比例)
        df_holders = self.api.taiwan_stock_holding_shares_per(stock_id=stock_id, start_date=start_date)
        db_instance.save_df_to_db(df_holders, 'large_holders')

    def fetch_market_health_data(self, start_date, db_instance):
        """金礦三：抓取大盤與外資期貨空單 (用於拔插頭協議)"""
        print("🚨 正在更新大盤防禦數據...")

        # 抓取加權指數 (TAIEX)
        df_taiex = self.api.taiwan_stock_daily(stock_id='TAIEX', start_date=start_date)
        db_instance.save_df_to_db(df_taiex, 'daily_kline')

        # 抓取外資台指期未平倉口數 (改用底層 get_data，避開版本參數衝突)
        df_futures = self.api.get_data(
            dataset='TaiwanFuturesInstitutionalInvestors',
            data_id='TX',
            start_date=start_date
        )
        db_instance.save_df_to_db(df_futures, 'futures_institutional')

# ==========================================
# 🛡️ 模組 3：DEFCON 1 緊急拔插頭協議 & 策略
# ==========================================
class NOCStrategy:
    def __init__(self, db_instance):
        self.db = db_instance

    def check_defcon_1_status(self):
        """檢查是否觸發大盤崩盤警報"""
        # 1. 取得加權指數最新 60 日均線 (季線
