import sqlite3
import pandas as pd
from FinMind.data import DataLoader
import datetime
import os

# ==========================================
# ⚙️ 模組 1：SQLite 戰情資料庫建置
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
        # 2. 融資融券表 (散戶動向)
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS margin_trades (
            date TEXT, stock_id TEXT, MarginPurchaseBuy INTEGER, MarginPurchaseSell INTEGER, MarginPurchaseBalance INTEGER,
            PRIMARY KEY (date, stock_id)
        )''')
        # 3. 大戶持股比例表 (籌碼集中度)
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS large_holders (
            date TEXT, stock_id TEXT, HoldingSharesLevel TEXT, percent REAL,
            PRIMARY KEY (date, stock_id, HoldingSharesLevel)
        )''')
        # 4. 外資期貨未平倉表 (大盤防空警報用)
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS futures_institutional (
            date TEXT, name TEXT, item TEXT, OpenInterestNetLot INTEGER,
            PRIMARY KEY (date, name, item)
        )''')
        self.conn.commit()

    def save_df_to_db(self, df, table_name):
        """將 Pandas DataFrame 高速寫入 SQLite"""
        if not df.empty:
            df.to_sql(table_name, self.conn, if_exists='append', index=False, method='multi')

# ==========================================
# 📡 模組 2：FinMind 數據抓取與更新 (包含三大金礦)
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
        
        # 2. 金礦一：融資融券餘額 (MarginPurchaseBalance)
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
        
        # 抓取外資台指期未平倉口數
        df_futures = self.api.taiwan_futures_institutional_investors(
            commodity_id='TX', start_date=start_date
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
        # 1. 取得加權指數最新 60 日均線 (季線) 狀態
        df_taiex = pd.read_sql("SELECT * FROM daily_kline WHERE stock_id='TAIEX' ORDER BY date DESC LIMIT 60", self.db.conn)
        if len(df_taiex) < 60:
            return False # 數據不足跳過
            
        current_close = df_taiex.iloc[0]['close']
        ma60 = df_taiex['close'].mean()
        is_broken_ma60 = current_close < ma60
        
        # 2. 取得外資最新期貨淨未平倉口數 (OpenInterestNetLot)
        # item: 'Foreign_Investor' (外資)
        df_fii = pd.read_sql("SELECT * FROM futures_institutional WHERE name='外資及陸資' ORDER BY date DESC LIMIT 1", self.db.conn)
        fii_net_oi = df_fii.iloc[0]['OpenInterestNetLot'] if not df_fii.empty else 0
        is_heavy_short = fii_net_oi < -30000  # 外資淨空單超過三萬口
        
        # ⚔️ 判斷邏輯：跌破季線 且 外資重兵做空
        if is_broken_ma60 and is_heavy_short:
            return True
        return False

    def analyze_stock_opportunity(self, stock_id):
        """整合 K線與籌碼金礦的進階過濾器"""
        # 這裡示範如何從 SQLite 讀取數據並判斷
        df_margin = pd.read_sql(f"SELECT * FROM margin_trades WHERE stock_id='{stock_id}' ORDER BY date DESC LIMIT 5", self.db.conn)
        
        if len(df_margin) >= 2:
            current_margin = df_margin.iloc[0]['MarginPurchaseBalance']
            prev_margin = df_margin.iloc[4]['MarginPurchaseBalance']
            
            # 判斷標準：近五日融資必須是「減少」的 (散戶退場)
            if current_margin < prev_margin:
                return "✅ 融資退潮，籌碼乾淨，允許狙擊！"
            else:
                return "⚠️ 融資暴增，散戶上車，建議觀望。"
        return "資料不足"

# ==========================================
# 🚀 總司令執行中樞 (Main)
# ==========================================
if __name__ == "__main__":
    # 1. 初始化資料庫與爬蟲
    db = NOCDatabase()
    fetcher = NOCDataFetcher() # 如果有 FinMind Token 請傳入 token="YOUR_TOKEN"
    strategy = NOCStrategy(db)
    
    # 設定抓取起點 (實戰中可設為近三個月，每日例行更新只需抓近三天)
    start_date = (datetime.datetime.now() - datetime.timedelta(days=90)).strftime("%Y-%m-%d")
    
    # 2. 執行每日數據更新 (建議設定在 Cronjob 每天下午 15:30 執行)
    try:
        fetcher.fetch_market_health_data(start_date, db)
        fetcher.fetch_and_store_stock_data("2353", start_date, db) # 以宏碁為例
        print("✅ 戰情資料庫更新完畢！")
    except Exception as e:
        print(f"⚠️ 數據抓取失敗，可能觸發 API 限制：{e}")

    # 3. 系統性風險掃描 (拔插頭協議)
    print("\n--- 🚨 執行 DEFCON 大盤掃描 ---")
    is_defcon_1 = strategy.check_defcon_1_status()
    
    if is_defcon_1:
        print("🟥【警告】觸發 DEFCON 1 拔插頭協議！大盤跌破季線且外資空單破三萬口！")
        print("指令：強制鎖定買盤，通知 Telegram，調高現有持股停損點！")
        # 這裡可以寫入發送 Telegram 的程式碼
    else:
        print("🟩 大盤狀態安全，防空警報解除。繼續執行個股雷達掃描...")
        
        # 4. 個股籌碼金礦分析 (以宏碁為例)
        analysis_result = strategy.analyze_stock_opportunity("2353")
        print(f"🎯 宏碁 (2353) 籌碼判定：{analysis_result}")
