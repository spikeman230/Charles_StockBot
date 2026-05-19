import sqlite3
import pandas as pd
from FinMind.data import DataLoader
import datetime
import os
import yfinance as yf
import logging

# ==========================================
# ⚙️ 模組 1：SQLite 戰情資料庫建置 (v2.3 雙引擎升級版)
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
        # 2. 融資融券表
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
        # 5. 美股/總經跨海連動表
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS macro_index (
            date TEXT, symbol TEXT, name TEXT, close REAL, pct_change REAL,
            PRIMARY KEY (date, symbol)
        )''')
        # 6. 🌟 新增：基本面財報透視表 (儲存 EPS, 毛利率, 營益率)
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS financial_statements (
            date TEXT, stock_id TEXT, type TEXT, value REAL,
            PRIMARY KEY (date, stock_id, type)
        )''')
        self.conn.commit()

    def save_df_to_db(self, df, table_name):
        """將 Pandas DataFrame 高速寫入 SQLite (具備自動過濾與防重複寫入機制)"""
        if df is None or df.empty:
            return

        self.cursor.execute(f"PRAGMA table_info({table_name})")
        db_columns = [col[1] for col in self.cursor.fetchall()]
        valid_cols = [c for c in df.columns if c in db_columns]
        df_filtered = df[valid_cols]

        try:
            df_filtered.to_sql(table_name, self.conn, if_exists='append', index=False, method='multi')
        except sqlite3.IntegrityError:
            for _, row in df_filtered.iterrows():
                try:
                    row.to_frame().T.to_sql(table_name, self.conn, if_exists='append', index=False)
                except sqlite3.IntegrityError:
                    pass


# ==========================================
# 📡 模組 2：數據抓取與更新引擎
# ==========================================

class NOCDataFetcher:
    def __init__(self, token=None):
        self.api = DataLoader()
        if token:
            self.api.login_by_token(api_token=token)

    def fetch_and_store_stock_data(self, stock_id, start_date, db_instance):
        """抓取個股戰情數據並存入本地庫"""
        print(f"📡 正在抓取 {stock_id} 的最新戰情數據...")
        # 1. 基本 K 線
        df_kline = self.api.taiwan_stock_daily(stock_id=stock_id, start_date=start_date)
        db_instance.save_df_to_db(df_kline, 'daily_kline')
        
        # 2. 融資融券餘額
        df_margin = self.api.taiwan_stock_margin_purchase_short_sale(stock_id=stock_id, start_date=start_date)
        db_instance.save_df_to_db(df_margin, 'margin_trades')

    def fetch_market_health_data(self, start_date, db_instance):
        """抓取大盤與外資期貨空單 (用於拔插頭協議)"""
        print("🚨 正在更新大盤防禦數據...")
        df_taiex = self.api.taiwan_stock_daily(stock_id='TAIEX', start_date=start_date)
        db_instance.save_df_to_db(df_taiex, 'daily_kline')
        
        df_futures = self.api.get_data(
            dataset='TaiwanFuturesInstitutionalInvestors',
            data_id='TX',
            start_date=start_date
        )
        db_instance.save_df_to_db(df_futures, 'futures_institutional')

    def fetch_us_macro_data(self, db_instance):
        """抓取美股四大關鍵指數與台積電 ADR"""
        print("🌎 啟動跨海雷達：掃描美股昨晚戰況...")
        macro_targets = {"^SOX": "費城半導體", "^IXIC": "納斯達克", "TSM": "台積電ADR"}
        macro_records = []
        tw_tz = datetime.timezone(datetime.timedelta(hours=8))
        
        for sym, name in macro_targets.items():
            try:
                ticker = yf.Ticker(sym)
                hist = ticker.history(period="5d")
                if len(hist) >= 2:
                    latest_date = hist.index[-1].astimezone(tw_tz).strftime('%Y-%m-%d')
                    latest_close = hist['Close'].iloc[-1]
                    prev_close = hist['Close'].iloc[-2]
                    pct_change = ((latest_close - prev_close) / prev_close) * 100
                    
                    macro_records.append({
                        "date": latest_date, "symbol": sym, "name": name,
                        "close": round(latest_close, 2), "pct_change": round(pct_change, 2)
                    })
            except Exception as e:
                logging.error(f"跨海雷達抓取 {sym} 失敗: {e}")
                
        if macro_records:
            df_macro = pd.DataFrame(macro_records)
            db_instance.save_df_to_db(df_macro, 'macro_index')

    def fetch_financial_statements(self, stock_id, db_instance):
        """🌟 新增：抓取財報基本面 (綜合損益表) 以計算毛利率與 EPS"""
        print(f"🏢 啟動基本面透視：掃描 {stock_id} 季報...")
        # 抓取過去 2 年的季報確保能做年增率(YoY)與季增率(QoQ)比較
        start_date = (datetime.datetime.now() - datetime.timedelta(days=730)).strftime('%Y-%m-%d')
        
        try:
            df_fin = self.api.taiwan_stock_financial_statement(stock_id=stock_id, start_date=start_date)
            if not df_fin.empty:
                # 篩選我們需要的關鍵財報指標：營業毛利率(GrossMargin)、營業利益率(OperatingMargin)、每股盈餘(EPS)
                target_types = ['GrossMargin', 'OperatingMargin', 'EPS']
                df_filtered = df_fin[df_fin['type'].isin(target_types)]
                db_instance.save_df_to_db(df_filtered, 'financial_statements')
        except Exception as e:
            logging.error(f"基本面財報抓取 {stock_id} 失敗: {e}")


# ==========================================
# 🛡️ 模組 3：戰略引擎 & 風控協議
# ==========================================

class NOCStrategy:
    def __init__(self, db_instance):
        self.db = db_instance

    def check_defcon_1_status(self):
        """檢查是否觸發大盤崩盤警報"""
        df_taiex = pd.read_sql("SELECT * FROM daily_kline WHERE stock_id='TAIEX' ORDER BY date DESC LIMIT 60", self.db.conn)
        if len(df_taiex) < 60: return False
            
        current_close = df_taiex.iloc[0]['close']
        ma60 = df_taiex['close'].mean()
        
        df_fii = pd.read_sql("SELECT * FROM futures_institutional WHERE name='外資及陸資' ORDER BY date DESC LIMIT 1", self.db.conn)
        fii_net_oi = df_fii.iloc[0]['OpenInterestNetLot'] if not df_fii.empty else 0
        
        return (current_close < ma60) and (fii_net_oi < -30000)

    def get_macro_sentiment(self):
        """讀取跨海雷達數據"""
        try:
            df_macro = pd.read_sql("SELECT * FROM macro_index ORDER BY date DESC LIMIT 3", self.db.conn)
            if df_macro.empty: return "🌎 跨海雷達未連線"
            
            sentiment_msg = "🌎 【昨夜美股戰況】: "
            is_danger = False
            for _, row in df_macro.iterrows():
                icon = "🔴" if row['pct_change'] < 0 else "🟢"
                sentiment_msg += f"{row['name']} {icon} {row['pct_change']}% | "
                if row['pct_change'] <= -3.0: is_danger = True
                    
            if is_danger: sentiment_msg += "\n⚠️ 警告：國際科技股重挫，今日台股突破訊號請縮小試單規模！"
            return sentiment_msg.strip(" | ")
        except Exception:
            return "🌎 跨海雷達解析異常"

    def analyze_stock_opportunity(self, stock_id):
        """整合 K線與籌碼金礦的進階過濾器"""
        df_margin = pd.read_sql(f"SELECT * FROM margin_trades WHERE stock_id='{stock_id}' ORDER BY date DESC LIMIT 5", self.db.conn)
        if len(df_margin) >= 2:
            if df_margin.iloc[0]['MarginPurchaseTodayBalance'] < df_margin.iloc[-1]['MarginPurchaseTodayBalance']:
                return "✅ 融資退潮，籌碼乾淨"
            else:
                return "⚠️ 融資暴增，散戶上車"
        return "資料不足"

    def get_fundamental_health(self, stock_id):
        """🌟 新增：讀取財報數據，回傳基本面健康度診斷"""
        try:
            # 取得該股票最近的財報資料
            df_fin = pd.read_sql(f"SELECT * FROM financial_statements WHERE stock_id='{stock_id}' ORDER BY date DESC", self.db.conn)
            if df_fin.empty:
                return ""

            # 抓取最新的毛利率與上一季的毛利率
            df_gm = df_fin[df_fin['type'] == 'GrossMargin'].head(2)
            df_eps = df_fin[df_fin['type'] == 'EPS'].head(1)
            
            msg = "🏢 【財報透視】: "
            
            if not df_eps.empty:
                latest_eps = df_eps.iloc[0]['value']
                msg += f"單季EPS {latest_eps:.2f} | "

            if len(df_gm) >= 2:
                curr_gm = df_gm.iloc[0]['value']
                prev_gm = df_gm.iloc[1]['value']
                msg += f"毛利率 {curr_gm:.2f}% "
                
                # 紅軍審計師攔截條件：毛利率顯著衰退 (例如下降超過 2%)
                if (prev_gm - curr_gm) >= 2.0:
                    msg += "📉 (⚠️ 毛利率嚴重衰退，當心做白工陷阱！)"
                elif curr_gm > prev_gm:
                    msg += "📈 (✅ 毛利率走升，體質轉佳)"
                else:
                    msg += "➖ (毛利率持平)"
                    
            return msg
        except Exception as e:
            logging.error(f"財報健康度診斷失敗: {e}")
            return ""
