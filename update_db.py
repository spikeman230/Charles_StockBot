# init_db.py
import datetime
import os
import time
import re
import logging
import yfinance as yf
from dotenv import load_dotenv

from noc_core import NOCDatabase, NOCDataFetcher
from noc_radar import RadarConfig
SCAN_LIST = RadarConfig.SCAN_LIST

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

load_dotenv()
FINMIND_TOKEN = os.getenv("FINMIND_TOKEN")

def fetch_stock_robust(sym: str, start_date: str, fetcher: NOCDataFetcher, db: NOCDatabase):
    """
    雙重保障抓取機制：
    1. 嘗試使用 NOCDataFetcher (FinMind) 寫入
    2. 若失敗或無數據，自動改用 yfinance 抓取並備援寫入 SQLite
    """
    # 防護 1: 嘗試透過預設流程抓取
    try:
        fetcher.fetch_and_store_stock_data(sym, start_date, db)
    except Exception as e:
        logger.warning(f"⚠️ {sym} 透過預設 Fetcher 抓取失敗: {e}")

    # 防護 2: 驗證資料庫是否有資料，若無資料則調用 yfinance 強制補給
    try:
        df = db.get_stock_dataframe(sym, days=5)
        if df is None or df.empty:
            logger.info(f"🔄 觸發 yfinance 備援機制補給: {sym}")
            ticker = yf.Ticker(sym)
            hist = ticker.history(period="10d")
            if not hist.empty:
                import sqlite3
                with sqlite3.connect(db.db_path) as conn:
                    for idx, row in hist.iterrows():
                        date_str = idx.strftime("%Y-%m-%d")
                        conn.execute('''
                            INSERT OR REPLACE INTO stock_prices (symbol, date, open, high, low, close, volume, adj_close)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (sym, date_str, row['Open'], row['High'], row['Low'], row['Close'], int(row['Volume']), row['Close']))
                logger.info(f"✅ {sym} 透過 yfinance 成功補給寫入資料庫！")
            else:
                logger.error(f"❌ {sym} Yahoo Finance 亦無數據（可能代號不符或興櫃/暫停交易）")
    except Exception as e:
        logger.error(f"❌ 備援寫入 {sym} 時發生錯誤: {e}")

# 掃描池（與您的雷達清單相同）
SCAN_LIST : list = [
    # =========================================================
    # 區塊 1：總司令旗艦權值股 (共 50 檔，保留 0050 作為大盤基準)
    # =========================================================
    "0050.TW", "2330.TW", "2317.TW", "2454.TW", "2382.TW", "2308.TW", "3231.TW", "3037.TW",
    "2303.TW", "3008.TW", "3034.TW", "3711.TW", "2357.TW", "2395.TW", "2408.TW", "2353.TW",
    "2379.TW", "4938.TW", "2301.TW", "2345.TW", "2324.TW", "3661.TW", "6669.TW", "3714.TW",
    "2881.TW", "2882.TW", "2891.TW", "2886.TW", "2884.TW", "2892.TW", "2885.TW", "2880.TW",
    "2883.TW", "2887.TW", "5871.TW", "2890.TW", "5880.TW", "2002.TW", "1216.TW", "1301.TW",
    "1303.TW", "1326.TW", "2912.TW", "9904.TW", "2603.TW", "2609.TW", "2615.TW", "2207.TW",
    "1101.TW", "1102.TW",

    # =========================================================
    # 區塊 2：高動能科技、AI、半導體、光通訊與機器人 (共 105 檔)
    # =========================================================
    "2356.TW", "3163.TWO", "5388.TW", "8299.TWO", "3260.TWO", "2377.TW", "2383.TW", "3017.TW",
    "2352.TW", "3443.TW", "3529.TWO", "3293.TWO", "6488.TWO", "8069.TWO", "6274.TWO", "6239.TW",
    "3044.TW", "2449.TW", "2344.TW", "2409.TW", "3481.TW", "6116.TW", "4958.TW", "6176.TW",
    "3532.TW", "2371.TW", "2404.TW", "3702.TW", "8046.TW", "5483.TWO", "3105.TWO", "5347.TWO",
    "6147.TWO", "6214.TW", "2313.TW", "2368.TW", "3013.TW", "3019.TW", "3042.TW", "3324.TWO",
    "3533.TW", "3583.TW", "3653.TW", "4966.TWO", "5269.TW", "6269.TW", "6415.TW", "6531.TW",
    "8016.TW", "8081.TW", "8150.TW", "3376.TW", "3035.TW", "3227.TWO", "3131.TWO", "2451.TW",
    "5469.TW", "3413.TW", "3450.TW", "4919.TW",
    # [散熱與 AI 機殼]
    "2421.TW", "3483.TW", "8996.TW", "8210.TW", "6117.TW", "5426.TWO",
    # [半導體、矽智財與 CoWoS 設備]
    "6643.TW", "3228.TWO", "3014.TW", "4961.TW", "6799.TW", "3587.TWO", "3289.TWO", "6146.TWO", 
    "6187.TWO", "6196.TW", "6640.TWO", "5443.TWO", "6139.TW", "2464.TW", "2388.TW", "2439.TW",
    # [被動元件與 PCB (含國巨)]
    "2327.TW", "2492.TW", "3026.TW", "6213.TW", "6153.TW",
    # [網通與矽光子/光通訊]
    "3596.TW", "3380.TW", "6285.TW", "4979.TW", "4908.TW", "4977.TW", "6442.TW", "8114.TWO",
    # [機器人、智慧自動化與先進裝備]
    "1590.TW", "2359.TW", "6188.TW", "4583.TW", "8374.TW", "2365.TW", "4510.TW", "3680.TW", 
    "6667.TW", "3167.TW",

    # =========================================================
    # 區塊 3：重電、電纜與綠能 (共 28 檔)
    # =========================================================
    "1513.TW", "1514.TW", "1519.TW", "1605.TW", "1504.TW", "1503.TW", "1515.TW", "1520.TW",
    "3708.TW", "1609.TW", "1608.TW", "1611.TW", "1612.TW", "1618.TW", "9958.TW", "3712.TW",
    "6409.TW", "1582.TW", "1522.TW", "1532.TW", "4536.TW", "8926.TW", "6869.TW", "1537.TW",
    # [太陽能與儲能政策概念股]
    "6806.TW", "6443.TW", "3576.TW", "6477.TW",

    # =========================================================
    # 區塊 4：生技醫療與美容保健 (共 25 檔)
    # =========================================================
    "6472.TW", "6446.TWO", "1795.TW", "4142.TW", "1701.TW", "1707.TW", "1720.TW", "4123.TWO", 
    "1762.TW", "4104.TW", "3176.TWO", "4114.TW", "4736.TWO", "4162.TWO", "6547.TWO", "6561.TWO", 
    "4128.TWO", "4105.TWO", "1736.TW", "8436.TW", "6491.TW", "4137.TW", "6666.TW", "1733.TW", 
    "4743.TWO",

    # =========================================================
    # 區塊 5：傳產塑化、汽車、軍工航太與航運觀光 (共 42 檔)
    # =========================================================
    # [汽車零組件 13 檔]
    "1536.TW", "2231.TW", "1521.TW", "1525.TW", "2228.TW", "2115.TW", "2201.TW", "2204.TW",
    "3346.TW", "1339.TW", "6279.TW", "1524.TW", "1568.TW",
    # [傳產塑化化學 15 檔 - 已移去重複的 1522.TW，補入台勝科 3532 / 精材 3374]
    "1314.TW", "1717.TW", "1304.TW", "1308.TW", "1309.TW", "1312.TW", "1305.TW", "1710.TW",
    "1704.TW", "4722.TW", "4739.TW", "1718.TW", "1319.TW", "6605.TW", "7736.TW",
    # [造船與軍工航太 5 檔]
    "2208.TW", "2634.TW", "4541.TW", "8222.TW", "2646.TW",
    # [軍工與強勢航太材料 5 檔 - 補入 3374.TW]
    "5009.TW", "3005.TW", "1584.TWO", "8033.TWO", "3374.TW",
    # [航空、散裝航運與內需觀光 4 檔]
    "2618.TW", "2610.TW", "2637.TW", "2731.TW"
]

if __name__ == "__main__":
    logger.info("🚀 開始執行 NOC 盤後戰情資料庫補給作業 (250 檔對齊版)...")
    db = NOCDatabase()
    fetcher = NOCDataFetcher(token=FINMIND_TOKEN)

    # 延長日期範圍至 10 天，防止連假或週末抓不到數據
    start_date = (datetime.datetime.now() - datetime.timedelta(days=10)).strftime("%Y-%m-%d")

    # 去重檢查 (維持 250 檔不重複標的)
    target_stocks = list(dict.fromkeys(SCAN_LIST))
    logger.info(f"📊 鎖定 {len(target_stocks)} 檔目標（不重複），準備開始日常補給！")

    try:
        logger.info("1. 正在更新大盤指數與市場海象數據...")
        fetcher.fetch_market_health_data(start_date, db)

        logger.info("2. 開始逐檔精算個股盤後數據...")
        for idx, sym in enumerate(target_stocks, 1):
            fetch_stock_robust(sym, start_date, fetcher, db)
            time.sleep(0.3)  # 適當間隔，避免請求過快被 API Ban

        # 計算資料庫中實際存有多少檔股票
        import sqlite3
        with sqlite3.connect(db.db_path) as conn:
            cur = conn.cursor()
            cur.execute("SELECT COUNT(DISTINCT symbol) FROM stock_prices")
            distinct_count = cur.fetchone()[0]

        logger.info(f"\n🎉 戰情資料庫補給完畢！資料庫目前實存不重複標的共：{distinct_count} 檔！")

    except Exception as e:
        logger.error(f"\n❌ 補給過程發生未預期錯誤: {e}")

