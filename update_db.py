# update_db.py
from noc_core import NOCDatabase, NOCDataFetcher
import datetime
import os
from dotenv import load_dotenv

load_dotenv()
FINMIND_TOKEN = os.getenv("FINMIND_TOKEN")

if __name__ == "__main__":
    print("🚚 開始執行盤後資料庫補給作業...")
    db = NOCDatabase()
    fetcher = NOCDataFetcher(token=FINMIND_TOKEN)
    
    # 預設抓取近 3 天的資料進行增量更新
    start_date = (datetime.datetime.now() - datetime.timedelta(days=3)).strftime("%Y-%m-%d")
    
    try:
        # 1. 更新大盤與外資期貨防空數據
        fetcher.fetch_market_health_data(start_date, db)
        
        # 2. 這裡可以匯入您的 150 檔 SCAN_LIST 迴圈
        # 示範更新宏碁、華新、力成、中鼎
        target_stocks = ["2353", "1605", "6239", "9933"]
        for sym in target_stocks:
            fetcher.fetch_and_store_stock_data(sym, start_date, db)
            
        print("✅ 戰情資料庫補給完畢！")
    except Exception as e:
        print(f"⚠️ 補給過程發生錯誤: {e}")
