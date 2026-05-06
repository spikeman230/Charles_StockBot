# init_db.py (只在系統建置第一天手動執行一次)
from noc_core import NOCDatabase, NOCDataFetcher
import datetime
import time
import os
from dotenv import load_dotenv

load_dotenv()
FINMIND_TOKEN = os.getenv("FINMIND_TOKEN")

if __name__ == "__main__":
    print("🚀 啟動 NOC 戰情室「建庫大補丸」歷史資料載入作業...")
    db = NOCDatabase()
    fetcher = NOCDataFetcher(token=FINMIND_TOKEN)
    
    # 🌟 關鍵：將時間往前推 400 天 (確保足以計算 240MA 年線)
    start_date = (datetime.datetime.now() - datetime.timedelta(days=400)).strftime("%Y-%m-%d")
    
    try:
        print("1️⃣ 正在下載大盤與防空歷史數據 (近400天)...")
        fetcher.fetch_market_health_data(start_date, db)
        
        # 假設這是您的 150 檔 SCAN_LIST
        # 這裡建議先用幾檔做測試，成功後再放 150 檔進去跑
        target_stocks = ["2353", "1605", "6239", "9933"] 
        
        print(f"2️⃣ 準備下載 {len(target_stocks)} 檔個股歷史數據...")
        for sym in target_stocks:
            fetcher.fetch_and_store_stock_data(sym, start_date, db)
            
            # ⚠️ 極度重要：FinMind API 有頻率限制！
            # 一次抓 400 天的資料負載較大，每抓完一檔務必強制休息 3~5 秒，否則會被鎖 IP。
            time.sleep(3) 
            
        print("✅ 歷史戰情資料庫 (SQLite) 初始建置與灌水完成！")
    except Exception as e:
        print(f"⚠️ 建庫過程發生錯誤: {e}")
