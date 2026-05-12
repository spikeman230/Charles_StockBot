# init_db.py (只在系統建置第一天、或刪除資料庫後手動執行一次)
from noc_core import NOCDatabase, NOCDataFetcher
import datetime
import time
import os
from dotenv import load_dotenv

load_dotenv()
FINMIND_TOKEN = os.getenv("FINMIND_TOKEN")

# 🌟 修正點 1：清單必須帶有 .TW / .TWO 後綴，確保與戰情室雷達同步
# 您可以先用這四檔測試，成功後再把 150 檔的 SCAN_LIST 貼過來
TARGET_STOCKS = ["2353.TW", "1605.TW", "6239.TW", "9933.TW"]

if __name__ == "__main__":
    print("🚀 啟動 NOC 戰情室「建庫大補丸」歷史資料載入作業 (v12.5 強固版)...")
    db = NOCDatabase()
    fetcher = NOCDataFetcher(token=FINMIND_TOKEN)
    
    # 🌟 關鍵：將時間往前推 400 天 (確保足以計算 240MA 年線與 60MA 季線)
    start_date = (datetime.datetime.now() - datetime.timedelta(days=400)).strftime("%Y-%m-%d")
    
    try:
        print("1️⃣ 正在下載大盤與防空歷史數據 (近 400 天)...")
        fetcher.fetch_market_health_data(start_date, db)
        
        print(f"2️⃣ 準備下載 {len(TARGET_STOCKS)} 檔個股歷史數據...")
        for full_sym in TARGET_STOCKS:
            # 切割後綴給 FinMind API 查詢使用
            pure_sym = full_sym.split('.')[0]
            if not pure_sym.isdigit():
                continue

            print(f"🔄 正在抓取 {full_sym} (近 400 天)...")
            try:
                # 🌟 修正點 2：抓取 K 線並強制掛上後綴 (.TW / .TWO)
                df_kline = fetcher.api.taiwan_stock_daily(stock_id=pure_sym, start_date=start_date)
                if df_kline is not None and not df_kline.empty:
                    df_kline['stock_id'] = full_sym 
                    db.save_df_to_db(df_kline, 'daily_kline')

                # 🌟 修正點 2：抓取融資融券並強制掛上後綴
                df_margin = fetcher.api.taiwan_stock_margin_purchase_short_sale(stock_id=pure_sym, start_date=start_date)
                if df_margin is not None and not df_margin.empty:
                    df_margin['stock_id'] = full_sym 
                    db.save_df_to_db(df_margin, 'margin_trades')

            except Exception as inner_e:
                print(f"   ⚠️ {full_sym} 抓取失敗，跳過並繼續下一檔: {inner_e}")

            # 🌟 修正點 3：強化防爬蟲機制，設定 1.5 秒延遲
            time.sleep(1.5) 
            
        print("\n✅ 「建庫大補丸」執行完畢！歷史資料庫已裝填完畢！")
        
    except Exception as e:
        print(f"\n💥 系統執行發生嚴重錯誤: {e}")
