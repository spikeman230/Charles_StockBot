# init_db.py
import datetime
import os
import json
import time
from pathlib import Path
from dotenv import load_dotenv

# 引入 NOC 核心防禦與後勤模組
from noc_core import NOCDatabase, NOCDataFetcher

# 🌟 引入主程式的 Trello 模組與設定 
# (⚠️ 注意：如果您的主程式不叫 stock_bot_0506.py，請修改此處的檔名)
from stock_bot import fetch_trello_deployment, cfg

load_dotenv()
FINMIND_TOKEN = os.getenv("FINMIND_TOKEN")

if __name__ == "__main__":
    print("🚀 啟動 NOC 戰情室「建庫大補丸」歷史資料載入作業...")
    db = NOCDatabase()
    fetcher = NOCDataFetcher(token=FINMIND_TOKEN)
    
    # 🌟 建庫大補丸的關鍵差異：抓取近 400 天的歷史資料 (涵蓋年線與籌碼長期動向)
    start_date = (datetime.datetime.now() - datetime.timedelta(days=400)).strftime("%Y-%m-%d")
    
    try:
        # 1. 優先下載大盤與防空警報歷史數據
        print("1️⃣ 正在下載大盤與防空歷史數據 (近400天)...")
        fetcher.fetch_market_health_data(start_date, db)
        
        # 2. 🌟 智能尋標系統：自動收集 Trello 與各級雷達檔案中的名單
        print("2️⃣ 正在從 Trello 與各級雷達檔案中提取實戰名單...")
        
        # 🛡️ 關鍵防線：使用 set() 建立集合，天生不允許重複！
        all_symbols = set() 
        
        # 【情報來源 A：Trello 實戰看板】
        try:
            TRELLO_DICT, TRELLO_PORTFOLIO = fetch_trello_deployment()
            if TRELLO_PORTFOLIO:
                for sym in TRELLO_PORTFOLIO.keys(): 
                    all_symbols.add(sym)
            if TRELLO_DICT:
                for stocks in TRELLO_DICT.values():
                    for sym in stocks.keys(): 
                        all_symbols.add(sym)
        except Exception as e:
            print(f"⚠️ 讀取 Trello 名單時發生錯誤: {e}")
                
        # 【情報來源 B：三大 JSON 雷達檔案】 (包含雷達、閃電、游擊隊)
        json_files = [cfg.RADAR_FILE, cfg.LIGHTNING_FILE, cfg.GUERRILLA_FILE]
        for fname in json_files: 
            if Path(fname).exists():
                try:
                    with open(fname, "r", encoding="utf-8") as f: 
                        data = json.load(f)
                        for sym in data.keys(): 
                            all_symbols.add(sym) # 把代號丟進集合 (重複的會自動無視)
                except Exception as e:
                    print(f"⚠️ 讀取 {fname} 時發生錯誤: {e}")

        # 轉換回標準陣列，準備發給補給兵
        target_stocks = list(all_symbols)
        
        # 確保不會抓到大盤(TAIEX)或空白代號 (過濾防呆)
        target_stocks = [s for s in target_stocks if s and not s.isalpha()]
        
        print(f"🎯 成功剔除重複目標！最終鎖定 {len(target_stocks)} 檔股票，準備進行 400 天歷史大補給！")
        print("⚠️ 因為是載入超過百檔的長天期歷史數據，預計耗時 10-20 分鐘，請耐心等候...")

        # 3. 執行各股籌碼與 K 線歷史抓取
        for i, sym in enumerate(target_stocks, 1):
            print(f"[{i}/{len(target_stocks)}] ", end="")
            fetcher.fetch_and_store_stock_data(sym, start_date, db)
            # 🌟 為了保護 FinMind API 連線不被強制中斷，這裡設定 1 秒的間隔
            time.sleep(1.0) 
            
        print("✅ 歷史戰情資料庫 (SQLite) 初始建置與灌水完成！")
        
    except Exception as e:
        print(f"⚠️ 建庫過程發生致命錯誤: {e}")
