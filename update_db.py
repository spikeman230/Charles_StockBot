# update_db.py
import datetime
import os
import json
import time
from pathlib import Path
from dotenv import load_dotenv

# 引入 NOC 核心防禦與後勤模組
from noc_core import NOCDatabase, NOCDataFetcher

# 🌟 引入主程式的 Trello 模組與設定 
# (⚠️ 注意：如果您的主程式不叫 stock_bot.py，請修改此處的檔名)
from stock_bot import fetch_trello_deployment, cfg

load_dotenv()
FINMIND_TOKEN = os.getenv("FINMIND_TOKEN")

if __name__ == "__main__":
    print("🚚 開始執行 NOC 盤後戰情資料庫補給作業...")
    db = NOCDatabase()
    fetcher = NOCDataFetcher(token=FINMIND_TOKEN)
    
    # 預設抓取近 3 天的資料進行增量更新 (SQLite 會自動過濾已存在的重複日期)
    start_date = (datetime.datetime.now() - datetime.timedelta(days=3)).strftime("%Y-%m-%d")
    
    try:
        # 1. 優先更新大盤與防空警報數據
        print("1️⃣ 正在更新大盤指數與外資期貨空單數據...")
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
        
        print(f"🎯 成功剔除重複目標！最終鎖定 {len(target_stocks)} 檔股票，準備大規模補給！")

        # 3. 執行各股籌碼與 K 線抓取
        for sym in target_stocks:
            fetcher.fetch_and_store_stock_data(sym, start_date, db)
            # 稍微暫停 0.5 秒，避免連續向 FinMind 發送過多請求被鎖 IP
            time.sleep(0.5) 
            
        print("✅ 戰情資料庫補給完畢！雷達彈藥充足！")
        
    except Exception as e:
        print(f"⚠️ 補給過程發生致命錯誤: {e}")
