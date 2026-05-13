import os
from dotenv import load_dotenv
from noc_core import NOCDatabase,NOCDataFetcher,NOCStrategy

# 載入環境變數 (確保 FINMIND_TOKEN 有抓到)
load_dotenv()
FINMIND_TOKEN = os.getenv("FINMIND_TOKEN")

print("🔌 [測試 1] 啟動本地軍火庫...")
db = NOCDatabase()
print("✅ 資料庫 noc_warroom.db 已連線並初始化資料表。")

print("\n📡 [測試 2] 測試 FinMind 數據管線 (以台積電 2330 為例)...")
fetcher = NOCDataFetcher(token=FINMIND_TOKEN)
try:
    # 抓取近幾個月的資料測試寫入
    fetcher.fetch_and_store_stock_data("2330", "2024-01-01", db)
    print("✅ 2330 K線與融資券資料已成功寫入 SQLite！")
except Exception as e:
    print(f"❌ 數據抓取或寫入失敗，請檢查網路或 FinMind Token: {e}")

print("\n🧠 [測試 3] 測試戰略引擎 (籌碼透視)...")
strategy = NOCStrategy(db)
result = strategy.analyze_stock_opportunity("2330")
print(f"👉 戰情室籌碼判讀結果：【 {result} 】")
