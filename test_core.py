import os
import sys
from dotenv import load_dotenv

# 🛡️ 裝甲自檢程序：先列出當前目錄的所有檔案，確認 GitHub 真的有抓到最新版
print("🔍 [裝甲自檢] 正在掃描當前雲端機房的檔案清單...")
print(os.listdir('.'))

# 確保 noc_core 存在於目錄中
if not os.path.exists('noc_core.py'):
    print("❌ 致命錯誤：找不到 noc_core.py！請確認檔案已成功 push 到 GitHub。")
    sys.exit(1)

# 🌟 正式載入模組
from noc_core import NOCDatabase, NOCDataFetcher, NOCStrategy

# 載入環境變數 (確保 FINMIND_TOKEN 有抓到)
load_dotenv()
FINMIND_TOKEN = os.getenv("FINMIND_TOKEN")
if not FINMIND_TOKEN:
    print("⚠️ 警告：未偵測到 FINMIND_TOKEN，可能導致抓取受限。")

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
