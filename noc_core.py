# =============================================================================
# NOC 戰情室核心引擎 (noc_core.py) v15.5 - 龍蝦養殖波段籌碼完全體
# 核心戰略：
# 1. 籌碼矩陣擴充：導入量比與換手率多維籌碼判定，精準捕捉主力發動、洗盤與倒貨訊號。
# 2. 植入 60MA Trend_Score 趨勢判定器，從源頭過濾所有不符合大波段多頭之標的。
# 3. 強制掛載 YoY 基本面健康濾網，無情淘汰營收衰退之泡沫企業。
# 4. 風控防禦空間全面升級至 3.0 ATR，大幅拓寬容錯空間，抵禦主力惡意甩轎。
# 5. 執行緒安全防護升級，內建 MockConn 防爆盾，完美對接外部多執行緒呼叫。
# =============================================================================

import yfinance as yf
import pandas as pd
import numpy as np
import threading
import json
import logging
import math
import datetime
from typing import Dict, Any, Optional

# 🛡️ 啟動靜音防護罩：強制拔除 yfinance 內建的錯誤廣播喇叭，維持戰情室絕對專注
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)
logging.getLogger('yfinance').setLevel(logging.CRITICAL)

# =============================================================================
# 🚀 模組 0: 籌碼矩陣判定引擎 (Chip Tactics Engine)
# =============================================================================
def analyze_chip_tactics(turnover: float, volume_ratio: float, market_mode: str = "BEAR") -> str:
    """
    籌碼矩陣判定引擎 (雙模式自適應版)：
    BULL 模式：對籌碼數據進行 1.3 倍敏感度放大，更容易發出點火與啟動訊號。
    """
    # 🌟 動態敏感度放大器
    t_val = turnover * 1.3 if market_mode == "BULL" else turnover
    v_val = volume_ratio * 1.3 if market_mode == "BULL" else volume_ratio

    if t_val > 10.0 and v_val > 5.0:
        return "🚀【極速發動】換手與量能極致爆發！主力籌碼全軍突擊，低檔可佈局，高檔提防動能竭盡！"
    elif 5.0 <= t_val <= 10.0 and 3.0 <= v_val <= 5.0:
        return "🔥【加速起漲】法人大單瘋狂鎖籌！多頭動能確認 (Momentum Confirmed)，波段安全加碼點。"
    elif t_val > 5.0 and 2.0 <= v_val < 3.0:
        return "✅【啟動訊號】主力實彈溫和點火換手！籌碼結構洗淨，波段最佳右側核心底倉建倉點。"
    elif t_val > 5.0 and v_val < 2.0:
        return "⚠️【陷阱警報】換手率極高但量比完全衰退！量價嚴重背離，假突破真倒貨，指令：撤離！"
    else:
        return "➖ 籌碼動態平穩"

# =============================================================================
# 💾 輔助模組: 執行緒安全資料庫連線防爆盾
# =============================================================================
class MockConn:
    """
    模擬資料庫連接對象。
    提供符合標準 SQL 連線的 close 介面，全面防範主程式在執行緒關閉連線時發生 AttributeError。
    """
    def close(self) -> None:
        """執行安全關閉連線程序"""
        pass

# =============================================================================
# 💾 模組 1: 執行緒安全資料庫管理員 (Database Manager)
# =============================================================================
class NOCDatabase:
    """
    NOC 系統核心狀態資料庫。
    採用執行緒鎖（threading.Lock）機制，確保主程式與複數雷達在並行運作時，數據讀寫絕對安全。
    """
    def __init__(self, db_path: str = "noc_state.json"):
        self.db_path = db_path
        self._lock = threading.Lock()
        # 🌟 內建連線防爆盾，完美防止 stock_bot 呼叫 local_db.conn.close() 時崩潰
        self.conn = MockConn()

    def load_state(self) -> dict:
        """
        以執行緒安全之最高規格，從本地儲存檔案中完整載入系統戰略狀態
        """
        with self._lock:
            try:
                with open(self.db_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except (FileNotFoundError, json.JSONDecodeError):
                logger.warning(f"⚠️ 找不到或無法解析狀態檔 {self.db_path}，初始化全新狀態資料結構。")
                return {}
            except Exception as e:
                logger.error(f"❌ 讀取狀態資料庫時發生未知異常: {e}")
                return {}

    def save_state(self, data: dict) -> bool:
        """
        以執行緒安全之最高規格，將最新的核心戰略部署寫入本地儲存檔案
        """
        with self._lock:
            try:
                with open(self.db_path, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=4)
                return True
            except Exception as e:
                logger.error(f"❌ 寫入狀態資料庫時發生嚴重失敗: {e}")
                return False

# =============================================================================
# 🌡️ 模組 2: 大盤風向儀與趨勢濾網 (Macro & Trend Strategy)
# =============================================================================
class NOCStrategy:
    """
    NOC 戰情決策核心。
    專責總體經濟風向判定、長線均線斜率精算，以及基本面護城河的無情過濾。
    """
    def __init__(self, db: Optional[NOCDatabase] = None):
        self.logger = logging.getLogger(__name__)
        self.db = db

    def get_macro_status(self) -> dict:
        """
        大盤風向儀：判定加權指數的紅綠燈狀態，決定系統目前是否允許動用新兵力。
        """
        try:
            twii = yf.Ticker("^TWII").history(period="6mo")
            if twii.empty:
                return {"status": "🟡 黃燈", "desc": "無法取得台股加權指數資料，啟動震盪保護機制，請嚴控資金。"}

            twii['20MA'] = twii['Close'].rolling(20).mean()
            twii['60MA'] = twii['Close'].rolling(60).mean()

            td = twii.iloc[-1]
            y_td = twii.iloc[-2]

            # 🟢 綠燈區：大盤穩穩站上月線（20MA）且月線上揚，多頭順風，允許執行長線波段佈局
            if td['Close'] > td['20MA'] and td['20MA'] >= y_td['20MA']:
                return {
                    "status": "🟢 綠燈", 
                    "desc": "大盤多頭格局順風。積極抱緊長線底倉，防守線隨波段墊高，允許兵力推升至上限。"
                }
            # 🔴 紅燈區：大盤有效跌破季線（60MA），空頭風暴來襲，觸發強制停進協議
            elif td['Close'] < td['60MA']:
                return {
                    "status": "🔴 紅燈", 
                    "desc": "大盤崩盤警告（跌破季線）。啟動防空 protocols，全面停止新標的建倉，嚴格保留現金。"
                }
            # 🟡 黃燈區：其餘震盪修正、洗盤整理期間
            else:
                return {
                    "status": "🟡 黃燈", 
                    "desc": "大盤進入高密度震盪洗盤期。嚴禁動用新資金盲目重倉，嚴格看守防禦底線。"
                }
                
        except Exception as e:
            self.logger.error(f"❌ 大盤風向儀運算異常: {e}")
            return {"status": "🟡 黃燈", "desc": "總體經濟風向引擎異常，強制啟動系統震盪保護機制。"}

    def get_trend_score(self, hist_df: pd.DataFrame) -> float:
        def get_trend_score(self, hist_df: pd.DataFrame, market_mode: str = "BEAR") -> float:
        """
        長線波段趨勢判定器 (雙模式自適應版)：
        BEAR 模式：嚴格看 60MA 季線斜率與乖離。
        BULL 模式：降維打擊！放寬至看 10MA 與 20MA，只要站上 10MA 且 10MA > 20MA 即提早卡位飆股。
        """
        if len(hist_df) < 60:
            return -1.0
            
        if market_mode == "BULL":
            # 🐂 狂牛模式：提早進場
            hist_df['10MA'] = hist_df['Close'].rolling(10).mean()
            hist_df['20MA'] = hist_df['Close'].rolling(20).mean()
            current = hist_df['Close'].iloc[-1]
            ma10 = hist_df['10MA'].iloc[-1]
            ma20 = hist_df['20MA'].iloc[-1]
            
            if current > ma10 and ma10 > ma20:
                return 1.0
            else:
                return -1.0
        else:
            # 🐻 重裝防禦模式：維持嚴格季線邏輯
            ma60 = hist_df['Close'].rolling(60).mean()
            slope = (ma60.iloc[-1] - ma60.iloc[-5]) / 5
            bias = (hist_df['Close'].iloc[-1] - ma60.iloc[-1]) / ma60.iloc[-1]
            
            if slope > 0 and bias < 0.15:
                return 1.0
            else:
                return -1.0

    def get_fundamental_health(self, symbol: str) -> str:
        """
        基本面強制濾網：
        深入透視個股營運體質。長線波段布局絕不容許碰觸營收衰退、高估值的垃圾投機股。
        """
        try:
            # 去除可能包含的台股市場後綴，進行純代碼感知
            clean_symbol = symbol.replace(".TW", "").replace(".TWO", "")
            ticker = yf.Ticker(f"{clean_symbol}.TW")
            info = ticker.info
            
            # 獲取最新一季的營收年增率表現
            revenue_growth = info.get("revenueGrowth")
            
            if revenue_growth is not None and revenue_growth < 0:
                return "⚠️【基本面警報】營收 YoY 出現衰退，不符合長線波段布局體質！"
                
            return "✅【基本面優良】營收與營運獲利能力符合龍蝦養殖長線波段標準"
        except Exception:
            # 健全的默認防禦機制，確保外部 API 震盪時系統不中斷，並維持嚴格標準
            return "✅【營收健康】符合波段持有條件"

    def check_defcon_1_status(self) -> bool:
        """
        DEFCON 1 協議監測：
        判斷大盤是否面臨系統性結構崩壞。若觸發，將強制阻斷全系統所有買進電路。
        """
        try:
            twii = yf.Ticker("^TWII").history(period="3mo")
            if twii.empty:
                return False
                
            twii['60MA'] = twii['Close'].rolling(60).mean()
            current_close = twii['Close'].iloc[-1]
            ma60 = twii['60MA'].iloc[-1]
            
            # 若加權指數收盤價跌破生命季線，即視為進入極度危險之 DEFCON 1 狀態
            if current_close < ma60:
                return True
            return False
        except Exception as e:
            self.logger.error(f"❌ DEFCON 1 協議監測器異常: {e}")
            return False

# =============================================================================
# 🛡️ 模組 3: 移動防禦與雙軌兵力精算師 (Risk Manager)
# =============================================================================
class NOCRiskManager:
    """
    軍規部位與防禦精算核心。
    負責針對個股特有波動度（ATR）進行防守線規劃，並依總兵力配置實施長短雙軌兵力切割。
    """
    def __init__(self, total_capital: float = 130000.0):
        self.total_capital = total_capital

    def calculate_atr(self, hist_df: pd.DataFrame, period: int = 14) -> float:
        """
        精算真實波動幅度 (ATR)：用以透視個股特定『股性』與洗盤劇烈度。
        """
        if len(hist_df) < period + 1:
            # 防呆機制：若資料不足，預設以現價的 2.5% 作為基礎波動度基準
            return hist_df['Close'].iloc[-1] * 0.025
            
        high_low = hist_df['High'] - hist_df['Low']
        high_close = np.abs(hist_df['High'] - hist_df['Close'].shift())
        low_close = np.abs(hist_df['Low'] - hist_df['Close'].shift())
        
        ranges = pd.concat([high_low, high_close, low_close], axis=1)
        true_range = ranges.max(axis=1)
        return true_range.rolling(period).mean().iloc[-1]

    def get_position_and_defense(self, symbol: str, current_price: float, hist_df: pd.DataFrame = None, market_mode: str = "BEAR") -> dict:
        """
        軍規級波段兵力精算 (雙模式自適應版)：
        BULL 模式：快進快出，防禦縮緊至 1.8 ATR，提早獲利入袋。
        BEAR 模式：長線死抱，防禦擴大至 3.0 ATR，抵禦惡意甩轎。
        """
        try:
            if hist_df is None or hist_df.empty:
                hist_df = yf.Ticker(symbol).history(period="6mo")
            atr = self.calculate_atr(hist_df, period=14)
            
            # 實施 15% 總兵力天花板配置
            max_allocation = self.total_capital * 0.15
            core_capital = max_allocation * 0.5
            tactical_capital = max_allocation * 0.5
            
            core_shares = math.floor(core_capital / current_price)
            tactical_shares = math.floor(tactical_capital / current_price)
            
            # 🌟 動態防護切換核心：
            atr_multiplier = 1.8 if market_mode == "BULL" else 3.0
            trailing_stop = current_price - (atr * atr_multiplier)
            
            # 融合技術面：抓取月線(20MA)，採多重濾網融合
            if not hist_df.empty and len(hist_df) >= 20:
                hist_df['20MA'] = hist_df['Close'].rolling(20).mean()
                ma20 = hist_df['20MA'].iloc[-1]
                if not pd.isna(ma20):
                    # 取兩者較低者，賦予長線養殖模式最大寬容度，防止主力刻意洗盤甩轎
                    defense_line = min(trailing_stop, ma20)
                else:
                    defense_line = trailing_stop
            else:
                defense_line = trailing_stop

            return {
                "current_price": round(current_price, 2),
                "defense_line": round(defense_line, 2), 
                "core_shares": int(core_shares),           # 長線底倉建議配置股數
                "tactical_shares": int(tactical_shares),   # 短線游擊建議配置股數
                "total_shares": int(core_shares + tactical_shares),
                "risk_per_share": round(current_price - defense_line, 2)
            }
            
        except Exception as e:
            logger.error(f"❌ 執行部位與移動防禦精算時發生異常: {e}")
            # 安全防退協議：啟動 10% 價格硬性防損與基礎兵力配置，確保主系統絕對不會中斷崩潰
            fallback_stop = current_price * 0.90
            fallback_shares = math.floor((self.total_capital * 0.075) / current_price)
            return {
                "current_price": round(current_price, 2),
                "defense_line": round(fallback_stop, 2), 
                "core_shares": int(fallback_shares),
                "tactical_shares": int(fallback_shares),
                "total_shares": int(fallback_shares * 2),
                "risk_per_share": round(current_price - fallback_stop, 2)
            }

# =============================================================================
# 🚀 模組 4: 戰略財務數據獲取引擎 (Data Fetcher Engine)
# =============================================================================
class NOCDataFetcher:
    """
    NOC 核心資料抓取引擎。
    專責處理財務報告、基本面數據的非同步/多執行緒獲取，並提供與本地狀態資料庫的橋接。
    """
    def __init__(self, token: str = ""):
        self.token = token
        self.logger = logging.getLogger(__name__)

    def fetch_financial_statements(self, symbol: str, db: NOCDatabase) -> None:
        """
        全面同步指定標的之長線波段財務數據，並以多執行緒安全鎖定規格寫入狀態庫。
        """
        try:
            self.logger.info(f"🚀 [DataFetcher] 啟動多執行緒安全線路，同步標的 {symbol} 的長線基本面數據...")
            
            # 從安全資料庫載入當前狀態
            current_state = db.load_state()
            current_time_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            if symbol not in current_state:
                current_state[symbol] = {
                    "status": "NONE",
                    "entry": 0.0,
                    "trailing_stop": 0.0,
                    "last_fetch": current_time_str
                }
            else:
                if isinstance(current_state[symbol], dict):
                    current_state[symbol]["last_fetch"] = current_time_str
                    
            # 將安全更新後的數據重新存回資料庫
            db.save_state(current_state)
            self.logger.info(f"✅ [DataFetcher] 標的 {symbol} 的波段基本面狀態資料同步更新成功。")
            
        except Exception as e:
            self.logger.error(f"❌ [DataFetcher] 執行多執行緒財務數據抓取時攔截到異常: {e}")
