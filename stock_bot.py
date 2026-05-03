# =============================================================================
# NOC 終極戰情室 v12.0 - 企業級效能與實戰體驗整合版
# 優化項目：背景非同步 Trello、防爬蟲隨機延遲、GitHub Secrets 友善、全 Emoji 支援
# =============================================================================

import yfinance as yf
import requests
import os
import datetime
import pandas as pd
import numpy as np
import csv
import json
import math
import re
import mplfinance as mpf
import smtplib
import sys
import logging
import time
import random
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field, asdict
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from dotenv import load_dotenv
from typing import Optional, Dict, Tuple, Any
from pathlib import Path

# =============================================================================
# === 0. 初始化：載入環境變數 & 日誌系統 ===
# =============================================================================
# load_dotenv() 幫助本地端讀取 .env，若在 GitHub Actions 則由 OS 環境變數(Secrets)接手
load_dotenv()

LOG_FILE = "noc_system.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(funcName)s - %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# =============================================================================
# === 1. 機密環境變數 (直接對接 GitHub Secrets) ===
# =============================================================================
TG_TOKEN        = os.getenv("TG_TOKEN")
TG_CHAT_ID      = os.getenv("TG_CHAT_ID")
EMAIL_USER      = os.getenv("EMAIL_USER")
EMAIL_PASS      = os.getenv("EMAIL_PASS")
EMAIL_TO        = os.getenv("EMAIL_TO")
FINMIND_TOKEN   = os.getenv("FINMIND_TOKEN")
TRELLO_KEY      = os.getenv("TRELLO_KEY")
TRELLO_TOKEN    = os.getenv("TRELLO_TOKEN")
TRELLO_BOARD_ID = os.getenv("TRELLO_BOARD_ID")

# =============================================================================
# === 1.1 量化風控常數 (集中管理) ===
# =============================================================================
class Config:
    TOTAL_CAPITAL      : float = float(os.getenv("TOTAL_CAPITAL", "1000000"))
    RISK_PER_TRADE     : float = float(os.getenv("RISK_PER_TRADE", "0.02"))
    ATR_MULTIPLIER     : float = float(os.getenv("ATR_MULTIPLIER", "2.0"))
    YOY_EXPLOSION_PCT  : float = float(os.getenv("YOY_EXPLOSION_PCT", "50.0"))
    PE_LIMIT           : float = float(os.getenv("PE_LIMIT", "40.0"))
    SILENT_MODE        : bool  = os.getenv("SILENT_MODE", "false").lower() == "true"
    CACHE_TTL_MINUTES  : int   = int(os.getenv("CACHE_TTL_MINUTES", "30"))
    CACHE_MAX_ITEMS    : int   = int(os.getenv("CACHE_MAX_ITEMS", "200"))
    MAX_WORKERS        : int   = int(os.getenv("MAX_WORKERS", "6")) # 稍微降低以防 Rate Limit
    STATE_FILE         : str   = "noc_state.json"
    LOG_FILE_CSV       : str   = "noc_trading_log.csv"
    RADAR_FILE         : str   = "radar_targets.json"
    LIGHTNING_FILE     : str   = "lightning_targets.json"

cfg = Config()

# =============================================================================
# === 1.2 強型別狀態管理 (改用 dataclasses) ===
# =============================================================================
@dataclass
class StockState:
    """單支股票的持倉狀態"""
    status         : str   = "NONE"   # NONE / HOLD / REAL_HOLD / REAL_HOLD_ETF
    entry          : float = 0.0
    trailing_stop  : float = 0.0

    def to_dict(self) -> dict:
        return asdict(self)

    @staticmethod
    def from_dict(d: dict) -> "StockState":
        return StockState(
            status        = d.get("status", "NONE"),
            entry         = float(d.get("entry", 0.0)),
            trailing_stop = float(d.get("trailing_stop", 0.0)),
        )

# =============================================================================
# === 1.3 快取管理器 (支援 TTL 過期 + 容量上限) ===
# =============================================================================
@dataclass
class CacheEntry:
    data      : Any
    timestamp : datetime.datetime = field(default_factory=datetime.datetime.utcnow)

class DataCacheManager:
    """帶 TTL 過期機制的記憶體快取管理器"""
    def __init__(self, ttl_minutes: int = 30, max_items: int = 200):
        self._cache   : Dict[str, CacheEntry] = {}
        self._ttl     = datetime.timedelta(minutes=ttl_minutes)
        self._max     = max_items

    def get(self, key: str) -> Optional[Any]:
        entry = self._cache.get(key)
        if entry is None: return None
        if datetime.datetime.utcnow() - entry.timestamp > self._ttl:
            del self._cache[key]
            logger.debug(f"快取過期，移除: {key}")
            return None
        return entry.data

    def set(self, key: str, data: Any) -> None:
        if len(self._cache) >= self._max:
            evict_count = max(1, self._max // 10)
            oldest_keys = sorted(self._cache, key=lambda k: self._cache[k].timestamp)[:evict_count]
            for k in oldest_keys: del self._cache[k]
            logger.info(f"快取已滿，清除 {evict_count} 筆最舊紀錄")
        self._cache[key] = CacheEntry(data=data)

    def clear(self) -> None:
        self._cache.clear()

DATA_CACHE = DataCacheManager(ttl_minutes=cfg.CACHE_TTL_MINUTES, max_items=cfg.CACHE_MAX_ITEMS)

# =============================================================================
# === 2. ETF 策略判定引擎 (恢復 Emoji) ===
# =============================================================================
_ETF_DIV_KEYS = ["高股息","優息","0056","00878","00919","00929","00915","00713","00939","00940","00936"]
_ETF_MKT_KEYS = ["0050","006208","市值","AAPL","NVDA","TSM","00881","科技","半導體","5G","00891","00892","009816"]

def get_etf_strategy(symbol: str, name: str) -> Tuple[str, float, str]:
    if any(k in name or k in symbol for k in _ETF_DIV_KEYS): return "💰高股息", 5.0, "控管殖利率 (5%乖離預警)"
    elif any(k in name or k in symbol for k in _ETF_MKT_KEYS): return "🚀市值/主題型", 10.0, "成長動能區 (10%乖離預警)"
    return "🔸一般型", 8.0, "趨勢防禦區 (8%乖離預警)"

# =============================================================================
# === 3. 交易日判斷 ===
# =============================================================================
def is_trading_day(curr_date: datetime.date) -> bool:
    try:
        tsm = yf.Ticker("2330.TW").history(period="1d")
        if tsm.empty: return False
        return tsm.index[-1].date() == curr_date
    except Exception as e:
        logger.warning(f"交易日 API 異常，降級為工作日判斷: {e}")
        return curr_date.weekday() < 5

# =============================================================================
# === 4. Trello 整合模組 (背景非同步化) ===
# =============================================================================
def _trello_params(**extra) -> dict:
    return {"key": TRELLO_KEY, "token": TRELLO_TOKEN, **extra}

def _trello_available() -> bool:
    return all([TRELLO_KEY, TRELLO_TOKEN, TRELLO_BOARD_ID])

def update_trello_system_status(status_msg: str, color: str = "🟢") -> None:
    """實際與 Trello 溝通的函數 (供背景執行緒呼叫)"""
    if not _trello_available(): return
    url = f"https://api.trello.com/1/boards/{TRELLO_BOARD_ID}/lists"
    try:
        res = requests.get(url, params=_trello_params(cards="open"), timeout=10)
        res.raise_for_status()
        lists_data = res.json()
        if not lists_data: return

        first_list_id = lists_data[0]["id"]
        status_card_id = None

        for lst in lists_data:
            for card in lst.get("cards", []):
                if "NOC 系統狀態" in card["name"]:
                    status_card_id = card["id"]
                    break
            if status_card_id: break

        tw_tz   = datetime.timezone(datetime.timedelta(hours=8))
        date_str = datetime.datetime.now(tw_tz).strftime("%m/%d")
        new_name = f"{color} NOC 系統狀態：{status_msg} ({date_str})"

        if status_card_id:
            requests.put(f"https://api.trello.com/1/cards/{status_card_id}", params=_trello_params(name=new_name), timeout=10)
        else:
            requests.post("https://api.trello.com/1/cards", params=_trello_params(idList=first_list_id, name=new_name, pos="top"), timeout=10)
        logger.info(f"Trello 狀態已更新：{new_name}")
    except Exception as e:
        logger.error(f"Trello 狀態更新失敗: {e}")

def update_trello_system_status_bg(status_msg: str, color: str = "🟢") -> None:
    """🔥 背景非同步啟動：避免競爭條件與阻塞主程式"""
    threading.Thread(target=update_trello_system_status, args=(status_msg, color), daemon=True).start()

def _parse_card_to_stock(card: dict) -> Tuple[str, dict]:
    raw_name = card["name"].strip()
    ticker_match = re.match(r"^[A-Za-z0-9.]+", raw_name)
    symbol    = ticker_match.group() if ticker_match else raw_name
    name_part = raw_name[len(symbol):].strip() if ticker_match else raw_name
    name      = re.sub(r"\(.*?\)", "", name_part).strip() or symbol

    title_tip_match = re.search(r"\((.*?)\)", name_part)
    trello_tip = title_tip_match.group(1) if title_tip_match else card.get("desc", "").strip()
    return symbol, {"name": name, "trello_tip": trello_tip}

def _parse_card_to_portfolio(card: dict) -> Tuple[str, dict]:
    raw_name     = card["name"].strip()
    ticker_match = re.match(r"^[A-Za-z0-9.]+", raw_name)
    symbol    = ticker_match.group() if ticker_match else raw_name
    name_part = raw_name[len(symbol):].strip() if ticker_match else raw_name
    name      = re.sub(r"\(.*?\)", "", name_part).strip() or symbol

    desc = card.get("desc", "")
    buy_price, shares = 0.0, 1000
    price_match  = re.search(r"成本[：:]\s*([0-9.]+)", desc)
    shares_match = re.search(r"股數[：:]\s*([0-9]+)", desc)
    if price_match: buy_price = float(price_match.group(1))
    if shares_match: shares = int(shares_match.group(1))

    return symbol, {"name": name, "buy_price": buy_price, "shares": shares, "trello_tip": desc}

def fetch_trello_deployment() -> Tuple[Optional[dict], Optional[dict]]:
    if not _trello_available():
        logger.warning("未偵測到 Trello 金鑰，將啟用靜態預設部署。")
        return None, None

    logger.info("正在連線 Trello 戰情看板...")
    url = f"https://api.trello.com/1/boards/{TRELLO_BOARD_ID}/lists"
    try:
        response = requests.get(url, params=_trello_params(cards="open"), timeout=10)
        response.raise_for_status()
        lists_data = response.json()

        trello_dict  : Dict[str, dict] = {}
        my_portfolio : Dict[str, dict] = {}

        for lst in lists_data:
            list_name = lst["name"]
            is_portfolio_list = "庫存" in list_name or "庫藏" in list_name
            for card in lst.get("cards", []):
                if "NOC 系統狀態" in card["name"]: continue
                if is_portfolio_list:
                    sym, info = _parse_card_to_portfolio(card)
                    my_portfolio[sym] = info
                else:
                    sym, info = _parse_card_to_stock(card)
                    trello_dict.setdefault(list_name, {})[sym] = info

        logger.info("成功從 Trello 載入最新部署！")
        return trello_dict, my_portfolio
    except Exception as e:
        logger.error(f"Trello 連線/解析失敗: {e}")
        return None, None

# =============================================================================
# === 5. 本地狀態記憶庫 ===
# =============================================================================
def load_state() -> Dict[str, StockState]:
    if not Path(cfg.STATE_FILE).exists(): return {}
    try:
        with open(cfg.STATE_FILE, "r", encoding="utf-8") as f:
            raw = json.load(f)
        return {sym: StockState.from_dict(d) for sym, d in raw.items()}
    except Exception as e:
        logger.error(f"狀態檔案讀取失敗: {e}")
        return {}

def save_state(state: Dict[str, StockState]) -> None:
    try:
        with open(cfg.STATE_FILE, "w", encoding="utf-8") as f:
            json.dump({sym: s.to_dict() for sym, s in state.items()}, f, ensure_ascii=False, indent=4)
    except OSError as e:
        logger.error(f"狀態檔案寫入失敗: {e}")

def write_noc_log(date, symbol, name, close_price, rsi, vol_status, status, predict, chip_signal, alert) -> None:
    log_exists = Path(cfg.LOG_FILE_CSV).exists()
    try:
        with open(cfg.LOG_FILE_CSV, mode="a", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            if not log_exists: writer.writerow(["日期","代號","名稱","收盤價","RSI","量能狀態","趨勢狀態","戰場預判","籌碼訊號","行動指令"])
            writer.writerow([date, symbol, name, f"{close_price:.2f}", f"{rsi:.2f}", vol_status, status, predict, chip_signal, alert])
    except Exception as e:
        logger.error(f"CSV 日誌寫入失敗: {e}")

# =============================================================================
# === 6. 環境感知：大盤、營收、估值 ===
# =============================================================================
def get_market_regime() -> Tuple[bool, str]:
    try:
        twii = yf.Ticker("^TWII").history(period="1mo")
        if twii.empty: raise ValueError("TWII 資料為空")
        twii["20MA"] = twii["Close"].rolling(20).mean()
        is_bull = twii["Close"].iloc[-1] > twii["20MA"].iloc[-1]
        msg = "🟢 多頭格局 (站上月線)" if is_bull else "🔴 空頭警戒 (跌破月線)"
        return is_bull, msg
    except Exception as e:
        logger.warning(f"大盤狀態取得失敗: {e}")
        return True, "🟡 大盤狀態未知"

def get_revenue_yoy(symbol: str):
    if not FINMIND_TOKEN: return "N/A"
    match = re.search(r"\d+", symbol)
    if not match: return "N/A"
    try:
        url    = "https://api.finmindtrade.com/api/v4/data"
        params = {"dataset": "TaiwanStockMonthRevenue", "data_id": match.group(), "start_date": (datetime.datetime.now() - datetime.timedelta(days=400)).strftime("%Y-%m-%d"), "token": FINMIND_TOKEN}
        r = requests.get(url, params=params, timeout=10)
        r.raise_for_status()
        data = r.json()
        if data.get("msg") == "success" and data.get("data"):
            df     = pd.DataFrame(data["data"])
            latest = df.iloc[-1]
            prev   = df[(df["revenue_year"] == latest["revenue_year"] - 1) & (df["revenue_month"] == latest["revenue_month"])]
            if not prev.empty and prev.iloc[-1]["revenue"] > 0:
                return float((latest["revenue"] - prev.iloc[-1]["revenue"]) / prev.iloc[-1]["revenue"] * 100)
    except Exception as e: pass
    return "N/A"

def get_pe_ratio(symbol: str):
    try:
        info = yf.Ticker(symbol).info
        pe   = info.get("trailingPE") or info.get("forwardPE")
        return pe if pe else "N/A"
    except: return "N/A"

# =============================================================================
# === 7. FinMind 籌碼分析 (恢復 Emoji) ===
# =============================================================================
def get_finmind_chip_data(symbol: str, start_date_str: str) -> pd.DataFrame:
    if not FINMIND_TOKEN: return pd.DataFrame()
    match = re.search(r"\d+", symbol)
    if not match: return pd.DataFrame()
    try:
        url    = "https://api.finmindtrade.com/api/v4/data"
        params = {"dataset": "TaiwanStockInstitutionalInvestorsBuySell", "data_id": match.group(), "start_date": start_date_str, "token": FINMIND_TOKEN}
        r = requests.get(url, params=params, timeout=10)
        r.raise_for_status()
        data = r.json()

        if data.get("msg") == "success" and data.get("data"):
            df            = pd.DataFrame(data["data"])
            df["net_buy"] = df["buy"] - df["sell"]
            df["type"]    = "Other"
            df.loc[df["name"].str.contains("外資"), "type"] = "Foreign_Inv"
            df.loc[df["name"].str.contains("投信"), "type"] = "Trust_Inv"
            df.loc[df["name"].str.contains("自營商"), "type"] = "Dealer_Inv"

            pivot_df = df.groupby(["date", "type"])["net_buy"].sum().unstack(fill_value=0).reset_index()
            for col in ["Foreign_Inv", "Trust_Inv", "Dealer_Inv"]:
                if col not in pivot_df.columns: pivot_df[col] = 0

            pivot_df["Date"] = pd.to_datetime(pivot_df["date"]).dt.date
            pivot_df.set_index("Date", inplace=True)
            return pivot_df[["Foreign_Inv", "Trust_Inv", "Dealer_Inv"]]
    except: pass
    return pd.DataFrame()

def calculate_chip_signals(hist: pd.DataFrame) -> pd.DataFrame:
    hist["Chip_Status"] = "無資料"
    hist["Trust_Streak"] = 0
    required = {"Foreign_Inv", "Trust_Inv", "Dealer_Inv"}
    if not required.issubset(hist.columns): return hist

    hist["Total_Institutional"] = hist["Foreign_Inv"] + hist["Trust_Inv"] + hist["Dealer_Inv"]
    hist["Signal_CoBuy"]        = (hist["Foreign_Inv"] > 0) & (hist["Trust_Inv"] > 0)
    hist["Signal_Trust_Trend"]  = ((hist["Trust_Inv"] > 0).astype(int).rolling(5).sum() >= 4) & (hist["Trust_Inv"] > 0)

    trust_dir            = np.sign(hist["Trust_Inv"])
    hist["Trust_Streak"] = trust_dir.groupby((trust_dir != trust_dir.shift()).cumsum()).cumsum()

    conds   = [hist["Signal_CoBuy"], hist["Signal_Trust_Trend"], hist["Total_Institutional"] > 0]
    choices = ["🤝 土洋齊買", "🏦 投信作帳", "📈 法人偏多"]
    hist["Chip_Status"] = np.select(conds, choices, default="➖ 中性/偏空")
    return hist

# =============================================================================
# === 8. 核心資料獲取引擎 ===
# =============================================================================
def get_stock_data(symbol: str, name: str) -> Optional[pd.DataFrame]:
    cached = DATA_CACHE.get(symbol)
    if cached is not None: return cached

    try:
        stock = yf.Ticker(symbol)
        hist  = stock.history(period="8mo")
        hist  = hist.dropna(subset=["Close"])

        if len(hist) < 40: return None

        hist["Date_Key"] = hist.index.date
        if FINMIND_TOKEN and (".TW" in symbol or ".TWO" in symbol):
            chip_df = get_finmind_chip_data(symbol, (datetime.datetime.now() - datetime.timedelta(days=200)).strftime("%Y-%m-%d"))
            if not chip_df.empty:
                hist = hist.merge(chip_df, left_on="Date_Key", right_index=True, how="left")
                hist = hist.fillna(0)

        hist = calculate_chip_signals(hist)

        curr_hour = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).hour
        vol_mult = {10: 4.5, 12: 1.5, 13: 1.1}.get(curr_hour, 1.0)

        hist["Est_Volume"] = hist["Volume"].copy()
        if len(hist) > 0:
            hist.iloc[-1, hist.columns.get_loc("Est_Volume")] = hist["Volume"].iloc[-1] * vol_mult

        hist["5MA"]   = hist["Close"].rolling(5).mean()
        hist["20MA"]  = hist["Close"].rolling(20).mean()
        hist["25MA"]  = hist["Close"].rolling(25).mean()
        hist["5VMA"]  = hist["Est_Volume"].rolling(5).mean()
        hist["60VMA"] = hist["Volume"].rolling(60).mean()

        cond_trend    = hist["25MA"] > hist["25MA"].shift(3)
        cond_vol_mom  = hist["5VMA"] > hist["60VMA"]
        cond_pullback = ((hist["Low"] <= hist["25MA"] * 1.015) & (hist["Close"] >= hist["25MA"] * 0.985))
        cond_shrink       = hist["Est_Volume"] < hist["5VMA"]
        hist["Signal_2560"] = cond_trend & cond_vol_mom & cond_pullback & cond_shrink

        hist["High_60"]        = hist["High"].rolling(window=60, min_periods=20).max()
        hist["Low_60"]         = hist["Low"].rolling(window=60, min_periods=20).min()
        price_range            = (hist["High_60"] - hist["Low_60"]).replace(0, np.nan)
        hist["Price_Position"] = (hist["Close"] - hist["Low_60"]) / price_range

        l9         = hist["Low"].rolling(9).min()
        h9         = hist["High"].rolling(9).max()
        kd_range   = (h9 - l9).replace(0, np.nan)
        hist["K"]  = ((hist["Close"] - l9) / kd_range * 100).ewm(com=2, adjust=False).mean()
        hist["D"]  = hist["K"].ewm(com=2, adjust=False).mean()

        delta      = hist["Close"].diff()
        gain       = delta.clip(lower=0)
        loss       = -delta.clip(upper=0)
        rs         = gain.ewm(com=13, adjust=False).mean() / loss.ewm(com=13, adjust=False).mean().replace(0, np.nan)
        hist["RSI"] = (100 - (100 / (1 + rs))).fillna(50)

        tr1        = hist["High"] - hist["Low"]
        tr2        = (hist["High"] - hist["Close"].shift(1)).abs()
        tr3        = (hist["Low"]  - hist["Close"].shift(1)).abs()
        hist["ATR"] = pd.concat([tr1, tr2, tr3], axis=1).max(axis=1).rolling(14).mean()

        hist["MACD"]      = hist["Close"].ewm(span=12, adjust=False).mean() - hist["Close"].ewm(span=26, adjust=False).mean()
        hist["MACD_Hist"] = hist["MACD"] - hist["MACD"].ewm(span=9, adjust=False).mean()

        hist["STD20"]    = hist["Close"].rolling(20).std()
        hist["BB_Width"] = (4 * hist["STD20"]) / hist["20MA"].replace(0, np.nan)

        hist["Is_Bottoming"] = ((hist["Close"] < hist["5MA"]) & (hist["MACD_Hist"].shift(2) < hist["MACD_Hist"].shift(1)) & (hist["MACD_Hist"].shift(1) < hist["MACD_Hist"]) & (hist["MACD_Hist"] < 0)).astype(int)
        hist["Is_Breakout"]      = ((hist["Close"].shift(1) < hist["5MA"].shift(1)) & (hist["Close"] > hist["5MA"]) & (hist["Est_Volume"] > hist["5VMA"] * 1.2))
        hist["Sniper_Signal"]    = (hist["Is_Bottoming"].rolling(3).max().fillna(0).astype(bool) & hist["Is_Breakout"])
        hist["Sniper_Memory_5D"] = hist["Sniper_Signal"].rolling(5).max().fillna(0)

        hist["20_High"]      = hist["High"].rolling(20).max().shift(1)
        hist["Shadow_Ratio"] = (hist["High"] - hist[["Open", "Close"]].max(axis=1)) / (hist["High"] - hist["Low"]).replace(0, 0.001)

        DATA_CACHE.set(symbol, hist)
        return hist
    except Exception as e:
        logger.error(f"[{symbol}] 資料獲取失敗: {e}")
        return None

# =============================================================================
# === 9. 並行資料預載入 (🔥 加入隨機延遲防爬蟲) ===
# =============================================================================
def preload_all_stocks(all_symbols: Dict[str, str]) -> None:
    logger.info(f"開始並行預載 {len(all_symbols)} 支股票資料 (workers={cfg.MAX_WORKERS})...")

    def _fetch(args):
        sym, name = args
        # 隨機延遲 0.2 ~ 1.5 秒，避免瞬間併發被 Yahoo 阻擋
        time.sleep(random.uniform(0.2, 1.5))
        return sym, get_stock_data(sym, name)

    with ThreadPoolExecutor(max_workers=cfg.MAX_WORKERS) as executor:
        futures = {executor.submit(_fetch, item): item[0] for item in all_symbols.items()}
        for future in as_completed(futures):
            sym = futures[future]
            try:
                future.result()
            except Exception as e:
                logger.error(f"[{sym}] 預載例外: {e}")
    logger.info("並行預載完成。")

# =============================================================================
# === 10. 圖表渲染引擎 ===
# =============================================================================
def draw_chart_if_needed(hist: pd.DataFrame, symbol: str) -> str:
    chart_file = f"{symbol}_chart.png"
    try:
        mc    = mpf.make_marketcolors(up="red", down="green", edge="black", wick="black", volume="gray")
        style = mpf.make_mpf_style(base_mpf_style="yahoo", marketcolors=mc)
        mpf.plot(hist[-60:], type="candle", style=style, volume=True, mav=(5, 20), title=f"Stock: {symbol}", savefig=chart_file)
    except:
        try: mpf.plot(hist[-60:], type="candle", style="yahoo", volume=True, mav=(5, 20), title=f"Stock: {symbol}", savefig=chart_file)
        except: pass
    return chart_file

# =============================================================================
# === 11. 發送模組 ===
# =============================================================================
def send_reports(subject: str, text_body: str, chart_files: list) -> None:
    if TG_TOKEN and TG_CHAT_ID:
        chunks = [text_body[i:i+4000] for i in range(0, len(text_body), 4000)]
        for chunk in chunks:
            try: requests.post(f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage", json={"chat_id": TG_CHAT_ID, "text": chunk, "disable_web_page_preview": True}, timeout=10)
            except Exception as e: logger.error(f"Telegram 發送失敗: {e}")
            
    if EMAIL_USER and EMAIL_PASS and EMAIL_TO:
        try:
            msg = MIMEMultipart()
            msg["From"], msg["To"], msg["Subject"] = EMAIL_USER, EMAIL_TO, subject
            msg.attach(MIMEText(text_body, "plain", "utf-8"))
            for chart in chart_files:
                if Path(chart).exists():
                    msg.attach(MIMEImage(open(chart, "rb").read(), name=Path(chart).name))
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30) as server:
                server.login(EMAIL_USER, EMAIL_PASS)
                server.send_message(msg)
        except Exception as e: logger.error(f"Email 發送失敗: {e}")

# =============================================================================
# === 12. 主程式 ===
# =============================================================================
if __name__ == "__main__":
    tw_tz      = datetime.timezone(datetime.timedelta(hours=8))
    curr_dt    = datetime.datetime.now(tw_tz)
    curr_date  = curr_dt.date()
    curr_time  = curr_dt.strftime("%Y-%m-%d %H:%M:%S")

    logger.info(f"NOC 終極戰情室 v12.0 啟動，時間：{curr_time}")

    # --- 交易日檢查與背景狀態更新 ---
    if not is_trading_day(curr_date):
        logger.info("今日為休市日，戰情室發送休市通知後休眠。")
        update_trello_system_status_bg("國定假日/休市", "🔴")
        
        # 🌟 既然排程是週一到週五，會觸發這裡通常就是遇到國定假日或颱風假。
        # 為了避免同一天內被排程重複打擾，我們限定只在「早上 10 點前」的那波喚醒才發送推播。
        curr_hour = curr_dt.hour
        if curr_hour <= 10:
            closed_msg = (
                f"📡 【NOC 終極戰情室 休市通知】\n"
                f"📅 時間：{curr_time}\n"
                f"━━━━━━━━━━━━━━\n"
                f"🔴 今日為國定假日或台股未開盤，戰情室暫停盤勢分析。\n"
                f"伺服器已自動進入休眠模式，祝您假期愉快！"
            )
            send_reports(f"NOC 戰情報告 {curr_date} (休市)", closed_msg, [])
        
        sys.exit(0)

    update_trello_system_status_bg("交易日運作中", "🟢")

    # --- 載入 Trello 部署 ---
    TRELLO_DICT, TRELLO_PORTFOLIO = fetch_trello_deployment()
    STOCK_DICT   = TRELLO_DICT   if TRELLO_DICT   is not None else {}
    MY_PORTFOLIO = TRELLO_PORTFOLIO if TRELLO_PORTFOLIO is not None else {}

    # --- 載入雷達名單 ---
    for fname, label in [(cfg.RADAR_FILE, "🎯 雷達鎖定 (新進火種區)"), (cfg.LIGHTNING_FILE, "⚡ 雷達鎖定 (短線飆股區)")]:
        if Path(fname).exists():
            try:
                with open(fname, "r", encoding="utf-8") as f: stocks = json.load(f)
                if stocks: STOCK_DICT[label] = stocks
            except: pass

    # --- 並行預載所有股票 ---
    all_symbols: Dict[str, str] = {}
    if MY_PORTFOLIO:
        for sym, data in MY_PORTFOLIO.items(): all_symbols[sym] = data["name"]
    for _, stocks in STOCK_DICT.items():
        if stocks:
            for sym, item in stocks.items(): all_symbols[sym] = item.get("name", sym) if isinstance(item, dict) else item

    preload_all_stocks(all_symbols)

    # --- 大盤狀態 ---
    is_bull_market, market_msg = get_market_regime()
    noc_state      : Dict[str, StockState] = load_state()

    msg_list       : list = [f"🌐 【大盤風向】: {market_msg}\n"]
    generated_charts: list = []
    has_data       = False

    # =========================================================================
    # === 戰區 1：庫藏股 ===
    # =========================================================================
    if MY_PORTFOLIO:
        msg_list.append("━━━━━━━━━━━━━━\n💼 【庫藏股 (實體持股動態防禦)】\n━━━━━━━━━━━━━━\n")
        for sym, data in MY_PORTFOLIO.items():
            hist = get_stock_data(sym, data["name"])
            if hist is None: continue

            td, has_data = hist.iloc[-1], True
            curr_price  = td["Close"]
            atr         = td["ATR"] if not pd.isna(td["ATR"]) else 0
            buy_price   = data["buy_price"]
            roi_pct     = ((curr_price - buy_price) / buy_price) * 100 if buy_price else 0

            etf_icon, bias_limit, etf_desc = get_etf_strategy(sym, data["name"])
            is_etf     = "一般型" not in etf_icon
            sym_state  = noc_state.get(sym, StockState())

            if is_etf:
                if sym_state.status != "REAL_HOLD_ETF":
                    noc_state[sym] = StockState(status="REAL_HOLD_ETF", entry=buy_price)
                if roi_pct <= -10.0: pnl_alert = f"💎【黃金坑加碼】帳面回檔 {roi_pct:.2f}%，啟動大額建倉！"
                elif roi_pct <= -5.0: pnl_alert = f"📉【紀律扣款】帳面回檔 {roi_pct:.2f}%，維持定期定額。"
                else: pnl_alert = "🧘‍♂️【長線鎖籌】無懼波動，靜待資產翻倍。"
            else:
                stop_dist = atr * cfg.ATR_MULTIPLIER if atr > 0 else 0
                if sym_state.status != "REAL_HOLD":
                    noc_state[sym] = StockState(status="REAL_HOLD", entry=buy_price, trailing_stop=curr_price - stop_dist)
                    sym_state = noc_state[sym]

                final_stop = max(sym_state.trailing_stop, curr_price - stop_dist)

                if curr_price < final_stop: pnl_alert = f"🩸【拔線警戒】跌破防守線 {final_stop:.1f}，請嚴格執行離場！"
                else:
                    noc_state[sym].trailing_stop = final_stop
                    pnl_alert = f"🔥 獲利巡航 | 📍 防線墊高至: {final_stop:.1f}" if roi_pct > 0 else f"🟡 浮虧防禦 | 📍 死守底線: {final_stop:.1f}"

            chart_file = draw_chart_if_needed(hist, sym)
            generated_charts.append(chart_file)
            msg_list.append(f"{etf_icon} {data['name']} ({sym})\n   成本: {buy_price:.2f} | 股數: {data['shares']} | 現價: {curr_price:.2f}\n   損益: {roi_pct:+.2f}% | 👉 指令: {pnl_alert}\n\n")

    # =========================================================================
    # === 戰區 2-5：Trello & 雷達觀察網域 ===
    # =========================================================================
    for cat, stocks in STOCK_DICT.items():
        if not stocks: continue

        is_etf_zone    = "ETF" in cat.upper()
        is_radar_zone  = "雷達" in cat
        is_key_obs     = "重點觀測" in cat
        is_normal_obs  = "觀察" in cat and not is_key_obs and not is_radar_zone

        cat_msg_list = []

        for sym, item in stocks.items():
            name = item.get("name", sym) if isinstance(item, dict) else item
            tips = item.get("trello_tip", "") if isinstance(item, dict) else ""

            hist = get_stock_data(sym, name)
            if hist is None: continue

            td, has_data = hist.iloc[-1], True

            close    = td["Close"]
            atr      = td["ATR"] if not pd.isna(td.get("ATR", float("nan"))) else 0
            rsi      = td["RSI"]
            vma5     = td["5VMA"]
            ma5      = td["5MA"]
            ma20     = td["20MA"]
            k, d     = td["K"], td["D"]
            est_vol  = td.get("Est_Volume", td["Volume"])
            pos      = td["Price_Position"] if not pd.isna(td.get("Price_Position", float("nan"))) else 0.5
            trust_streak = int(td.get("Trust_Streak", 0)) if not pd.isna(td.get("Trust_Streak", float("nan"))) else 0
            bias     = ((close - ma20) / ma20) * 100 if ma20 else 0

            pe, yoy  = get_pe_ratio(sym), get_revenue_yoy(sym)

            # 量能、趨勢標籤
            if est_vol > vma5 * 1.2: vol_status = "📈 出量"
            elif est_vol < vma5 * 0.8: vol_status = "📉 量縮"
            else: vol_status = "➖ 量平"

            if close > ma5 > ma20: trend_status = "🔥 多頭"
            elif close < ma5 < ma20: trend_status = "🧊 空頭"
            else: trend_status = "🔄 盤整"

            yoy_label = f"{yoy:.2f}%" if isinstance(yoy, float) else yoy
            if isinstance(yoy, float) and yoy >= cfg.YOY_EXPLOSION_PCT: yoy_label += " (🌟 業績大爆發)"

            kd_str = f"K:{k:.1f} D:{d:.1f}"
            if k < 30 and k > d and hist["K"].iloc[-2] <= hist["D"].iloc[-2]: kd_str += " (🌟 KD金叉)"
            elif k > 80: kd_str += " (⚠️ 短線過熱)"

            pe_str = f"{pe:.1f}" if isinstance(pe, float) else pe
            is_overvalued = isinstance(pe, float) and pe > cfg.PE_LIMIT

            chip_msg = td.get("Chip_Status", "➖ 無資料")
            if trust_streak > 0: chip_msg += f" (連買 {trust_streak} 天)"
            elif trust_streak < 0: chip_msg += f" (連賣 {abs(trust_streak)} 天)"

            # 預判訊息 (恢復 Emoji)
            predict_msg = "無特殊徵兆"
            shadow_ratio = td.get("Shadow_Ratio", 0)
            bb_width     = td.get("BB_Width", 1)
            if est_vol > vma5 * 2:
                if pos > 0.7: predict_msg = "💀【動能竭盡】高檔爆量轉折！"
                elif pos < 0.3: predict_msg = "🔥【底部換手】低檔爆量，醞釀反彈！"
                else: predict_msg = "⚠️【中繼爆量】留意方向表態！"
            elif shadow_ratio > 0.5 and est_vol > vma5 * 1.5:
                if pos > 0.7: predict_msg = "⚠️【避雷針陷阱】高檔長上影線！"
                elif pos < 0.3: predict_msg = "🌟【仙人指路】低檔長上影線試盤！"
            elif close > td.get("20_High", 0) and est_vol > vma5 * 1.2: predict_msg = "🚀【無壓巡航】突破 20 日高！"
            elif not pd.isna(bb_width) and bb_width < 0.08: predict_msg = "⚠️【大變盤預警】通道極度壓縮！"

            stop_dist = atr * cfg.ATR_MULTIPLIER if atr > 0 else 999999
            safe_stop = stop_dist if stop_dist > 0 else 999999
            suggested_shares = min(math.floor((cfg.TOTAL_CAPITAL * cfg.RISK_PER_TRADE) / safe_stop), math.floor(cfg.TOTAL_CAPITAL / (close if close > 0 else 1.0)))

            sym_state = noc_state.get(sym, StockState())
            alert     = "✅ 持股觀望"

            if sym_state.status == "REAL_HOLD":
                alert = f"💼 已列入持股防禦區 | 📍 防線: {sym_state.trailing_stop:.1f}"
            elif sym_state.status == "NONE":
                if td.get("Sniper_Signal", False):
                    if not is_bull_market: alert = "🛡️【大盤攔截】大盤偏空，放棄狙擊。"
                    elif isinstance(yoy, float) and yoy < 0: alert = "🛡️【基本面攔截】營收衰退，避開地雷。"
                    elif is_overvalued: alert = f"🛡️【估值攔截】PE {pe_str} 過高，風險極大。"
                    else:
                        stop_price = close - stop_dist
                        noc_state[sym] = StockState(status="HOLD", entry=close, trailing_stop=stop_price)
                        prefix = "⚔️【雙劍合璧】" if isinstance(yoy, float) and yoy >= cfg.YOY_EXPLOSION_PCT else "🚀【啟動狙擊】"
                        alert  = f"{prefix}買入 {suggested_shares/1000:.1f} 張，停損 {stop_price:.1f}"
                elif td.get("Sniper_Memory_5D", 0) == 1:
                    alert = "🔥【狙擊延續】站穩5日線！" if close > ma5 else "⚠️【狙擊失效】跌破5日線！"
            elif sym_state.status == "HOLD":
                new_stop = max(sym_state.trailing_stop, close - stop_dist)
                if close < new_stop:
                    alert = f"🩸【拔線離場】跌破防守線 {new_stop:.1f}！"
                    noc_state[sym] = StockState(status="NONE")
                else:
                    noc_state[sym].trailing_stop = new_stop
                    alert = f"🔥【波段抱牢】防守線: {new_stop:.1f}"

            write_noc_log(curr_date, sym, name, close, rsi, vol_status, trend_status, predict_msg, chip_msg, alert)

            # --- ETF 戰區 ---
            if is_etf_zone:
                etf_type, bias_limit, etf_desc = get_etf_strategy(sym, name)
                if bias > bias_limit: etf_cmd = "⚠️ 乖離過熱，建議分批獲利了結"
                elif k < 30 and k > d: etf_cmd = "🔥 KD低檔金叉，建議佈局買進"
                elif close > ma5: etf_cmd = "✅ 趨勢向上，續抱"
                else: etf_cmd = "⏳ 趨勢偏弱，觀望"

                s = f"{etf_type} {name} ({sym})\n   現價: {close:.2f} | 乖離: {bias:+.1f}% ({'🚨過熱' if bias > bias_limit else '✅穩定'})\n   屬性: {etf_desc}\n   👉 指令: {etf_cmd}\n"
                if tips: s += f"   💡 戰略提示: {tips}\n"
                cat_msg_list.append(s + "\n")
                chart_file = draw_chart_if_needed(hist, sym)
                if chart_file not in generated_charts: generated_charts.append(chart_file)

            # --- 雷達戰區 ---
            elif is_radar_zone:
                s = f"🎯 {name} ({sym})\n   現價: {close:.2f} | 狀態: {trend_status} | {vol_status}\n   指標: {kd_str} | RSI: {rsi:.1f}\n   💰 籌碼: {chip_msg}\n   👉 指令: {alert}\n"
                if tips: s += f"   💡 戰略提示: {tips}\n"
                cat_msg_list.append(s + "\n")
                chart_file = draw_chart_if_needed(hist, sym)
                if chart_file not in generated_charts: generated_charts.append(chart_file)

            # --- 重點觀測 ---
            elif is_key_obs:
                etf_icon, _, etf_desc = get_etf_strategy(sym, name)
                s = f"{etf_icon} {name} ({sym})\n   現價: {close:.2f} | 乖離: {bias:+.1f}% | PE: {pe_str}\n   指標: {kd_str} | RSI: {rsi:.1f} | 類型: {etf_desc}\n   狀態: {trend_status} | YoY: {yoy_label}\n   💰 籌碼: {chip_msg}\n   🔮 預判: {predict_msg}\n   👉 指令: {alert}\n"
                if tips: s += f"   💡 戰略提示: {tips}\n"
                cat_msg_list.append(s + "\n")
                chart_file = draw_chart_if_needed(hist, sym)
                if chart_file not in generated_charts: generated_charts.append(chart_file)

            # --- 普通觀察（條件觸發才顯示）---
            elif is_normal_obs:
                is_2560    = bool(td.get("Signal_2560", False))
                trap_msgs  = {"💀【動能竭盡】高檔爆量轉折！", "⚠️【避雷針陷阱】高檔長上影線！", "⚠️【大變盤預警】通道極度壓縮！"}
                is_trap    = predict_msg in trap_msgs
                is_recovery = bool(td.get("Sniper_Signal", False)) or (k < 30 and k > d) or ("止跌" in tips or "支撐" in tips) or predict_msg in {"🔥【底部換手】低檔爆量，醞釀反彈！", "🌟【仙人指路】低檔長上影線試盤！"} or is_2560
                
                if not (is_trap or is_recovery): continue

                if is_2560:
                    predict_msg = "🎯【2560戰法】量縮回踩 25MA，絕佳左側佈局點！"
                    alert       = "✅ 準備進場 (請留意停損設 25MA 下方 3%)"

                chart_file = draw_chart_if_needed(hist, sym)
                if chart_file not in generated_charts: generated_charts.append(chart_file)

                trigger_label = "🌟 高勝率回踩狙擊" if is_2560 else ("🚨 陷阱預警" if is_trap else "🔥 復甦/狙擊訊號")
                s = f"👀 {name} ({sym})\n   現價: {close:.2f} | RSI: {rsi:.1f} | 乖離: {bias:+.1f}%\n   💰 籌碼: {chip_msg}\n   🎯 條件觸發: {trigger_label}\n   👉 預判/指令: {predict_msg if is_trap else alert}\n"
                if tips: s += f"   💡 戰略提示: {tips}\n"
                cat_msg_list.append(s + "\n")

            # --- 未分類 ---
            else:
                s = f"🔸 {name} ({sym})\n   現價: {close:.2f} | 狀態: {trend_status}\n   👉 指令: {alert}\n"
                if tips: s += f"   💡 戰略提示: {tips}\n"
                cat_msg_list.append(s + "\n")
                chart_file = draw_chart_if_needed(hist, sym)
                if chart_file not in generated_charts: generated_charts.append(chart_file)

        if cat_msg_list:
            msg_list.append(f"━━━━━━━━━━━━━━\n📂 【{cat}】\n━━━━━━━━━━━━━━\n")
            msg_list.extend(cat_msg_list)

    # =========================================================================
    # === ETF 績效競技場 ===
    # =========================================================================
    etf_arena   = {"💰高股息防禦組": [], "🚀市值與主題成長組": []}
    current_year = curr_date.year

    all_etfs: Dict[str, str] = {}
    if MY_PORTFOLIO:
        for sym, data in MY_PORTFOLIO.items(): all_etfs[sym] = data["name"]
    for _, stocks in STOCK_DICT.items():
        if stocks:
            for sym, item in stocks.items(): all_etfs[sym] = item.get("name", sym) if isinstance(item, dict) else item

    for sym, name in all_etfs.items():
        etf_icon, _, _ = get_etf_strategy(sym, name)
        if "一般型" in etf_icon: continue

        hist = get_stock_data(sym, name)
        if hist is None or len(hist) < 10: continue

        close_price = hist["Close"].iloc[-1]
        qtr_days    = min(60, len(hist) - 1)
        qtr_price   = hist["Close"].iloc[-(qtr_days + 1)]
        qtr_roi     = ((close_price - qtr_price) / qtr_price) * 100 if qtr_price else 0

        hist_ytd = hist[hist.index.year == current_year]
        ytd_roi  = (((close_price - hist_ytd["Close"].iloc[0]) / hist_ytd["Close"].iloc[0]) * 100 if not hist_ytd.empty and hist_ytd["Close"].iloc[0] != 0 else qtr_roi)

        group_key = "💰高股息防禦組" if "高股息" in etf_icon else "🚀市值與主題成長組"
        etf_arena[group_key].append({"name": name, "sym": sym, "qtr_roi": qtr_roi, "ytd_roi": ytd_roi})

    arena_lines = []
    if any(etf_arena.values()):
        arena_lines.append("━━━━━━━━━━━━━━\n🏆 【ETF 雙引擎績效競技場 (自動汰弱留強)】\n━━━━━━━━━━━━━━\n")
        medals = ["🥇", "🥈", "🥉"]
        for group_name, group_data in etf_arena.items():
            if not group_data: continue
            arena_lines.append(f"**{group_name}**\n")
            for idx, etf in enumerate(sorted(group_data, key=lambda x: x["qtr_roi"], reverse=True)):
                rank   = medals[idx] if idx < 3 else "🔸"
                q_str  = f"{etf['qtr_roi']:+.1f}%"
                y_str  = f"{etf['ytd_roi']:+.1f}%"
                if etf["qtr_roi"] > 5.0 and etf["ytd_roi"] > 10.0: status = "🔥 雙料強勢"
                elif etf["qtr_roi"] < 0 and etf["ytd_roi"] > 0: status = "⏳ 短線洗盤，長線穩健"
                elif etf["qtr_roi"] < -2.0 and etf["ytd_roi"] < 0: status = "⚠️ 嚴重落後，請檢視佔比"
                else: status = "✅ 穩定跟隨"
                arena_lines.append(f"{rank} {etf['name']} ({etf['sym']})\n   季動能 {q_str} ｜ 本年累計 {y_str} ({status})\n")
            arena_lines.append("\n")
        msg_list.extend(arena_lines)

    # =========================================================================
    # === 收尾：靜默模式判斷 & 發送 ===
    # =========================================================================
    if not has_data:
        logger.warning("無法取得任何股票資料，程式結束。")
        sys.exit(0)

    save_state(noc_state)

    if len(msg_list) == 1 and "大盤風向" in msg_list[0]:
        if cfg.SILENT_MODE:
            logger.info("靜默模式啟用：無觸發條件，本次不發送報告。")
            sys.exit(0)
        else:
            msg_list.append("\n🔕 【靜默模式】無觸發條件。")

    header    = f"📡 【NOC 終極戰情室 v12.0 (效能與體驗優化版)】\n📅 時間：{curr_time}\n━━━━━━━━━━━━━━\n"
    final_text = header + "".join(msg_list)

    send_reports(f"NOC 戰情報告 {curr_date}", final_text, generated_charts)
    logger.info("報告發送完成，清理暫存圖表...")

    for chart in generated_charts:
        try:
            if Path(chart).exists(): Path(chart).unlink()
        except OSError as e: logger.warning(f"圖表清理失敗 [{chart}]: {e}")

    logger.info("NOC 終極戰情室 v12.0 執行完畢。")
