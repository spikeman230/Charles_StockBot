# =============================================================================
# NOC 終極戰情室 v16.12 長短雙軌版（含除錯模式）
# 核心功能：初升段即時偵測、過熱攔截、白名單強制輸出、四象限矩陣
# 整合：旱地拔蔥、狙擊金叉（統一使用 noc_core 函數）
# 除錯模式：DEBUG_FORCE_PUSH = True 時，忽略過熱/四象限/黃燈/攻擊信號，強制推播所有股票
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
import hashlib
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field, asdict
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from dotenv import load_dotenv
from typing import Optional, Dict, Tuple, Any
from pathlib import Path

from noc_core import (
    NOCDatabase, NOCStrategy, NOCDataFetcher, NOCRiskManager,
    analyze_chip_tactics, NOCChipMatrix, is_high_quality_signal,
    assess_volume_turnover_signal, is_overheated, detect_initial_breakout,
    calculate_monster_breakout, calculate_sniper_signal
)

# =============================================================================
# 除錯模式開關 (True = 強制推播所有觀察股，忽略過熱/大盤/信號過濾)
# =============================================================================
DEBUG_FORCE_PUSH = False # 除錯完成後改為 False / True

# =============================================================================
# 初始化與組態
# =============================================================================
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
logging.getLogger('yfinance').setLevel(logging.CRITICAL)

TG_TOKEN = os.getenv("TG_TOKEN")
TG_CHAT_ID = os.getenv("TG_CHAT_ID")
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")
EMAIL_TO = os.getenv("EMAIL_TO")
FINMIND_TOKEN = os.getenv("FINMIND_TOKEN")
TRELLO_KEY = os.getenv("TRELLO_KEY")
TRELLO_TOKEN = os.getenv("TRELLO_TOKEN")
TRELLO_BOARD_ID = os.getenv("TRELLO_BOARD_ID")

class Config:
    TOTAL_CAPITAL : float = float(os.getenv("TOTAL_CAPITAL", "130000"))
    RISK_PER_TRADE : float = float(os.getenv("RISK_PER_TRADE", "0.02"))
    ATR_MULTIPLIER : float = float(os.getenv("ATR_MULTIPLIER", "3.0"))
    YOY_EXPLOSION_PCT : float = float(os.getenv("YOY_EXPLOSION_PCT", "10.0"))
    PE_LIMIT : float = float(os.getenv("PE_LIMIT", "40.0"))
    SILENT_MODE : bool = os.getenv("SILENT_MODE", "false").lower() == "true"
    CACHE_TTL_MINUTES : int = int(os.getenv("CACHE_TTL_MINUTES", "30"))
    CACHE_MAX_ITEMS : int = int(os.getenv("CACHE_MAX_ITEMS", "200"))
    MAX_WORKERS : int = int(os.getenv("MAX_WORKERS", "6"))
    STATE_FILE : str = "noc_state.json"
    LOG_FILE_CSV : str = "noc_trading_log.csv"
    RADAR_FILE : str = "radar_targets.json"
    LIGHTNING_FILE : str = "lightning_targets.json"
    GUERRILLA_FILE : str = "guerrilla_targets.json"

    ACTION_WHITELIST : list = ["建倉", "試單", "波段", "佈局", "長線鎖籌", "加碼", "扣款", "獲利巡航", "浮虧防禦", "洗盤耐受", "戰術撤離", "基本面瓦解", "物理防爆門", "護城河瓦解", "洗盤耐受區", "籌碼動能"]
    ACTION_BLACKLIST : list = ["持股觀望", "暫停進場", "嚴格觀望", "不建議進場", "等待", "不動用資金", "不適用", "營收衰退", "大盤進入震盪洗盤期"]

cfg = Config()

# =============================================================================
# 波段狀態管理
# =============================================================================
@dataclass
class StockState:
    status : str = "NONE"
    entry : float = 0.0
    trailing_stop : float = 0.0
    last_alert_hash : str = ""

    def to_dict(self) -> dict:
        return asdict(self)

    @staticmethod
    def from_dict(d: dict) -> "StockState":
        return StockState(
            status = d.get("status", "NONE"),
            entry = float(d.get("entry", 0.0)),
            trailing_stop = float(d.get("trailing_stop", 0.0)),
            last_alert_hash = d.get("last_alert_hash", "")
        )

# =============================================================================
# 快取管理器
# =============================================================================
@dataclass
class CacheEntry:
    data : Any
    timestamp : datetime.datetime = field(default_factory=datetime.datetime.utcnow)

class DataCacheManager:
    def __init__(self, ttl_minutes: int = 30, max_items: int = 200):
        self._cache : Dict[str, CacheEntry] = {}
        self._ttl = datetime.timedelta(minutes=ttl_minutes)
        self._max = max_items
    def get(self, key: str) -> Optional[Any]:
        entry = self._cache.get(key)
        if entry is None:
            return None
        if datetime.datetime.utcnow() - entry.timestamp > self._ttl:
            del self._cache[key]
            return None
        return entry.data
    def set(self, key: str, data: Any) -> None:
        if len(self._cache) >= self._max:
            evict_count = max(1, self._max // 10)
            oldest_keys = sorted(self._cache, key=lambda k: self._cache[k].timestamp)[:evict_count]
            for k in oldest_keys:
                del self._cache[k]
        self._cache[key] = CacheEntry(data=data)

DATA_CACHE = DataCacheManager(ttl_minutes=cfg.CACHE_TTL_MINUTES, max_items=cfg.CACHE_MAX_ITEMS)

# =============================================================================
# 戰略資產歸類與建倉計劃
# =============================================================================
_ETF_DIV_KEYS = ["高股息","優息","0056","00878","00919","00929","00915","00713","00939","00940","00936"]
_ETF_MKT_KEYS = ["0050","006208","市值","00881","科技","半導體","5G","00891","00892","009816"]

def get_etf_strategy(symbol: str, name: str) -> Tuple[str, float, str]:
    if any(k in name or k in symbol for k in _ETF_DIV_KEYS):
        return "💰高股息", 5.0, "控管殖利率 (5%乖離預警)"
    elif any(k in name or k in symbol for k in _ETF_MKT_KEYS):
        return "🚀市值/主題型", 10.0, "成長動能區 (10%乖離預警)"
    return "🔸一般型", 8.0, "趨勢防禦區 (8%乖離預警)"

def build_tactical_plan(symbol: str, close: float, hist: pd.DataFrame, trend_score: float, fund_health: str, manual_stop: float = 0.0, market_mode: str = "BEAR") -> str:
    if "衰退" in fund_health or "警報" in fund_health:
        return f" 🛡️ 【基本面攔截】營收年增率衰退，不予執行任何長線養殖建倉！\n"
    if trend_score < 0:
        return f" 🛡️ 【趨勢面攔截】該標的未符合長線多頭條件，放棄長線佈局計畫。\n"

    risk_calculator = NOCRiskManager(total_capital=cfg.TOTAL_CAPITAL)
    defense_data = risk_calculator.get_position_and_defense(symbol, close, hist, market_mode=market_mode, is_yellow_light=False)
    stop_loss = defense_data["defense_line"]
    stop_reason = f"融合風控防禦底線 (ATR倍數: {'1.8' if market_mode=='BULL' else '3.0'})"
    if manual_stop > 0:
        stop_loss = manual_stop
        stop_reason = "總司令絕對防線 (Trello 覆寫價)"

    plan = (
        f" 👉 【長線波段作戰指令】\n"
        f" * 戰術策略：積極長線鎖籌 (符合雙重長線濾網)\n"
        f" * 建議底倉 (長線 7.5%)：{defense_data['core_shares']} 股\n"
        f" * 建議游擊 (短線 7.5%)：{defense_data['tactical_shares']} 股\n"
        f" * 移動防禦底線：{stop_loss:.2f} ({stop_reason})\n"
        f" * 鐵律聲明：收盤價若有效跌破此防線，強制執行變現撤離，嚴禁逆勢加碼平攤！\n"
    )
    return plan

def build_light_plan(symbol: str, close: float, hist: pd.DataFrame, manual_stop: float, market_mode: str) -> str:
    risk_calculator = NOCRiskManager(total_capital=cfg.TOTAL_CAPITAL)
    defense_data = risk_calculator.get_position_and_defense(symbol, close, hist, market_mode=market_mode, is_yellow_light=False)
    stop_loss = defense_data["defense_line"]
    if manual_stop > 0:
        stop_loss = manual_stop
    return (
        f" 👉 【初升段試單指令】\n"
        f" * 建議試單股數：{defense_data['total_shares']} 股 (總資金5-10%)\n"
        f" * 移動防禦底線：{stop_loss:.2f}\n"
        f" * 鐵律：若三日內未站穩，立即減碼。\n"
    )

# =============================================================================
# 交易日感知
# =============================================================================
def is_trading_day(curr_date: datetime.date) -> bool:
    # 除錯模式下強制為交易日
    if DEBUG_FORCE_PUSH:
        return True
    if curr_date.weekday() >= 5:
        return False
    try:
        tsm = yf.Ticker("2330.TW").history(period="5d")
        if tsm.empty:
            return True
        last_trading_date = tsm.index[-1].date()
        diff_days = (curr_date - last_trading_date).days
        return diff_days <= 1
    except Exception as e:
        logger.warning(f"交易日感知異常，默認開啟放行: {e}")
        return True

# =============================================================================
# Trello 整合（修正：過濾無效卡片與 NOC 狀態卡）
# =============================================================================
def _trello_params(**extra) -> dict:
    return {"key": TRELLO_KEY, "token": TRELLO_TOKEN, **extra}

def _trello_available() -> bool:
    return all([TRELLO_KEY, TRELLO_TOKEN, TRELLO_BOARD_ID])

def update_trello_system_status(status_msg: str, color: str = "🟢") -> None:
    if not _trello_available():
        return
    url = f"https://api.trello.com/1/boards/{TRELLO_BOARD_ID}/lists"
    try:
        res = requests.get(url, params=_trello_params(cards="open"), timeout=10)
        res.raise_for_status()
        lists_data = res.json()
        if not lists_data:
            return
        first_list_id = lists_data[0]["id"]
        status_card_id = None
        for lst in lists_data:
            for card in lst.get("cards", []):
                if "NOC 系統狀態" in card["name"]:
                    status_card_id = card["id"]
                    break
            if status_card_id:
                break
        tw_tz = datetime.timezone(datetime.timedelta(hours=8))
        date_str = datetime.datetime.now(tw_tz).strftime("%m/%d")
        new_name = f"{color} NOC 系統狀態：{status_msg} ({date_str})"
        if status_card_id:
            requests.put(f"https://api.trello.com/1/cards/{status_card_id}", params=_trello_params(name=new_name), timeout=10)
        else:
            requests.post("https://api.trello.com/1/cards", params=_trello_params(idList=first_list_id, name=new_name, pos="top"), timeout=10)
    except Exception as e:
        logger.error(f"Trello 看板系統狀態更新失敗: {e}")

def update_trello_system_status_bg(status_msg: str, color: str = "🟢") -> None:
    threading.Thread(target=update_trello_system_status, args=(status_msg, color), daemon=True).start()

def _parse_card_to_stock(card: dict) -> Optional[Tuple[str, dict]]:
    raw_name = card["name"].strip()
    ticker_match = re.match(r"^[A-Za-z0-9.]+", raw_name)
    if not ticker_match:
        return None
    symbol = ticker_match.group()
    name_part = raw_name[len(symbol):].strip() if ticker_match else raw_name
    name = re.sub(r"\(.*?\)", "", name_part).strip() or symbol
    title_tip_match = re.search(r"\((.*?)\)", name_part)
    trello_tip = title_tip_match.group(1) if title_tip_match else card.get("desc", "").strip()
    return symbol, {"name": name, "trello_tip": trello_tip}

def _parse_card_to_portfolio(card: dict) -> Optional[Tuple[str, dict]]:
    raw_name = card["name"].strip()
    ticker_match = re.match(r"^[A-Za-z0-9.]+", raw_name)
    if not ticker_match:
        return None
    symbol = ticker_match.group()
    name_part = raw_name[len(symbol):].strip() if ticker_match else raw_name
    name = re.sub(r"\(.*?\)", "", name_part).strip() or symbol
    desc = card.get("desc", "")
    buy_price, shares, manual_stop = 0.0, 1000, 0.0
    price_match = re.search(r"成本[：:]\s*([0-9.]+)", desc)
    shares_match = re.search(r"股數[：:]\s*([0-9]+)", desc)
    stop_match = re.search(r"(防線|停損|防守)[：:]\s*([0-9.]+)", desc)
    if price_match:
        buy_price = float(price_match.group(1))
    if shares_match:
        shares = int(shares_match.group(1))
    if stop_match:
        manual_stop = float(stop_match.group(2))
    return symbol, {"name": name, "buy_price": buy_price, "shares": shares, "trello_tip": desc, "manual_stop": manual_stop}

def fetch_trello_deployment() -> Tuple[Optional[dict], Optional[dict]]:
    if not _trello_available():
        return None, None
    url = f"https://api.trello.com/1/boards/{TRELLO_BOARD_ID}/lists"
    try:
        response = requests.get(url, params=_trello_params(cards="open"), timeout=10)
        response.raise_for_status()
        lists_data = response.json()
        trello_dict, my_portfolio = {}, {}
        for lst in lists_data:
            list_name = lst["name"]
            # 跳過美股相關列表（避免混亂）
            if "美股" in list_name:
                continue
            is_portfolio_list = "庫存" in list_name or "庫藏" in list_name
            for card in lst.get("cards", []):
                card_name = card["name"]
                # 跳過任何包含 "NOC" 的卡片（系統狀態卡）
                if "NOC" in card_name:
                    continue
                if is_portfolio_list:
                    parsed = _parse_card_to_portfolio(card)
                    if parsed is None:
                        continue
                    sym, info = parsed
                    my_portfolio[sym] = info
                else:
                    parsed = _parse_card_to_stock(card)
                    if parsed is None:
                        continue
                    sym, info = parsed
                    trello_dict.setdefault(list_name, {})[sym] = info
        return trello_dict, my_portfolio
    except Exception as e:
        logger.error(f"無法完整拉取 Trello 看板配置: {e}")
        return None, None

# =============================================================================
# 本地狀態管理
# =============================================================================
def load_state() -> Dict[str, StockState]:
    if not Path(cfg.STATE_FILE).exists():
        return {}
    try:
        with open(cfg.STATE_FILE, "r", encoding="utf-8") as f:
            return {sym: StockState.from_dict(d) for sym, d in json.load(f).items()}
    except:
        return {}

def save_state(state: Dict[str, StockState]) -> None:
    try:
        with open(cfg.STATE_FILE, "w", encoding="utf-8") as f:
            json.dump({sym: s.to_dict() for sym, s in state.items()}, f, ensure_ascii=False, indent=4)
    except:
        pass

def write_noc_log(date, symbol, name, close_price, rsi, vol_status, status, predict, chip_signal, alert) -> None:
    log_exists = Path(cfg.LOG_FILE_CSV).exists()
    try:
        with open(cfg.LOG_FILE_CSV, mode="a", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            if not log_exists:
                writer.writerow(["日期","代號","名稱","收盤價","RSI","量能狀態","趨勢狀態","戰場預判","籌碼訊號","行動指令"])
            writer.writerow([date, symbol, name, f"{close_price:.2f}", f"{rsi:.2f}", vol_status, status, predict, chip_signal, alert])
    except:
        pass

# =============================================================================
# 大盤風向與基本面輔助
# =============================================================================
def get_market_regime() -> Tuple[bool, str]:
    try:
        twii = yf.Ticker("^TWII").history(period="2mo")
        if twii.empty:
            raise ValueError("加權指數歷史數據下載失敗")
        twii["20MA"] = twii["Close"].rolling(20).mean()
        is_bull = twii["Close"].iloc[-1] > twii["20MA"].iloc[-1]
        return is_bull, "🟢 多頭格局 (站上月線軌道)" if is_bull else "🔴 空頭警戒 (跌破月線防禦)"
    except Exception as e:
        logger.error(f"大盤技術風向判斷異常: {e}")
        return True, "搶修中 - 🟡 大盤海象未知"

def get_revenue_yoy(symbol: str):
    if not FINMIND_TOKEN:
        return "N/A"
    match = re.search(r"\d+", symbol)
    if not match:
        return "N/A"
    try:
        url = "https://api.finmindtrade.com/api/v4/data"
        params = {
            "dataset": "TaiwanStockMonthRevenue",
            "data_id": match.group(),
            "start_date": (datetime.datetime.now() - datetime.timedelta(days=400)).strftime("%Y-%m-%d"),
            "token": FINMIND_TOKEN
        }
        r = requests.get(url, params=params, timeout=10)
        data = r.json()
        if data.get("msg") == "success" and data.get("data"):
            df = pd.DataFrame(data["data"])
            latest = df.iloc[-1]
            prev = df[(df["revenue_year"] == latest["revenue_year"] - 1) & (df["revenue_month"] == latest["revenue_month"])]
            if not prev.empty and prev.iloc[-1]["revenue"] > 0:
                return float((latest["revenue"] - prev.iloc[-1]["revenue"]) / prev.iloc[-1]["revenue"] * 100)
    except:
        pass
    return "N/A"

def get_pe_ratio(symbol: str):
    try:
        info = yf.Ticker(symbol).info
        pe = info.get("trailingPE") or info.get("forwardPE")
        return pe if pe else "N/A"
    except:
        return "N/A"

def get_finmind_chip_data(symbol: str, start_date_str: str) -> pd.DataFrame:
    if not FINMIND_TOKEN:
        return pd.DataFrame()
    match = re.search(r"\d+", symbol)
    if not match:
        return pd.DataFrame()
    try:
        url = "https://api.finmindtrade.com/api/v4/data"
        params = {
            "dataset": "TaiwanStockInstitutionalInvestorsBuySell",
            "data_id": match.group(),
            "start_date": start_date_str,
            "token": FINMIND_TOKEN
        }
        r = requests.get(url, params=params, timeout=10)
        data = r.json()
        if data.get("msg") == "success" and data.get("data"):
            df = pd.DataFrame(data["data"])
            df["net_buy"] = df["buy"] - df["sell"]
            df["type"] = "Other"
            df.loc[df["name"].str.contains("外資"), "type"] = "Foreign_Inv"
            df.loc[df["name"].str.contains("投信"), "type"] = "Trust_Inv"
            df.loc[df["name"].str.contains("自營商"), "type"] = "Dealer_Inv"
            pivot_df = df.groupby(["date", "type"])["net_buy"].sum().unstack(fill_value=0).reset_index()
            for col in ["Foreign_Inv", "Trust_Inv", "Dealer_Inv"]:
                if col not in pivot_df.columns:
                    pivot_df[col] = 0
            pivot_df["Date"] = pd.to_datetime(pivot_df["date"]).dt.date
            pivot_df.set_index("Date", inplace=True)
            return pivot_df[["Foreign_Inv", "Trust_Inv", "Dealer_Inv"]]
    except:
        pass
    return pd.DataFrame()

def calculate_chip_signals(hist: pd.DataFrame) -> pd.DataFrame:
    hist["Chip_Status"] = "➖ 中性/偏空"
    hist["Trust_Streak"] = 0
    if not {"Foreign_Inv", "Trust_Inv", "Dealer_Inv"}.issubset(hist.columns):
        return hist
    hist["Total_Institutional"] = hist["Foreign_Inv"] + hist["Trust_Inv"] + hist["Dealer_Inv"]
    hist["Signal_CoBuy"] = (hist["Foreign_Inv"] > 0) & (hist["Trust_Inv"] > 0)
    hist["Signal_Trust_Trend"] = ((hist["Trust_Inv"] > 0).astype(int).rolling(5).sum() >= 4) & (hist["Trust_Inv"] > 0)
    trust_dir = np.sign(hist["Trust_Inv"])
    hist["Trust_Streak"] = trust_dir.groupby((trust_dir != trust_dir.shift()).cumsum()).cumsum()
    conds = [hist["Signal_CoBuy"], hist["Signal_Trust_Trend"], hist["Total_Institutional"] > 0]
    hist["Chip_Status"] = np.select(conds, ["🤝 土洋齊買", "🏦 投信作帳", "📈 法人偏多"], default="➖ 中性/偏空")
    return hist

# =============================================================================
# 核心數據抓取與技術指標（使用統一函數計算狙擊金叉與旱地拔蔥）
# =============================================================================
def get_stock_data(symbol: str, name: str) -> Optional[pd.DataFrame]:
    cached = DATA_CACHE.get(symbol)
    if cached is not None:
        return cached
    try:
        match = re.search(r"\d+", symbol)
        if match and FINMIND_TOKEN:
            raw_id = match.group()
            local_db = NOCDatabase()
            local_fetcher = NOCDataFetcher(token=FINMIND_TOKEN)
            local_fetcher.fetch_financial_statements(raw_id, local_db)

        stock = yf.Ticker(symbol)
        info = stock.info
        shares_out = info.get("sharesOutstanding") or info.get("impliedSharesOutstanding")
        hist = stock.history(period="8mo").dropna(subset=["Close"])
        if len(hist) < 60:
            return None

        hist["Shares_Out"] = shares_out if shares_out else np.nan
        hist["Date_Key"] = hist.index.date
        if FINMIND_TOKEN and (".TW" in symbol or ".TWO" in symbol):
            chip_df = get_finmind_chip_data(symbol, (datetime.datetime.now() - datetime.timedelta(days=200)).strftime("%Y-%m-%d"))
            if not chip_df.empty:
                hist = hist.merge(chip_df, left_on="Date_Key", right_index=True, how="left").ffill().fillna(0)

        hist = calculate_chip_signals(hist)

        # 動態量能預估
        now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8)))
        market_open = now.replace(hour=9, minute=0, second=0, microsecond=0)
        market_close = now.replace(hour=13, minute=30, second=0, microsecond=0)
        total_trading_minutes = (market_close - market_open).total_seconds() / 60.0
        if market_open < now < market_close:
            elapsed_mins = max(1.0, (now - market_open).total_seconds() / 60.0)
            vol_mult = total_trading_minutes / elapsed_mins
        else:
            vol_mult = 1.0
        hist["Est_Volume"] = hist["Volume"].copy()
        if len(hist) > 0:
            hist.iloc[-1, hist.columns.get_loc("Est_Volume")] = int(hist["Volume"].iloc[-1] * vol_mult)

        hist["5MA"] = hist["Close"].rolling(5).mean()
        hist["20MA"] = hist["Close"].rolling(20).mean()
        hist["25MA"] = hist["Close"].rolling(25).mean()
        hist["60MA"] = hist["Close"].rolling(60).mean()
        hist["5VMA"] = hist["Est_Volume"].rolling(5).mean()
        hist["60VMA"] = hist["Volume"].rolling(60).mean()

        hist["Turnover_Rate"] = ((hist["Est_Volume"] / hist["Shares_Out"]) * 100).fillna(1.5)
        hist["Volume_Ratio"] = (hist["Est_Volume"] / hist["5VMA"].shift(1)).fillna(1.0)

        # K線特徵
        hist['Candle_Ratio'] = (hist['High'] - hist[['Open','Close']].max(axis=1)) / (hist['High'] - hist['Low'] + 1e-9)
        hist['Close_vs_High'] = hist['Close'] / hist['High']
        hist['Is_Red'] = hist['Close'] >= hist['Open']

        # 乖離與漲幅（用於過熱攔截）
        hist['Bias_20MA'] = (hist['Close'] - hist['20MA']) / hist['20MA'] * 100
        hist['Bias_60MA'] = (hist['Close'] - hist['60MA']) / hist['60MA'] * 100
        hist['Return_5D'] = hist['Close'].pct_change(5) * 100
        hist['Return_10D'] = hist['Close'].pct_change(10) * 100

        hist["25MA_Rising"] = hist["25MA"] > hist["25MA"].shift(1)
        hist["Is_Red_Candle"] = hist["Close"] > hist["Open"]
        hist["Lower_Shadow_Ratio"] = (hist[["Open", "Close"]].min(axis=1) - hist["Low"]) / (hist["High"] - hist["Low"]).replace(0, 0.001)

        hist["Signal_2560"] = (hist["25MA"] > hist["25MA"].shift(3)) & (hist["5VMA"] > hist["60VMA"]) & (hist["Low"] <= hist["25MA"] * 1.015) & (hist["Close"] >= hist["25MA"] * 0.985) & (hist["Est_Volume"] < hist["5VMA"])
        hist["High_60"] = hist["High"].rolling(window=60, min_periods=20).max()
        hist["Low_60"] = hist["Low"].rolling(window=60, min_periods=20).min()
        hist["Price_Position"] = (hist["Close"] - hist["Low_60"]) / (hist["High_60"] - hist["Low_60"]).replace(0, np.nan)

        l9, h9 = hist["Low"].rolling(9).min(), hist["High"].rolling(9).max()
        hist["K"] = ((hist["Close"] - l9) / (h9 - l9).replace(0, np.nan) * 100).ewm(com=2, adjust=False).mean()
        hist["D"] = hist["K"].ewm(com=2, adjust=False).mean()

        delta = hist["Close"].diff()
        rs = delta.clip(lower=0).ewm(com=13, adjust=False).mean() / (-delta.clip(upper=0)).ewm(com=13, adjust=False).mean().replace(0, np.nan)
        hist["RSI"] = (100 - (100 / (1 + rs))).fillna(50)

        hist["ATR"] = pd.concat([hist["High"] - hist["Low"], (hist["High"] - hist["Close"].shift(1)).abs(), (hist["Low"] - hist["Close"].shift(1)).abs()], axis=1).max(axis=1).rolling(14).mean()

        hist["MACD"] = hist["Close"].ewm(span=12, adjust=False).mean() - hist["Close"].ewm(span=26, adjust=False).mean()
        hist["MACD_Hist"] = hist["MACD"] - hist["MACD"].ewm(span=9, adjust=False).mean()
        hist["STD20"] = hist["Close"].rolling(20).std()
        hist["BB_Width"] = (4 * hist["STD20"]) / hist["20MA"].replace(0, np.nan)

        # ========== 使用統一函數計算狙擊金叉與旱地拔蔥 ==========
        sniper_val = calculate_sniper_signal(hist)
        hist['Sniper_Signal'] = sniper_val
        hist['Sniper_Memory_5D'] = hist['Sniper_Signal'].rolling(5).max().fillna(0)

        td_temp = hist.iloc[-1]
        monster_val = calculate_monster_breakout(hist, td_temp)
        hist['Monster_Breakout'] = monster_val

        hist["20_High"] = hist["High"].rolling(20).max().shift(1)
        hist["Shadow_Ratio"] = (hist["High"] - hist[["Open", "Close"]].max(axis=1)) / (hist["High"] - hist["Low"]).replace(0, 0.001)

        hist["PE"] = get_pe_ratio(symbol)
        hist["YoY"] = get_revenue_yoy(symbol)

        DATA_CACHE.set(symbol, hist)
        return hist
    except Exception as e:
        logger.error(f"❌ 標的 [{symbol}] 執行技術分析精算失敗: {e}")
        return None

# =============================================================================
# 並行預載入
# =============================================================================
def preload_all_stocks(all_symbols: Dict[str, str]) -> None:
    logger.info(f"啟動高並行預載快取電路，共計 {len(all_symbols)} 檔標的...")
    def _fetch(args):
        sym, name = args
        time.sleep(random.uniform(0.1, 1.0))
        return sym, get_stock_data(sym, name)

    with ThreadPoolExecutor(max_workers=cfg.MAX_WORKERS) as executor:
        futures = {executor.submit(_fetch, item): item[0] for item in all_symbols.items()}
        for future in as_completed(futures):
            try:
                future.result()
            except Exception as e:
                logger.error(f"標的 {futures[future]} 預載入失敗: {e}")

# =============================================================================
# 圖表與推播
# =============================================================================
def draw_chart_if_needed(hist: pd.DataFrame, symbol: str) -> str:
    chart_file = f"{symbol}_chart.png"
    try:
        mc = mpf.make_marketcolors(up="red", down="green", edge="black", wick="black", volume="gray")
        mpf.plot(hist[-60:], type="candle", style=mpf.make_mpf_style(base_mpf_style="yahoo", marketcolors=mc), volume=True, mav=(5, 20, 60), title=f"Stock: {symbol} (Long-Term Wave)", savefig=chart_file)
    except:
        try:
            mpf.plot(hist[-60:], type="candle", style="yahoo", volume=True, mav=(5, 20), title=f"Stock: {symbol}", savefig=chart_file)
        except:
            pass
    return chart_file

def send_reports(subject: str, text_body: str, chart_files: list) -> None:
    if TG_TOKEN and TG_CHAT_ID:
        for chunk in [text_body[i:i+4000] for i in range(0, len(text_body), 4000)]:
            try:
                requests.post(f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage", json={"chat_id": TG_CHAT_ID, "text": chunk, "disable_web_page_preview": True}, timeout=10)
            except:
                pass
    if EMAIL_USER and EMAIL_PASS and EMAIL_TO:
        try:
            msg = MIMEMultipart()
            msg["From"], msg["To"], msg["Subject"] = EMAIL_USER, EMAIL_TO, subject
            msg.attach(MIMEText(text_body, "plain", "utf-8"))
            for chart in chart_files:
                if Path(chart).exists():
                    with open(chart, "rb") as f:
                        msg.attach(MIMEImage(f.read(), name=Path(chart).name))
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30) as server:
                server.login(EMAIL_USER, EMAIL_PASS)
                server.send_message(msg)
        except:
            pass

# =============================================================================
# 主程式
# =============================================================================
if __name__ == "__main__":
    tw_tz = datetime.timezone(datetime.timedelta(hours=8))
    curr_dt = datetime.datetime.now(tw_tz)
    curr_date, curr_time = curr_dt.date(), curr_dt.strftime("%Y-%m-%d %H:%M:%S")

    logger.info(f"NOC 終極戰情室 v16.12 長短雙軌版（除錯模式={'ON' if DEBUG_FORCE_PUSH else 'OFF'}）啟動。時間：{curr_time}")

    db = NOCDatabase()
    strategy = NOCStrategy(db)
    chip_matrix_analyzer = NOCChipMatrix()

    msg_list = []

    macro_info = strategy.get_macro_status()
    is_yellow_light = False

    # 除錯模式：忽略大盤紅燈
    if not DEBUG_FORCE_PUSH and macro_info["status"] == "🔴 紅燈":
        logger.warning("🚨🚨🚨 觸發戰略級拔插頭熔斷協議！大盤環境進入極度危險空頭階段。")
        update_trello_system_status_bg("⚠️ 觸發空頭防禦協議 (全面停止建倉)", "🔴")
        send_reports(
            f"🚨 NOC 系統最高防空警報 {curr_date}",
            f"📡 【NOC 系統強制熔斷通知】\n📅 時間：{curr_time}\n━━━━━━━━━━━━━━\n大盤目前狀態為：{macro_info['status']} - {macro_info['desc']}\n已觸發最高資產保護協議，全系統雷達冷卻關閉，嚴格禁止任何開新倉買進動作！請總司令檢視既有長線持股！",
            []
        )
        sys.exit(0)
    elif "黃燈" in macro_info["status"] or macro_info["status"] == "🟡 黃燈":
        logger.warning("🟡 觸發大盤黃燈防禦電路！總兵力天花板強制鎖定 50% 水位 (6.5萬) / 雷達新火種禁止開新倉 / 防守線緊縮至 2.0 ATR 或月線。")
        cfg.TOTAL_CAPITAL = float(os.getenv("TOTAL_CAPITAL", "130000")) * 0.5
        cfg.ATR_MULTIPLIER = 2.0
        is_yellow_light = True
        update_trello_system_status_bg("🟡 黃燈防禦協議 (半倉/收緊防護)", "🟡")

    if not is_trading_day(curr_date):
        logger.info("今日非台股交易日。戰情室啟動靜默休眠機制。")
        update_trello_system_status_bg("非交易日/休市靜默", "🔴")
        if curr_dt.hour <= 10:
            send_reports(f"NOC 戰情報告 {curr_date} (休市)", f"📡 【NOC 戰情室靜默休眠】\n📅 時間：{curr_time}\n━━━━━━━━━━━━━━\n🔴 今日市場休市，全系統處於資產監守維護狀態，不推播繁雜雜訊。", [])

    if not is_yellow_light:
        logger.info("通過環境感知檢查，開始同步雲端 Trello 看板部署...")
        update_trello_system_status_bg("交易日波段追蹤中", "🟢")

    TRELLO_DICT, TRELLO_PORTFOLIO = fetch_trello_deployment()
    STOCK_DICT = TRELLO_DICT if TRELLO_DICT else {}
    MY_PORTFOLIO = TRELLO_PORTFOLIO if TRELLO_PORTFOLIO else {}

    for fname, label in [(cfg.RADAR_FILE, "👀 長線觀察區 (雷達自動火種)"), (cfg.LIGHTNING_FILE, "👀 短線觀察區 (閃電自動火種)")]:
        if Path(fname).exists():
            try:
                with open(fname, "r", encoding="utf-8") as f:
                    STOCK_DICT[label] = json.load(f)
            except Exception as e:
                logger.error(f"讀取 {fname} 失敗: {e}")
    
    # 整合短線飆股搜尋器結果 (Momentum)
    MOMENTUM_FILE = "momentum_targets.json"
    if Path(MOMENTUM_FILE).exists():
        try:
            with open(MOMENTUM_FILE, "r", encoding="utf-8") as f:
                momentum_data = json.load(f)
            if momentum_data:
                STOCK_DICT["⚡短線飆股區 (Momentum)"] = momentum_data
                logger.info(f"✅ 已載入 {len(momentum_data)} 檔短線飆股至追蹤區")
        except Exception as e:
            logger.error(f"讀取短線飆股清單失敗: {e}")
    
    all_symbols = {sym: data["name"] for sym, data in MY_PORTFOLIO.items()}
    for stocks in STOCK_DICT.values():
        for sym, item in stocks.items():
            all_symbols[sym] = item.get("name", sym) if isinstance(item, dict) else item

    preload_all_stocks(all_symbols)

    is_bull_market, market_msg = get_market_regime()
    market_mode = "BULL" if is_bull_market else "BEAR"
    logger.info(f"📡 市場模式切換 => {market_mode} (大盤訊號: {market_msg})")

    noc_state = load_state()

    macro_msg = f"🌐 【大盤風向儀】：{macro_info['status']} | {market_msg}\n"
    if is_yellow_light:
        macro_msg += "⚠️ 【黃燈防禦】總兵力天花板強制鎖定 50% 水位 (6.5萬) / 雷達新火種禁止開新倉 / 防守線緊縮至 2.0 ATR 或月線\n"

    msg_list = [macro_msg]
    generated_charts = []
    has_data = False
    has_actionable_alerts = False

    # =========================================================================
    # 戰區 1：庫藏股 (白名單強制輸出)
    # =========================================================================
    if MY_PORTFOLIO:
        msg_list.append("━━━━━━━━━━━━━━\n💼 【庫藏股 (長線鎖籌動態防禦動態)】\n━━━━━━━━━━━━━━\n")
        for sym, data in MY_PORTFOLIO.items():
            hist = get_stock_data(sym, data["name"])
            if hist is None:
                continue

            raw_id = re.search(r"\d+", sym).group() if re.search(r"\d+", sym) else sym
            td, has_data = hist.iloc[-1], True
            curr_price, atr = td["Close"], td["ATR"] if not pd.isna(td.get("ATR", float("nan"))) else 0
            buy_price = data["buy_price"]
            roi_pct = ((curr_price - buy_price) / buy_price) * 100 if buy_price else 0

            etf_icon, _, _ = get_etf_strategy(sym, data["name"])

            if sym not in noc_state:
                noc_state[sym] = StockState()
            sym_state = noc_state[sym]

            ma20 = td["20MA"]
            ma60 = td["60MA"]
            turnover = td["Turnover_Rate"]
            vol_ratio = td["Volume_Ratio"]
            yoy_single = td["YoY"]

            if is_yellow_light:
                current_atr_multiplier = 2.0
            else:
                current_atr_multiplier = 1.8 if market_mode == "BULL" else 3.0
            calculated_stop = curr_price - (atr * current_atr_multiplier)
            calculated_stop = min(calculated_stop, ma20) if not pd.isna(ma20) else calculated_stop

            if sym_state.status != "REAL_HOLD" and sym_state.status != "REAL_HOLD_ETF":
                noc_state[sym] = StockState(status="REAL_HOLD", entry=buy_price, trailing_stop=calculated_stop)
                sym_state = noc_state[sym]

            trello_stop = data.get("manual_stop", 0.0)
            if trello_stop > 0:
                final_stop = max(trello_stop, sym_state.trailing_stop, calculated_stop)
            else:
                final_stop = max(sym_state.trailing_stop, calculated_stop)

            fund_health = strategy.get_fundamental_health(raw_id)
            is_accumulated_recession = "衰退" in fund_health

            if is_accumulated_recession:
                pnl_alert = "💀【護城河瓦解】累計營收衰退，明日開盤即刻清倉！"
                noc_state[sym] = StockState(status="NONE")
            elif isinstance(yoy_single, (int, float)) and yoy_single < -10:
                pnl_alert = f"💀【營收急遽衰退】單月年增率 {yoy_single:.1f}%，明日開盤清倉！"
                noc_state[sym] = StockState(status="NONE")
            elif isinstance(yoy_single, (int, float)) and yoy_single < 0:
                pnl_alert = f"⚠️【營運動能轉弱】單月年增率 {yoy_single:.1f}%，減碼觀察，不加碼。"
            elif roi_pct <= -15.0 or curr_price < ma60 or curr_price < final_stop:
                pnl_alert = f"🩸【戰術撤離】跌破防守底線 ({final_stop:.2f})，無條件停損變現！"
                noc_state[sym] = StockState(status="NONE")
            elif trello_stop > 0 and sym_state.trailing_stop != final_stop:
                pnl_alert = f"🛡️【手動指揮】已依據 Trello 覆寫防守線至 {final_stop:.2f}"
                noc_state[sym].trailing_stop = final_stop
            elif roi_pct > 0 and curr_price > ma20:
                pnl_alert = f"🔥【獲利巡航】獲利奔跑中，防禦線上移至 {final_stop:.2f}！"
                noc_state[sym].trailing_stop = final_stop
            elif roi_pct <= 0 and curr_price >= ma60 and curr_price >= final_stop:
                pnl_alert = f"🛡️【洗盤耐受區】嚴禁攤平加碼，死守底線 ({final_stop:.2f})！"
                noc_state[sym].trailing_stop = final_stop
            else:
                pnl_alert = f"🔍【中立觀察】價格震盪，監控防禦底線 ({final_stop:.2f})。"
                noc_state[sym].trailing_stop = final_stop

            silent_keywords = ["中立觀察", "長線鎖籌", "洗盤耐受區"]
            is_silent = any(kw in pnl_alert for kw in silent_keywords)
            if is_silent and cfg.SILENT_MODE:
                logger.info(f"🔇 [靜默模式] 庫藏股 {sym} 指令為 '{pnl_alert}'，符合靜默關鍵字，不進行推播與繪圖。")
            else:
                generated_charts.append(draw_chart_if_needed(hist, sym))
                inv_str = f"{etf_icon} {data['name']} ({sym})\n"
                inv_str += f" 現價: {curr_price:.2f} | 成本: {buy_price:.2f}\n"
                chip_msg = td["Chip_Status"]
                matrix_signal = chip_matrix_analyzer.analyze(hist, market_mode=market_mode)
                inv_str += f" 換手: {turnover:.2f}% | 量比: {vol_ratio:.2f}倍 | 籌碼戰術: {matrix_signal}\n"
                inv_str += f" 💰 法人籌碼: {chip_msg}\n"
                inv_str += f" 📊 累計財報: {fund_health}\n"
                yoy_display = f"{yoy_single:.1f}%" if isinstance(yoy_single, (int, float)) else str(yoy_single)
                inv_str += f" 📈 單月YoY: {yoy_display}\n"
                inv_str += f" 損益: {roi_pct:+.2f}% | 👉 作戰指令: {pnl_alert}\n\n"
                msg_list.append(inv_str)
                has_actionable_alerts = True

    # =========================================================================
    # 戰區 2：觀察區 (白名單: 長線觀測區, 短線觀測區, ⚡ 短線飆股區 (Momentum))
    # =========================================================================
    force_include_categories = ["長線觀測區", "短線觀測區", "⚡ 短線飆股區 (Momentum)"]
    for cat, stocks in STOCK_DICT.items():
        if not stocks:
            continue

        cat_msg_list = []
        for sym, item in stocks.items():
            name = item.get("name", sym) if isinstance(item, dict) else item
            tips = item.get("trello_tip", "") if isinstance(item, dict) else ""

            manual_stop_price = 0.0
            stop_match = re.search(r"(?:死線|防線|停損)[:：]\s*([0-9.]+)", tips)
            if stop_match:
                manual_stop_price = float(stop_match.group(1))

            hist = get_stock_data(sym, name)
            if hist is None:
                continue

            raw_id = re.search(r"\d+", sym).group() if re.search(r"\d+", sym) else sym
            td, has_data = hist.iloc[-1], True
            close, rsi, ma5, ma20 = td["Close"], td["RSI"], td["5MA"], td["20MA"]
            vma5, est_vol = td["5VMA"], td["Est_Volume"]

            atr = td["ATR"] if not pd.isna(td["ATR"]) else 0
            price_position = td["Price_Position"] if not pd.isna(td["Price_Position"]) else 0.5
            trust_streak = int(td["Trust_Streak"])
            bias = ((close - ma20) / ma20) * 100 if ma20 else 0
            pe = td["PE"]
            yoy = td["YoY"]

            turnover = td["Turnover_Rate"]
            vol_ratio = td["Volume_Ratio"]
            shares_out = td.get("Shares_Out", 0.0)

            is_lightning = "短線" in cat
            local_market_mode = "BULL" if is_lightning else market_mode

            trend_score = strategy.get_trend_score(hist, market_mode=local_market_mode)
            fund_health = strategy.get_fundamental_health(raw_id)

            vol_status = "📈 出量" if est_vol > vma5 * 1.2 else ("📉 量縮" if est_vol < vma5 * 0.8 else "➖ 量平")
            trend_status = "🔥 多頭" if close > ma5 > ma20 else ("🧊 空頭" if close < ma5 < ma20 else "🔄 盤整")

            yoy_label = f"{yoy:.2f}%" if isinstance(yoy, float) else str(yoy)
            pe_str = f"{pe:.1f}" if isinstance(pe, float) else str(pe)

            chip_msg = td["Chip_Status"]
            if trust_streak > 0:
                chip_msg += f" (連買 {trust_streak} 天)"
            elif trust_streak < 0:
                chip_msg += f" (連賣 {abs(trust_streak)} 天)"

            if sym not in noc_state:
                noc_state[sym] = StockState()
            sym_state = noc_state[sym]

            alert = "✅ 趨勢追蹤中，尚未觸發佈局點"
            trigger_label = ""
            action_plan_text = ""

            # 黃燈攔截非白名單（除錯模式跳過）
            if not DEBUG_FORCE_PUSH:
                if is_yellow_light and cat not in force_include_categories:
                    logger.debug(f"🟡 黃燈模式跳過 {sym} (分類: {cat})")
                    continue

            # 過熱攔截（除錯模式跳過）
            if not DEBUG_FORCE_PUSH:
                ma20_val = td['20MA'] if not pd.isna(td['20MA']) else 0
                ma60_val = td['60MA'] if not pd.isna(td['60MA']) else 0
                return_5d = td.get('Return_5D', 0)
                return_10d = td.get('Return_10D', 0)
                overheated, over_reason = is_overheated(close, ma20_val, ma60_val, return_5d, return_10d, price_position, vol_ratio)
                if overheated:
                    logger.info(f"🛑 [過熱攔截] {sym} 原因: {over_reason}，強制封鎖推播。")
                    continue

            # 四象限信號（除錯模式跳過危險檢查）
            quadrant_signal = assess_volume_turnover_signal(
                vol_ratio=vol_ratio,
                turnover=turnover,
                shares_out=shares_out,
                price_position=price_position,
                candle_ratio=td['Candle_Ratio'],
                is_red=td['Is_Red'],
                close_vs_high=td['Close_vs_High']
            )
            danger_signals = ("🔴 主力出貨區", "⚠️ 量價背離陷阱", "🔴 爆量長上影 (假突破/出貨)", "⚠️ 黑K出量 (賣壓沉重)")
            if not DEBUG_FORCE_PUSH:
                if quadrant_signal in danger_signals:
                    logger.info(f"🛑 [四象限攔截] {sym} 信號為 {quadrant_signal}，強制封鎖推播。")
                    continue

            # 狀態機觸發判斷（優先級：初升段突破 > 旱地拔蔥 > 狙擊金叉）
            if sym_state.status == "REAL_HOLD":
                alert = f"💼 持股防禦區 | 📍 最新防線: {sym_state.trailing_stop:.1f}"
            elif sym_state.status == "NONE":
                initial_break, break_type, _ = detect_initial_breakout(hist, td)
                if initial_break and not is_yellow_light:
                    trigger_label = break_type
                    risk_calculator = NOCRiskManager(total_capital=cfg.TOTAL_CAPITAL)
                    defense_info = risk_calculator.get_position_and_defense(sym, close, hist, market_mode=local_market_mode, is_yellow_light=False)
                    stop_price = defense_info["defense_line"]
                    noc_state[sym] = StockState(status="HOLD", entry=close, trailing_stop=stop_price)
                    alert = "⚡【初升段起漲】放量突破關鍵價位，小注試單！"
                    action_plan_text = build_light_plan(sym, close, hist, manual_stop_price, local_market_mode)
                elif td.get("Monster_Breakout", False):
                    trigger_label = "🔥【旱地拔蔥】底部極端爆量，長紅突破季線！"
                    if not is_lightning and ("衰退" in fund_health or "警報" in fund_health):
                        alert = "🛡️【基本面攔截】營收 YoY 衰退，無情淘汰。"
                    elif trend_score < 0:
                        alert = "🛡️【趨勢攔截】長線多頭條件未滿足，拒絕追高。"
                    elif is_yellow_light:
                        alert = "🟡【黃燈強制攔截】大盤震盪洗盤，強制攔截新倉。"
                        action_plan_text = ""
                    else:
                        matrix_signal = chip_matrix_analyzer.analyze(hist, market_mode=local_market_mode)
                        td['Trend_Score'] = trend_score
                        if not is_high_quality_signal(hist, td, matrix_signal, local_market_mode):
                            logger.info(f"🔇 低品質訊號攔截 {sym} : {trigger_label}")
                            trigger_label = ""
                            alert = "📉 訊號品質不足 (未突破20日高點/量比<2/籌碼弱勢)"
                        else:
                            risk_calculator = NOCRiskManager(total_capital=cfg.TOTAL_CAPITAL)
                            defense_info = risk_calculator.get_position_and_defense(sym, close, hist, market_mode=local_market_mode, is_yellow_light=is_yellow_light)
                            stop_price = defense_info["defense_line"]
                            noc_state[sym] = StockState(status="HOLD", entry=close, trailing_stop=stop_price)
                            alert = "🐉【妖股起漲預警】資金強勢介入，無視基本面，強烈建議觀察試單！"
                            action_plan_text = build_tactical_plan(sym, close, hist, trend_score, fund_health, manual_stop_price, market_mode=local_market_mode)
                elif td.get("Sniper_Signal", False):
                    trigger_label = "🌟 狙擊金叉 (底部扭轉)"
                    if not is_lightning and ("衰退" in fund_health or "警報" in fund_health):
                        alert = "🛡️【基本面攔截】營收 YoY 衰退，無情淘汰。"
                    elif trend_score < 0:
                        alert = "🛡️【趨勢攔截】長線多頭條件未滿足，拒絕追高。"
                    elif is_yellow_light:
                        alert = "🟡【黃燈強制攔截】大盤進入震盪洗盤期，戰情室強制攔截，禁止盲目開新倉建倉。"
                        action_plan_text = ""
                    else:
                        matrix_signal = chip_matrix_analyzer.analyze(hist, market_mode=local_market_mode)
                        td['Trend_Score'] = trend_score
                        if not is_high_quality_signal(hist, td, matrix_signal, local_market_mode):
                            logger.info(f"🔇 低品質訊號攔截 {sym} : {trigger_label}")
                            trigger_label = ""
                            alert = "📉 訊號品質不足 (未突破20日高點/量比<2/籌碼弱勢)"
                        else:
                            risk_calculator = NOCRiskManager(total_capital=cfg.TOTAL_CAPITAL)
                            defense_info = risk_calculator.get_position_and_defense(sym, close, hist, market_mode=local_market_mode, is_yellow_light=is_yellow_light)
                            stop_price = defense_info["defense_line"]
                            noc_state[sym] = StockState(status="HOLD", entry=close, trailing_stop=stop_price)
                            alert = f"🚀【長線波段佈局觸發】"
                            action_plan_text = build_tactical_plan(sym, close, hist, trend_score, fund_health, manual_stop_price, market_mode=local_market_mode)

            # ------------------- 組裝推播訊息（明確化） -------------------
            if trigger_label:
                header = f"🎯 {name} ({sym}) —— {trigger_label}\n"
            else:
                header = f"🎯 {name} ({sym})\n"

            s = header
            s += f" 現價: {close:.2f} | RSI: {rsi:.1f} | 乖離: {bias:+.1f}%\n"
            s += f" 趨勢: {trend_status} | 估值 PE: {pe_str} | 營收 YoY: {yoy_label}\n"

            matrix_signal = chip_matrix_analyzer.analyze(hist, market_mode=local_market_mode)
            s += f" 換手: {turnover:.2f}% | 量比: {vol_ratio:.2f}倍 | 籌碼戰術: {matrix_signal}\n"
            s += f" 💰 法人動向: {chip_msg}\n"
            s += f" 📊 財報透視: {fund_health}\n"

            # 安全顯示單月營收年增率
            yoy_display = f"{yoy:.1f}%" if isinstance(yoy, (int, float)) else str(yoy)
            s += f" 📈 單月YoY: {yoy_display} | 累計描述: {fund_health}\n"

            if action_plan_text:
                s += f"{action_plan_text}\n"
            else:
                s += f" 👉 作戰指令: {alert}\n"

            action_command = s

            is_force_output = cat in force_include_categories
            if is_force_output:
                is_active = (close > ma20) or (vol_ratio > 1.5)
                if not is_active:
                    is_force_output = False
                    logger.debug(f"強制輸出分類 {cat} 中 {sym} 不活躍，降級過濾。")

            fatal_flaws = cfg.ACTION_BLACKLIST + ["攔截", "衰退", "警報", "無情淘汰", "拒絕追高", "黃燈強制攔截"]
            has_fatal_flaw = any(keyword in action_command for keyword in fatal_flaws)

            # 除錯模式：跳過所有過濾，直接推播
            if DEBUG_FORCE_PUSH:
                if tips:
                    s += f" 💡 Trello 決策提示: {tips}\n"
                cat_msg_list.append(s + "\n")
                generated_charts.append(draw_chart_if_needed(hist, sym))
                has_actionable_alerts = True
            else:
                if is_force_output:
                    if has_fatal_flaw:
                        logger.info(f"🛑 [強制分類攔截] {sym} 屬於強制輸出區，但觸發致命缺陷，強制封鎖推播。")
                    else:
                        if tips:
                            s += f" 💡 Trello 決策提示: {tips}\n"
                        cat_msg_list.append(s + "\n")
                        generated_charts.append(draw_chart_if_needed(hist, sym))
                        has_actionable_alerts = True
                else:
                    has_valid_signal = bool(trigger_label) or "主力點火" in matrix_signal
                    if has_valid_signal:
                        if has_fatal_flaw:
                            logger.info(f"🛑 [過濾器攔截] {sym} 雖有訊號，但觸發致命缺陷，強制封鎖推播。")
                        else:
                            if tips:
                                s += f" 💡 Trello 決策提示: {tips}\n"
                            cat_msg_list.append(s + "\n")
                            generated_charts.append(draw_chart_if_needed(hist, sym))
                            has_actionable_alerts = True
                    else:
                        logger.debug(f"🔇 [靜默跳過] {sym} 無重要觸發訊號，不推播。")

        if cat_msg_list:
            msg_list.append(f"━━━━━━━━━━━━━━\n📂 【{cat}】\n━━━━━━━━━━━━━━\n")
            msg_list.extend(cat_msg_list)

    # =========================================================================
    # 戰區 3：ETF 績效競技場 (週報)
    # =========================================================================
    if curr_dt.weekday() == 4:
        etf_arena = {"💰高股息防禦組": [], "🚀市值與主題成長組": []}
        current_year = curr_date.year
        for sym, name in all_symbols.items():
            etf_icon, _, _ = get_etf_strategy(sym, name)
            if "一般型" in etf_icon:
                continue
            hist = get_stock_data(sym, name)
            if hist is None or len(hist) < 10:
                continue
            close_price = hist["Close"].iloc[-1]
            qtr_days = min(60, len(hist) - 1)
            qtr_price = hist["Close"].iloc[-(qtr_days + 1)]
            qtr_roi = ((close_price - qtr_price) / qtr_price) * 100 if qtr_price else 0
            hist_ytd = hist[hist.index.year == current_year]
            ytd_roi = (((close_price - hist_ytd["Close"].iloc[0]) / hist_ytd["Close"].iloc[0]) * 100 if not hist_ytd.empty and hist_ytd["Close"].iloc[0] != 0 else qtr_roi)
            group_key = "💰高股息防禦組" if "高股息" in etf_icon else "🚀市值與主題成長組"
            etf_arena[group_key].append({"name": name, "sym": sym, "qtr_roi": qtr_roi, "ytd_roi": ytd_roi})
        if any(etf_arena.values()):
            msg_list.append("━━━━━━━━━━━━━━\n🏆 【ETF 雙引擎長線績效競技場 (週報)】\n━━━━━━━━━━━━━━\n")
            for group_name, group_data in etf_arena.items():
                if not group_data:
                    continue
                msg_list.append(f"**{group_name}**\n")
                for idx, etf in enumerate(sorted(group_data, key=lambda x: x["qtr_roi"], reverse=True)):
                    rank = ["🥇", "🥈", "🥉"][idx] if idx < 3 else "🔸"
                    status = "🔥 雙料強勢" if etf["qtr_roi"] > 5.0 and etf["ytd_roi"] > 10.0 else ("⏳ 長線沉澱修正" if etf["qtr_roi"] < 0 and etf["ytd_roi"] > 0 else ("⚠️ 績效嚴重落後" if etf["qtr_roi"] < -2.0 and etf["ytd_roi"] < 0 else "✅ 穩健向上跟隨"))
                    msg_list.append(f"{rank} {etf['name']} ({etf['sym']})\n 季度動能: {etf['qtr_roi']:+.1f}% ｜ 本年累計: {etf['ytd_roi']:+.1f}% ({status})\n")
                msg_list.append("\n")
    else:
        logger.info("非週五，跳過 ETF 績效推播 (週報模式)")

    # =========================================================================
    # 最終儲存與推播
    # =========================================================================
    if not has_data:
        logger.info("無有效標的運算數據，終止行程。")
        sys.exit(0)

    save_state(noc_state)

    if not has_actionable_alerts and cfg.SILENT_MODE:
        logger.info("🔇 [靜默模式] 今日無任何可行動警報（無建倉/停損/獲利巡航等重要事件），系統靜默退出。")
        sys.exit(0)

    send_reports(f"NOC 戰情報告 {curr_date}", f"📡 【NOC 終極戰情室 v16.12】\n📅 執行時間：{curr_time}\n━━━━━━━━━━━━━━\n" + "".join(msg_list), generated_charts)

    for chart in generated_charts:
        if Path(chart).exists():
            Path(chart).unlink()

    logger.info("🚀 全系統長線波段精算追蹤程序執行完畢，資料安全回存本地庫。")
