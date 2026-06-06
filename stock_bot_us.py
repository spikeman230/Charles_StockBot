# =============================================================================
# NOC 美股戰情室 v1.0 長短雙軌版
# 適用市場：美股 (NYSE/NASDAQ)
# 核心功能：初升段即時偵測、過熱攔截、白名單強制輸出、四象限矩陣
# 整合：旱地拔蔥、狙擊金叉
# 使用現有 Trello 面板，新增「美股長線觀測區」、「美股短線觀測區」列表
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

# 導入美股核心引擎
from noc_core_us import (
    get_stock_data, NOCStrategy_US, calculate_all_indicators,
    assess_volume_turnover_signal, is_overheated, detect_initial_breakout,
    calculate_monster_breakout, calculate_sniper_signal, is_high_quality_signal
)

# =============================================================================
# 初始化與組態
# =============================================================================
load_dotenv()

LOG_FILE = "noc_us_system.log"
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
TRELLO_KEY = os.getenv("TRELLO_KEY")
TRELLO_TOKEN = os.getenv("TRELLO_TOKEN")
TRELLO_BOARD_ID = os.getenv("TRELLO_BOARD_ID")

class Config:
    TOTAL_CAPITAL_USD : float = float(os.getenv("TOTAL_CAPITAL_USD", "50000"))
    RISK_PER_TRADE : float = float(os.getenv("RISK_PER_TRADE", "0.02"))
    ATR_MULTIPLIER : float = float(os.getenv("ATR_MULTIPLIER", "2.5")) # 美股波動較小
    SILENT_MODE : bool = os.getenv("SILENT_MODE", "false").lower() == "true"
    CACHE_TTL_MINUTES : int = int(os.getenv("CACHE_TTL_MINUTES", "30"))
    CACHE_MAX_ITEMS : int = int(os.getenv("CACHE_MAX_ITEMS", "200"))
    MAX_WORKERS : int = int(os.getenv("MAX_WORKERS", "6"))
    STATE_FILE : str = "noc_us_state.json"
    LOG_FILE_CSV : str = "noc_us_trading_log.csv"
    RADAR_FILE : str = "radar_us_targets.json"

    ACTION_WHITELIST : list = ["建倉", "試單", "波段", "佈局", "長線鎖籌", "加碼", "獲利巡航", "洗盤耐受", "戰術撤離"]
    ACTION_BLACKLIST : list = ["持股觀望", "暫停進場", "嚴格觀望", "不建議進場", "等待", "不動用資金"]

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
# 統一的數據獲取（使用 noc_core_us 完整版，並支援快取）
# =============================================================================
def get_stock_data_cached(symbol: str) -> Optional[pd.DataFrame]:
    cached = DATA_CACHE.get(symbol)
    if cached is not None:
        return cached
    hist = get_stock_data(symbol)
    if hist is not None:
        DATA_CACHE.set(symbol, hist)
    return hist

# =============================================================================
# 風險管理 (美股版)
# =============================================================================
class NOCRiskManager_US:
    def __init__(self, total_capital: float = 50000.0):
        self.total_capital = total_capital

    def calculate_atr(self, hist_df: pd.DataFrame, period: int = 14) -> float:
        if len(hist_df) < period + 1:
            return hist_df['Close'].iloc[-1] * 0.02
        hl = hist_df['High'] - hist_df['Low']
        hc = np.abs(hist_df['High'] - hist_df['Close'].shift())
        lc = np.abs(hist_df['Low'] - hist_df['Close'].shift())
        tr = pd.concat([hl, hc, lc], axis=1).max(axis=1)
        return tr.rolling(period).mean().iloc[-1]

    def get_position_and_defense(self, symbol: str, current_price: float, hist_df: pd.DataFrame = None,
                                 market_mode: str = "BULL", is_yellow_light: bool = False) -> dict:
        try:
            if hist_df is None or hist_df.empty:
                hist_df = yf.Ticker(symbol).history(period="6mo")
            atr = self.calculate_atr(hist_df, 14)
            mult = 2.0 if is_yellow_light else (1.5 if market_mode == "BULL" else 2.5)
            stop = current_price - (atr * mult)
            if not hist_df.empty and len(hist_df) >= 20:
                ma20 = hist_df['Close'].rolling(20).mean().iloc[-1]
                if not pd.isna(ma20):
                    stop = min(stop, ma20)
            risk_per_share = current_price - stop
            if risk_per_share <= 0:
                risk_per_share = current_price * 0.05
            max_shares = math.floor((self.total_capital * 0.02) / risk_per_share)
            max_alloc = math.floor((self.total_capital * 0.15) / current_price)
            total = min(max_shares, max_alloc)
            core = total // 2
            tactical = total - core
            return {
                "current_price": round(current_price, 2),
                "defense_line": round(stop, 2),
                "core_shares": int(core),
                "tactical_shares": int(tactical),
                "total_shares": int(total),
                "risk_per_share": round(risk_per_share, 2)
            }
        except Exception as e:
            logger.error(f"❌ 部位精算異常: {e}")
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
# Trello 整合（共用台股 Trello 面板，新增美股專區）
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
                if "NOC 美股系統狀態" in card["name"]:
                    status_card_id = card["id"]
                    break
            if status_card_id:
                break
        ny_tz = datetime.timezone(datetime.timedelta(hours=-4))
        date_str = datetime.datetime.now(ny_tz).strftime("%m/%d")
        new_name = f"{color} NOC 美股系統狀態：{status_msg} ({date_str})"
        if status_card_id:
            requests.put(f"https://api.trello.com/1/cards/{status_card_id}", params=_trello_params(name=new_name), timeout=10)
        else:
            requests.post("https://api.trello.com/1/cards", params=_trello_params(idList=first_list_id, name=new_name, pos="top"), timeout=10)
    except Exception as e:
        logger.error(f"Trello 看板更新失敗: {e}")

def update_trello_system_status_bg(status_msg: str, color: str = "🟢") -> None:
    threading.Thread(target=update_trello_system_status, args=(status_msg, color), daemon=True).start()

def _parse_card_to_stock(card: dict) -> Tuple[str, dict]:
    raw_name = card["name"].strip()
    ticker_match = re.match(r"^[A-Za-z0-9.]+", raw_name)
    symbol = ticker_match.group() if ticker_match else raw_name
    name_part = raw_name[len(symbol):].strip() if ticker_match else raw_name
    name = re.sub(r"\(.*?\)", "", name_part).strip() or symbol
    title_tip_match = re.search(r"\((.*?)\)", name_part)
    trello_tip = title_tip_match.group(1) if title_tip_match else card.get("desc", "").strip()
    return symbol, {"name": name, "trello_tip": trello_tip}

def _parse_card_to_portfolio(card: dict) -> Tuple[str, dict]:
    raw_name = card["name"].strip()
    ticker_match = re.match(r"^[A-Za-z0-9.]+", raw_name)
    symbol = ticker_match.group() if ticker_match else raw_name
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
        # 定義美股相關列表名稱（請在 Trello 中手動新增）
        us_watch_categories = ["美股長線觀測區", "美股短線觀測區", "美股庫藏股"]
        for lst in lists_data:
            list_name = lst["name"]
            is_portfolio_list = "庫存" in list_name or "庫藏" in list_name
            if list_name not in us_watch_categories and "美股" not in list_name:
                continue # 只處理美股專區
            for card in lst.get("cards", []):
                if "NOC" in card["name"]:
                    continue
                if is_portfolio_list:
                    sym, info = _parse_card_to_portfolio(card)
                    my_portfolio[sym] = info
                else:
                    sym, info = _parse_card_to_stock(card)
                    trello_dict.setdefault(list_name, {})[sym] = info
        return trello_dict, my_portfolio
    except Exception as e:
        logger.error(f"無法拉取 Trello 看板配置: {e}")
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
# 交易日感知（美股）
# =============================================================================
def is_trading_day_ny() -> bool:
    now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=-4)))
    if now.weekday() >= 5:
        return False
    try:
        spy = yf.Ticker("SPY").history(period="5d")
        if spy.empty:
            return True
        last_date = spy.index[-1].date()
        diff_days = (now.date() - last_date).days
        return diff_days <= 1
    except:
        return True

# =============================================================================
# 建倉計劃函數
# =============================================================================
def build_tactical_plan(symbol: str, close: float, hist: pd.DataFrame, trend_score: float, fund_health: str, manual_stop: float = 0.0, market_mode: str = "BULL") -> str:
    if "衰退" in fund_health:
        return f" 🛡️ 【基本面攔截】營收衰退，不予執行建倉！\n"
    if trend_score < 0:
        return f" 🛡️ 【趨勢面攔截】長線多頭條件未滿足\n"

    risk_calculator = NOCRiskManager_US(total_capital=cfg.TOTAL_CAPITAL_USD)
    defense_data = risk_calculator.get_position_and_defense(symbol, close, hist, market_mode=market_mode, is_yellow_light=False)
    stop_loss = defense_data["defense_line"]
    stop_reason = f"ATR風控底線 (倍數: {'1.5' if market_mode=='BULL' else '2.5'})"
    if manual_stop > 0:
        stop_loss = manual_stop
        stop_reason = "Trello 覆寫價"

    plan = (
        f" 👉 【美股長線波段指令】\n"
        f" * 建議底倉：{defense_data['core_shares']} 股\n"
        f" * 建議游擊：{defense_data['tactical_shares']} 股\n"
        f" * 移動防禦底線：{stop_loss:.2f} ({stop_reason})\n"
        f" * 鐵律：收盤跌破防線即撤離！\n"
    )
    return plan

def build_light_plan(symbol: str, close: float, hist: pd.DataFrame, manual_stop: float, market_mode: str) -> str:
    risk_calculator = NOCRiskManager_US(total_capital=cfg.TOTAL_CAPITAL_USD)
    defense_data = risk_calculator.get_position_and_defense(symbol, close, hist, market_mode=market_mode, is_yellow_light=False)
    stop_loss = defense_data["defense_line"]
    if manual_stop > 0:
        stop_loss = manual_stop
    return (
        f" 👉 【初升段試單指令】\n"
        f" * 試單股數：{defense_data['total_shares']} 股 (資金5-10%)\n"
        f" * 防線：{stop_loss:.2f}\n"
        f" * 三日未站穩即減碼\n"
    )
# =============================================================================
# 並行預載入
# =============================================================================
def preload_all_stocks(all_symbols: Dict[str, str]) -> None:
    logger.info(f"啟動並行預載，共計 {len(all_symbols)} 檔美股...")
    def _fetch(args):
        sym, name = args
        time.sleep(random.uniform(0.1, 0.5))
        return sym, get_stock_data_cached(sym)

    with ThreadPoolExecutor(max_workers=cfg.MAX_WORKERS) as executor:
        futures = {executor.submit(_fetch, item): item[0] for item in all_symbols.items()}
        for future in as_completed(futures):
            try:
                future.result()
            except Exception as e:
                logger.error(f"預載失敗 {futures[future]}: {e}")

# =============================================================================
# 圖表與推播
# =============================================================================
def draw_chart_if_needed(hist: pd.DataFrame, symbol: str) -> str:
    chart_file = f"{symbol}_us_chart.png"
    try:
        mc = mpf.make_marketcolors(up="green", down="red", edge="black", wick="black", volume="gray")
        mpf.plot(hist[-60:], type="candle", style=mpf.make_mpf_style(base_mpf_style="yahoo", marketcolors=mc), volume=True, mav=(5, 20, 60), title=f"{symbol} (US)", savefig=chart_file)
    except:
        try:
            mpf.plot(hist[-60:], type="candle", style="yahoo", volume=True, mav=(5, 20), title=f"{symbol}", savefig=chart_file)
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
    ny_tz = datetime.timezone(datetime.timedelta(hours=-4))
    curr_dt = datetime.datetime.now(ny_tz)
    curr_date, curr_time = curr_dt.date(), curr_dt.strftime("%Y-%m-%d %H:%M:%S")

    logger.info(f"NOC 美股戰情室 v1.0 啟動。時間：{curr_time} (美東)")

    strategy = NOCStrategy_US()
    msg_list = []

    macro_info = strategy.get_macro_status("SPY")
    is_yellow_light = False

    if macro_info["status"] == "🔴 紅燈":
        logger.warning("🚨 美股大盤空頭，停止建倉！")
        update_trello_system_status_bg("空頭防禦 (停止建倉)", "🔴")
        send_reports(f"NOC 美股防空警報 {curr_date}", f"大盤：{macro_info['status']} - {macro_info['desc']}\n停止新倉！", [])
        sys.exit(0)
    elif macro_info["status"] == "🟡 黃燈":
        logger.warning("🟡 美股黃燈，半倉/收緊防線")
        is_yellow_light = True
        update_trello_system_status_bg("黃燈防禦", "🟡")

    #if not is_trading_day_ny():
        #logger.info("美股休市，靜默休眠")
        #update_trello_system_status_bg("休市靜默", "🔴")
        #if curr_dt.hour <= 10:
            #send_reports(f"NOC 美股休市 {curr_date}", "美股休市，系統休眠", [])
        # 不退出，仍可測試

    if not is_yellow_light:
        update_trello_system_status_bg("交易日追蹤中", "🟢")

    TRELLO_DICT, TRELLO_PORTFOLIO = fetch_trello_deployment()
    STOCK_DICT = TRELLO_DICT if TRELLO_DICT else {}
    MY_PORTFOLIO = TRELLO_PORTFOLIO if TRELLO_PORTFOLIO else {}

    # 讀取雷達自動火種檔案（若有）
    if Path(cfg.RADAR_FILE).exists():
        try:
            with open(cfg.RADAR_FILE, "r", encoding="utf-8") as f:
                STOCK_DICT["👀 美股雷達自動火種"] = json.load(f)
        except Exception as e:
            logger.error(f"讀取雷達檔失敗: {e}")

    all_symbols = {sym: data.get("name", sym) for sym, data in MY_PORTFOLIO.items()}
    for stocks in STOCK_DICT.values():
        for sym, item in stocks.items():
            all_symbols[sym] = item.get("name", sym) if isinstance(item, dict) else item

    preload_all_stocks(all_symbols)

    # 大盤模式固定為 BULL（美股多頭），可依 SPY 月線判斷
    market_mode = "BULL" if macro_info["status"] == "🟢 綠燈" else "BEAR"
    logger.info(f"市場模式 => {market_mode}")

    noc_state = load_state()

    macro_msg = f"🌐 【美股大盤】：{macro_info['status']} | {macro_info['desc']}\n"
    if is_yellow_light:
        macro_msg += "⚠️ 【黃燈】半倉操作，防線收緊\n"

    msg_list = [macro_msg]
    generated_charts = []
    has_data = False
    has_actionable_alerts = False

    # =========================================================================
    # 戰區 1：庫藏股（美股白名單強制輸出）
    # =========================================================================
    if MY_PORTFOLIO:
        msg_list.append("━━━━━━━━━━━━━━\n💼 【美股庫藏股】\n━━━━━━━━━━━━━━\n")
        for sym, data in MY_PORTFOLIO.items():
            hist = get_stock_data_cached(sym)
            if hist is None:
                continue

            td, has_data = hist.iloc[-1], True
            curr_price = td["Close"]
            buy_price = data.get("buy_price", 0)
            roi_pct = ((curr_price - buy_price) / buy_price) * 100 if buy_price else 0

            if sym not in noc_state:
                noc_state[sym] = StockState()
            sym_state = noc_state[sym]

            ma20 = td["20MA"]
            ma60 = td["60MA"]
            turnover = td["Turnover_Rate"]
            vol_ratio = td["Volume_Ratio"]
            yoy = td.get("YoY", "N/A")

            # 停損計算（簡化）
            atr = td["ATR"] if not pd.isna(td.get("ATR", 0)) else curr_price * 0.02
            calculated_stop = curr_price - (atr * 2.0)
            calculated_stop = min(calculated_stop, ma20) if not pd.isna(ma20) else calculated_stop

            if sym_state.status != "REAL_HOLD":
                noc_state[sym] = StockState(status="REAL_HOLD", entry=buy_price, trailing_stop=calculated_stop)
                sym_state = noc_state[sym]

            if isinstance(yoy, (int, float)) and yoy < -15:
                pnl_alert = "💀 營收大幅衰退，清倉！"
                noc_state[sym] = StockState(status="NONE")
            elif roi_pct <= -15.0 or curr_price < ma60:
                pnl_alert = f"🩸 跌破防守，停損變現！"
                noc_state[sym] = StockState(status="NONE")
            elif roi_pct > 0 and curr_price > ma20:
                pnl_alert = f"🔥 獲利巡航中，防線上移"
            else:
                pnl_alert = f"🔍 中立觀察"

            silent_keywords = ["中立觀察", "獲利巡航"]
            is_silent = any(kw in pnl_alert for kw in silent_keywords)
            if is_silent and cfg.SILENT_MODE:
                logger.info(f"🔇 靜默模式跳過 {sym}")
            else:
                generated_charts.append(draw_chart_if_needed(hist, sym))
                inv_str = f"🎯 {data.get('name', sym)} ({sym})\n"
                inv_str += f" 現價: {curr_price:.2f} | 成本: {buy_price:.2f} | 損益: {roi_pct:+.1f}%\n"
                inv_str += f" 換手: {turnover:.2f}% | 量比: {vol_ratio:.2f}\n"
                inv_str += f" 📊 財報: {yoy}\n"
                inv_str += f" 👉 指令: {pnl_alert}\n\n"
                msg_list.append(inv_str)
                has_actionable_alerts = True

    # =========================================================================
    # 戰區 2：觀察區（白名單強制輸出）
    # =========================================================================
    force_include_categories = ["美股長線觀測區", "美股短線觀測區"]

    for cat, stocks in STOCK_DICT.items():
        if not stocks:
            continue

        cat_msg_list = []
        for sym, item in stocks.items():
            name = item.get("name", sym) if isinstance(item, dict) else item
            tips = item.get("trello_tip", "") if isinstance(item, dict) else ""

            hist = get_stock_data_cached(sym)
            if hist is None:
                continue

            td, has_data = hist.iloc[-1], True
            close, rsi, ma5, ma20 = td["Close"], td["RSI"], td["5MA"], td["20MA"]
            ma60 = td["60MA"]
            vol_ratio = td["Volume_Ratio"]
            turnover = td["Turnover_Rate"]
            price_position = td["Price_Position"] if not pd.isna(td["Price_Position"]) else 0.5
            pe = td.get("PE", "N/A")
            yoy = td.get("YoY", "N/A")

            bias = ((close - ma20) / ma20) * 100 if ma20 else 0
            trend_status = "🔥 多頭" if close > ma5 > ma20 else ("🧊 空頭" if close < ma5 < ma20 else "🔄 盤整")

            # 基本面
            fund_health = strategy.get_fundamental_health(sym)

            if sym not in noc_state:
                noc_state[sym] = StockState()
            sym_state = noc_state[sym]

            alert = "✅ 追蹤中"
            trigger_label = ""
            action_plan_text = ""

            # 黃燈攔截非白名單
            if is_yellow_light and cat not in force_include_categories:
                continue

            # 過熱攔截
            ma20_val = td['20MA'] if not pd.isna(td['20MA']) else 0
            ma60_val = td['60MA'] if not pd.isna(td['60MA']) else 0
            return_5d = td.get('Return_5D', 0)
            return_10d = td.get('Return_10D', 0)
            overheated, over_reason = is_overheated(close, ma20_val, ma60_val, return_5d, return_10d, price_position, vol_ratio)
            if overheated:
                logger.info(f"🛑 [過熱攔截] {sym}: {over_reason}")
                continue

            # 四象限信號
            market_cap = td.get('Market_Cap', 0)
            quadrant_signal = assess_volume_turnover_signal(
                vol_ratio=vol_ratio,
                turnover=turnover,
                market_cap=market_cap,
                price_position=price_position,
                candle_ratio=td['Candle_Ratio'],
                is_red=td['Is_Red'],
                close_vs_high=td['Close_vs_High']
            )
            danger = ("🔴 主力出貨區", "⚠️ 量價背離陷阱", "🔴 爆量長上影", "⚠️ 黑K出量")
            if quadrant_signal in danger:
                continue

            # 狀態機觸發
            if sym_state.status == "NONE":
                initial_break, break_type, _ = detect_initial_breakout(hist, td)
                if initial_break and not is_yellow_light:
                    trigger_label = break_type
                    risk_calculator = NOCRiskManager_US(total_capital=cfg.TOTAL_CAPITAL_USD)
                    defense_info = risk_calculator.get_position_and_defense(sym, close, hist, market_mode=market_mode, is_yellow_light=False)
                    stop_price = defense_info["defense_line"]
                    noc_state[sym] = StockState(status="HOLD", entry=close, trailing_stop=stop_price)
                    alert = "⚡【初升段起漲】小注試單！"
                    action_plan_text = build_light_plan(sym, close, hist, 0, market_mode)
                elif td.get("Monster_Breakout", False):
                    trigger_label = "🔥【旱地拔蔥】"
                    alert = "🐉 爆量突破季線"
                    action_plan_text = build_tactical_plan(sym, close, hist, 1, fund_health, 0, market_mode)
                elif td.get("Sniper_Signal", False):
                    trigger_label = "🌟【狙擊金叉】"
                    alert = "🚀 底部扭轉"
                    action_plan_text = build_tactical_plan(sym, close, hist, 1, fund_health, 0, market_mode)

            # 組裝訊息
            s = f"🎯 {name} ({sym})"
            if trigger_label:
                s += f" —— {trigger_label}"
            s += "\n"
            s += f" 現價: {close:.2f} | RSI: {rsi:.1f} | 乖離: {bias:+.1f}%\n"
            s += f" 趨勢: {trend_status} | PE: {pe} | YoY: {yoy}\n"
            s += f" 換手: {turnover:.2f}% | 量比: {vol_ratio:.2f}\n"
            s += f" 📊 財報: {fund_health}\n"
            s += f" 📐 量價四象限: {quadrant_signal}\n"

            if action_plan_text:
                s += f"{action_plan_text}"
            else:
                s += f" 👉 指令: {alert}\n"

            is_force_output = cat in force_include_categories
            if is_force_output:
                if tips:
                    s += f" 💡 提示: {tips}\n"
                cat_msg_list.append(s)
                generated_charts.append(draw_chart_if_needed(hist, sym))
                has_actionable_alerts = True
            else:
                has_valid_signal = bool(trigger_label)
                if has_valid_signal:
                    if tips:
                        s += f" 💡 提示: {tips}\n"
                    cat_msg_list.append(s)
                    generated_charts.append(draw_chart_if_needed(hist, sym))
                    has_actionable_alerts = True

        if cat_msg_list:
            msg_list.append(f"━━━━━━━━━━━━━━\n📂 【{cat}】\n━━━━━━━━━━━━━━\n")
            msg_list.extend(cat_msg_list)

    # =========================================================================
    # 戰區 3：簡易績效追蹤（可選）
    # =========================================================================
    # 每週五推播主要 ETF 表現
    if curr_dt.weekday() == 4:
        etf_list = ["SPY", "QQQ", "IWM", "AAPL", "MSFT", "NVDA"]
        etf_msg = "━━━━━━━━━━━━━━\n🏆 美股主要標的週報\n━━━━━━━━━━━━━━\n"
        for etf in etf_list:
            try:
                hist = yf.Ticker(etf).history(period="5d")
                if not hist.empty:
                    close = hist['Close'].iloc[-1]
                    week_ago = hist['Close'].iloc[0] if len(hist) >= 5 else close
                    change = ((close - week_ago) / week_ago) * 100
                    etf_msg += f"{etf}: {close:.2f} | 週漲跌 {change:+.1f}%\n"
            except:
                pass
        msg_list.append(etf_msg + "\n")

    # =========================================================================
    # 最終儲存與推播
    # =========================================================================
    if not has_data:
        logger.info("無有效標的，終止")
        sys.exit(0)

    save_state(noc_state)

    if not has_actionable_alerts and cfg.SILENT_MODE:
        logger.info("靜默模式：無重要警報")
        sys.exit(0)

    send_reports(f"NOC 美股戰情報告 {curr_date}", f"📡 【NOC 美股戰情室 v1.0】\n📅 {curr_time}\n" + "".join(msg_list), generated_charts)

    for chart in generated_charts:
        if Path(chart).exists():
            Path(chart).unlink()

    logger.info("美股戰情室執行完畢")
