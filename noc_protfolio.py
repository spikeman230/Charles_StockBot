"""
NOC 投資組合軍需官 (Portfolio Quartermaster)
獨立程式，專屬資料庫 noc_portfolio.db，不影響 noc_warroom.db。
功能：
1. 建立交易總帳 (trade_ledger)
2. 從 Trello 同步庫存 (新增 / 平倉)
3. 計算未實現損益與 ATR 防線
4. 生成並推播每日後勤戰報 (Telegram)
"""

import os
import sys
import re
import logging
import datetime
import sqlite3
from typing import List, Dict, Optional, Any

import requests
import yfinance as yf
import pandas as pd
import numpy as np
from dotenv import load_dotenv

# 從 noc_core 導入風險管理器（不涉及資料庫寫入）
try:
    from noc_core import NOCRiskManager
except ImportError:
    # 若 noc_core 不在路徑，則提供一個簡易替代（不推薦，但可防止崩潰）
    logging.warning("無法從 noc_core 導入 NOCRiskManager，使用內建簡易 ATR 計算")
    class NOCRiskManager:
        def calculate_atr(self, hist_df: pd.DataFrame, period: int = 14) -> float:
            if len(hist_df) < period + 1:
                return hist_df['Close'].iloc[-1] * 0.025
            hl = hist_df['High'] - hist_df['Low']
            hc = np.abs(hist_df['High'] - hist_df['Close'].shift())
            lc = np.abs(hist_df['Low'] - hist_df['Close'].shift())
            tr = pd.concat([hl, hc, lc], axis=1).max(axis=1)
            return tr.rolling(period).mean().iloc[-1]

# ========== 環境變數載入 ==========
load_dotenv()

TRELLO_API_KEY = os.getenv("TRELLO_KEY")
TRELLO_TOKEN = os.getenv("TRELLO_TOKEN")
TRELLO_BOARD_ID = os.getenv("TRELLO_BOARD_ID")
TRELLO_LIST_NAME = os.getenv("TRELLO_LIST_NAME", "💼 庫存機櫃")
TELEGRAM_BOT_TOKEN = os.getenv("TG_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TG_CHAT_ID")

DB_PATH = "noc_portfolio.db" # 專屬資料庫

# ========== 日誌設定 ==========
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger(__name__)

# ========== 資料庫初始化 ==========
def init_db() -> None:
    """建立 noc_portfolio.db 與 trade_ledger 表格"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS trade_ledger (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            symbol TEXT NOT NULL,
            name TEXT,
            status TEXT CHECK(status IN ('OPEN', 'CLOSED')) NOT NULL,
            entry_date TEXT,
            exit_date TEXT,
            entry_price REAL,
            exit_price REAL,
            shares INTEGER,
            realized_pnl REAL
        )
    ''')
    conn.commit()
    conn.close()
    logger.info(f"資料庫初始化完成: {DB_PATH}")

# ========== Trello 資料獲取 ==========
def fetch_trello_deployment() -> List[Dict[str, Any]]:
    """
    從 Trello 指定列表讀取庫存卡片。
    卡片名稱視為股票代號 (例如 "2330.TW")，
    卡片描述須包含 成本:123.45 與 股數:1000 (不區分大小寫，支援冒號或等號)。
    回傳: [{'symbol': str, 'name': str, 'entry_price': float, 'shares': int}, ...]
    """
    if not TRELLO_API_KEY or not TRELLO_TOKEN or not TRELLO_BOARD_ID:
        logger.error("Trello 環境變數未完整設定，無法獲取庫存")
        return []

    # 1. 取得所有列表，找出目標列表 ID
    lists_url = f"https://api.trello.com/1/boards/{TRELLO_BOARD_ID}/lists"
    params = {'key': TRELLO_API_KEY, 'token': TRELLO_TOKEN}
    try:
        resp = requests.get(lists_url, params=params, timeout=10)
        resp.raise_for_status()
        lists = resp.json()
    except Exception as e:
        logger.error(f"無法取得 Trello 列表: {e}")
        return []

    target_list = next((lst for lst in lists if lst['name'] == TRELLO_LIST_NAME), None)
    if not target_list:
        logger.error(f"找不到 Trello 列表: {TRELLO_LIST_NAME}")
        return []

    list_id = target_list['id']

    # 2. 取得該列表的所有卡片
    cards_url = f"https://api.trello.com/1/lists/{list_id}/cards"
    try:
        resp = requests.get(cards_url, params=params, timeout=10)
        resp.raise_for_status()
        cards = resp.json()
    except Exception as e:
        logger.error(f"無法取得 Trello 卡片: {e}")
        return []

    result = []
    for card in cards:
        name = card['name'].strip()
        if not name:
            continue
        desc = card.get('desc', '') or ''
        # 解析 成本 與 股數 (忽略大小寫，支援冒號或等號)
        cost_match = re.search(r'成本\s*[:=]\s*([\d.]+)', desc, re.IGNORECASE)
        shares_match = re.search(r'股數\s*[:=]\s*([\d]+)', desc, re.IGNORECASE)
        try:
            entry_price = float(cost_match.group(1)) if cost_match else 0.0
            shares = int(shares_match.group(1)) if shares_match else 0
        except (ValueError, AttributeError):
            logger.warning(f"卡片 {name} 描述格式錯誤，跳過")
            continue

        if entry_price <= 0 or shares <= 0:
            logger.warning(f"卡片 {name} 成本或股數無效，跳過")
            continue

        result.append({
            'symbol': name, # 直接使用卡片名稱作為股票代號
            'name': name,
            'entry_price': entry_price,
            'shares': shares
        })
        logger.debug(f"從 Trello 讀取: {name} 成本 {entry_price} 股數 {shares}")

    logger.info(f"共讀取 {len(result)} 筆 Trello 庫存")
    return result

# ========== 資料庫輔助函數 ==========
def get_open_positions(conn: sqlite3.Connection) -> List[Dict[str, Any]]:
    """讀取所有 OPEN 狀態的紀錄"""
    c = conn.cursor()
    c.execute("""
        SELECT id, symbol, name, entry_date, entry_price, shares
        FROM trade_ledger
        WHERE status = 'OPEN'
    """)
    rows = c.fetchall()
    return [
        {
            'id': row[0],
            'symbol': row[1],
            'name': row[2],
            'entry_date': row[3],
            'entry_price': row[4],
            'shares': row[5]
        }
        for row in rows
    ]

def get_closed_today(conn: sqlite3.Connection) -> List[Dict[str, Any]]:
    """讀取今日平倉的紀錄"""
    today = datetime.date.today().isoformat()
    c = conn.cursor()
    c.execute("""
        SELECT symbol, name, exit_price, realized_pnl
        FROM trade_ledger
        WHERE status = 'CLOSED' AND exit_date = ?
    """, (today,))
    rows = c.fetchall()
    return [
        {'symbol': row[0], 'name': row[1], 'exit_price': row[2], 'realized_pnl': row[3]}
        for row in rows
    ]

# ========== 同步邏輯 ==========
def sync_trello_positions(conn: sqlite3.Connection, trello_positions: List[Dict[str, Any]]) -> None:
    """
    比對 Trello 與資料庫 OPEN 紀錄，進行新增或平倉。
    """
    open_pos = get_open_positions(conn)
    open_symbols = {p['symbol'] for p in open_pos}
    trello_symbols = {p['symbol'] for p in trello_positions}

    # 新增：Trello 有但 OPEN 沒有的
    new_symbols = trello_symbols - open_symbols
    for p in trello_positions:
        if p['symbol'] in new_symbols:
            c = conn.cursor()
            today = datetime.date.today().isoformat()
            c.execute("""
                INSERT INTO trade_ledger
                (symbol, name, status, entry_date, entry_price, shares)
                VALUES (?, ?, 'OPEN', ?, ?, ?)
            """, (p['symbol'], p['name'], today, p['entry_price'], p['shares']))
            logger.info(f"新增持倉: {p['symbol']} 成本 {p['entry_price']} 股數 {p['shares']}")
    conn.commit()

    # 平倉：OPEN 有但 Trello 沒有的
    closed_symbols = open_symbols - trello_symbols
    for p in open_pos:
        if p['symbol'] in closed_symbols:
            # 取得前一個交易日收盤價
            exit_price = p['entry_price'] # fallback
            try:
                ticker = yf.Ticker(p['symbol'])
                hist = ticker.history(period="2d")
                if len(hist) >= 2:
                    exit_price = hist['Close'].iloc[-2] # 前一日
                elif len(hist) == 1:
                    exit_price = hist['Close'].iloc[-1] # 只有今日？但可能不完整
                else:
                    logger.warning(f"{p['symbol']} 無足夠歷史資料，使用成本價平倉")
            except Exception as e:
                logger.error(f"取得 {p['symbol']} 收盤價失敗: {e}")

            realized_pnl = (exit_price - p['entry_price']) * p['shares']
            today = datetime.date.today().isoformat()
            c = conn.cursor()
            c.execute("""
                UPDATE trade_ledger
                SET status = 'CLOSED',
                    exit_date = ?,
                    exit_price = ?,
                    realized_pnl = ?
                WHERE id = ?
            """, (today, exit_price, realized_pnl, p['id']))
            logger.info(f"平倉: {p['symbol']} 出場價 {exit_price:.2f} 損益 {realized_pnl:.2f}")
    conn.commit()

# ========== 未實現損益與 ATR 防線計算 ==========
def calculate_open_positions(conn: sqlite3.Connection) -> List[Dict[str, Any]]:
    """
    針對所有 OPEN 持倉，計算目前市價、未實現損益、ATR 防線與狀態。
    """
    open_pos = get_open_positions(conn)
    if not open_pos:
        return []

    risk_mgr = NOCRiskManager()
    results = []

    for p in open_pos:
        try:
            ticker = yf.Ticker(p['symbol'])
            hist = ticker.history(period="60d")
            if hist.empty:
                logger.warning(f"{p['symbol']} 無歷史資料，跳過")
                continue

            current_price = hist['Close'].iloc[-1]
            # 未實現損益
            unrealized_pnl = (current_price - p['entry_price']) * p['shares']
            unrealized_pnl_pct = (current_price / p['entry_price'] - 1) * 100

            # ATR 防線 (2倍ATR)
            atr = risk_mgr.calculate_atr(hist, period=14)
            defense_line = current_price - 2.0 * atr

            # 狀態判定
            if current_price > defense_line:
                if current_price - defense_line > atr:
                    status = "✅ 安全"
                else:
                    status = "⚠️ 警戒"
            else:
                status = "🔴 跌破"

            results.append({
                'symbol': p['symbol'],
                'name': p['name'],
                'entry_price': p['entry_price'],
                'current_price': current_price,
                'shares': p['shares'],
                'unrealized_pnl': unrealized_pnl,
                'unrealized_pnl_pct': unrealized_pnl_pct,
                'defense_line': defense_line,
                'status': status
            })
        except Exception as e:
            logger.error(f"計算 {p['symbol']} 時發生錯誤: {e}")

    return results

# ========== Telegram 推播 ==========
def send_telegram(message: str) -> None:
    """透過 Telegram Bot 發送訊息"""
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        logger.warning("Telegram 環境變數未設定，跳過推播")
        return

    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = {
        'chat_id': TELEGRAM_CHAT_ID,
        'text': message,
        'parse_mode': 'HTML'
    }
    try:
        resp = requests.post(url, data=payload, timeout=10)
        resp.raise_for_status()
        logger.info("Telegram 訊息推送成功")
    except Exception as e:
        logger.error(f"Telegram 推送失敗: {e}")

# ========== 戰報生成 ==========
def generate_report(open_positions: List[Dict[str, Any]],
                    closed_today: List[Dict[str, Any]]) -> str:
    """組裝每日後勤戰報文字"""
    today = datetime.date.today().isoformat()
    lines = [
        f"📊 軍需官每日庫存戰報 ({today})",
        "━━━━━━━━━━━━━━━━━━━━━━━━━━"
    ]

    # 總投入與未實現總損益
    total_invested = sum(p['entry_price'] * p['shares'] for p in open_positions)
    total_unrealized = sum(p['unrealized_pnl'] for p in open_positions)
    pct = (total_unrealized / total_invested * 100) if total_invested else 0.0
    lines.append(f"💰 總投入資金: {total_invested:,.2f}")
    lines.append(f"📈 未實現總損益: {total_unrealized:+,.2f} ({pct:+.2f}%)")
    lines.append("")

    # 進行中戰役
    lines.append("⚔️ 進行中戰役:")
    if open_positions:
        for p in open_positions:
            lines.append(
                f" {p['symbol']} {p['name']} | "
                f"成本 {p['entry_price']:.2f} | "
                f"現價 {p['current_price']:.2f} | "
                f"損益 {p['unrealized_pnl']:+,.2f} ({p['unrealized_pnl_pct']:+.2f}%) | "
                f"防線: {p['status']}"
            )
    else:
        lines.append(" 目前無 OPEN 持倉")
    lines.append("")

    # 今日結算戰果
    lines.append("🎯 今日結算戰果:")
    if closed_today:
        for c in closed_today:
            lines.append(
                f" {c['symbol']} {c['name']} | "
                f"出場 {c['exit_price']:.2f} | "
                f"損益 {c['realized_pnl']:+,.2f}"
            )
    else:
        lines.append(" 無平倉紀錄")

    return "\n".join(lines)

# ========== 主程式 ==========
def main() -> None:
    """執行完整流程"""
    logger.info("=== NOC 投資組合軍需官啟動 ===")

    # 1. 初始化資料庫
    init_db()

    conn = sqlite3.connect(DB_PATH)
    try:
        # 2. 從 Trello 獲取最新庫存
        trello_positions = fetch_trello_deployment()

        # 3. 同步資料庫 (新增/平倉)
        sync_trello_positions(conn, trello_positions)

        # 4. 計算 OPEN 持倉的未實現損益與防線
        open_positions = calculate_open_positions(conn)

        # 5. 取得今日平倉紀錄
        closed_today = get_closed_today(conn)

        # 6. 生成戰報
        report = generate_report(open_positions, closed_today)
        logger.info("\n" + report)

        # 7. 推播 Telegram
        send_telegram(report)

    except Exception as e:
        logger.exception(f"執行過程發生未預期錯誤: {e}")
    finally:
        conn.close()
        logger.info("=== 投資組合軍需官執行完畢 ===")

if __name__ == "__main__":
    main()

