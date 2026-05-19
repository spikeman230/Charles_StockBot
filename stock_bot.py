# =============================================================================
# NOC 終極戰情室 v14.0 - 長線波段鎖籌重構版
# 核心戰術：
# 1. 強制掛載 Trend_Score 趨勢濾網，60MA 向上方可進場。
# 2. 基本面 YoY 成長檢查，虧損企業禁止長線建倉。
# 3. ATR_MULTIPLIER 鎖定 3.0，確保波段持有空間。
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
from dotenv import load_dotenv
from typing import Optional, Dict, Tuple, Any
from pathlib import Path

# 引入 NOC 核心防禦模組 (需確保 noc_core.py 位於同層目錄)
from noc_core import NOCDatabase, NOCStrategy, NOCDataFetcher, NOCRiskManager

# =============================================================================
# === 0. 初始化：載入環境變數與日誌 ===
# =============================================================================
load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(funcName)s - %(message)s",
    handlers=[
        logging.FileHandler("noc_system.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# =============================================================================
# === 1. 全域組態設定 ===
# =============================================================================
class Config:
    TOTAL_CAPITAL      : float = float(os.getenv("TOTAL_CAPITAL", "130000"))
    ATR_MULTIPLIER     : float = float(os.getenv("ATR_MULTIPLIER", "3.0"))  # 波段防禦寬容度
    YOY_EXPLOSION_PCT  : float = float(os.getenv("YOY_EXPLOSION_PCT", "10.0"))
    PE_LIMIT           : float = float(os.getenv("PE_LIMIT", "40.0"))
    MAX_WORKERS        : int   = int(os.getenv("MAX_WORKERS", "6"))
    STATE_FILE         : str   = "noc_state.json"
    LOG_FILE_CSV       : str   = "noc_trading_log.csv"

cfg = Config()
risk_mgr = NOCRiskManager(total_capital=cfg.TOTAL_CAPITAL)

# =============================================================================
# === 2. 長線決策引擎 (波段轉型) ===
# =============================================================================
def build_tactical_plan(symbol: str, close: float, hist: pd.DataFrame, trend_score: float, fund_health: str) -> str:
    """
    軍規級部署建議：廢除目標價，改以移動防禦線作為唯一出場依據
    """
    # 檢查基本面護城河
    if "禁止" in fund_health:
        return f"   🛡️ {fund_health}"
    
    # 檢查長線趨勢得分
    if trend_score < 0:
        return "   🛡️【趨勢警戒】60MA 趨勢向下，禁止長線佈局。"

    # 取得雙軌部位建議 (15% 總兵力)
    defense_data = risk_mgr.get_position_and_defense(symbol, close, hist)
    
    plan = (
        f"   💎【長線鎖籌 (波段模式)】\n"
        f"      * 移動防禦線 (Trailing Stop): {defense_data['defense_line']:.2f}\n"
        f"      * 建議長線底倉 (7.5%): {defense_data['core_shares']} 股\n"
        f"      * 建議短線游擊 (7.5%): {defense_data['tactical_shares']} 股\n"
        f"      * 風險係數 (ATR): {defense_data['risk_per_share']:.2f} / 股\n"
        f"      * 鐵律聲明: 跌破防禦線即刻執行總司令戰術撤離，禁止逆勢攤平！"
    )
    return plan

# =============================================================================
# === 3. 主程式架構 ===
# =============================================================================
if __name__ == "__main__":
    logger.info("NOC 終極戰情室 v14.0 (波段鎖籌版) 啟動...")
    
    # 初始化核心物件
    db = NOCDatabase()
    strategy = NOCStrategy()
    
    # 大盤風向儀檢查
    macro_status = strategy.get_macro_status()
    if macro_status["status"] == "🔴 紅燈":
        logger.warning("🔴 [致命警報] 觸發拔插頭協議，系統進入資產保護模式。")
        sys.exit(0)
        
    logger.info(f"大盤風向儀狀態: {macro_status['status']} - {macro_status['desc']}")
    
    # 執行波段掃描與監控邏輯
    # (註：此處為架構框架，詳細的迴圈與雷達邏輯會與您原有的資料庫對接)
    logger.info("戰情室已切換為「長線鎖籌」模式，準備進行波段掃描與庫存部位監控。")
    
    # 範例調用：確保與核心策略對接
    # plan = build_tactical_plan(symbol, close, hist, trend_score, fund_health)
    # logger.info(plan)
    
    logger.info("重構部署完成。系統已就緒，等待指令進行下一步作業。")
