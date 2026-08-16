# -*- coding: utf-8 -*-
"""
DUAL QLD + TQQQ Anchor Dashboard v8 - Google Colab 실행용

목적
- QLD 최종 앵커와 TQQQ 최종 앵커를 한 파일에서 동시에 판단
- 사용자가 QLD/TQQQ 목표비중을 수동으로 조절
- 현금, QLD 보유수량/평단, TQQQ 보유수량/평단 수동 입력
- 각 자산별 BUY / SELL_ALL / HOLD / WAIT 판단
- 목표비중 대비 매수 가이드: 매도는 매도 4조건에서만 실행
- Excel 리포트, 통합 PNG 현황판, QLD/TQQQ 단독 현황판, 오늘의 주문요약 PNG, 그래프 PNG, CSV 요약 생성

최종 전략 요약
[QLD 매수]
1) QLD > QLD MA20
2) QLD MACD Histogram > 0
3) QQQ > QQQ MA50

[QLD 매도]
1) QQQ < QQQ MA200
2) QLD < QLD MA100
3) QLD MACD Histogram < 0
4) QLD 위험점수 >= 8

[TQQQ 매수]
1) TQQQ > TQQQ MA20
2) TQQQ MACD Histogram > 0
3) QQQ > QQQ MA50

[TQQQ 매도]
1) QQQ < QQQ MA200
2) TQQQ < TQQQ MA100
3) TQQQ MACD Histogram < 0
4) TQQQ 위험점수 >= 8

실행 예시
    python DUAL_QLD_TQQQ_anchor_dashboard.py
    python DUAL_QLD_TQQQ_anchor_dashboard.py --cash 10000 --qld-weight 0.6 --tqqq-weight 0.4
    python DUAL_QLD_TQQQ_anchor_dashboard.py --cash 5000 --qld-shares 20 --qld-avg-price 80 --tqqq-shares 10 --tqqq-avg-price 60 --qld-weight 0.6 --tqqq-weight 0.4 --period 15y

주의
- 투자 조언이 아니라 사용자가 정한 규칙 기반 신호 계산 도구입니다.
- 실제 주문 전 가격, 세금, 수수료, 계좌 상황을 직접 확인하세요.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import platform
import subprocess
import sys
from datetime import datetime
from typing import Dict, List, Tuple, Optional


def ensure_package(package: str, import_name: Optional[str] = None) -> None:
    name = import_name or package
    try:
        __import__(name)
    except ModuleNotFoundError:
        print(f"[install] {package} 설치 중...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])


for pkg, imp in [
    ("pandas", "pandas"),
    ("numpy", "numpy"),
    ("yfinance", "yfinance"),
    ("xlsxwriter", "xlsxwriter"),
    ("matplotlib", "matplotlib"),
    ("pillow", "PIL"),
    ("requests", "requests"),
]:
    ensure_package(pkg, imp)

import numpy as np
import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt
from PIL import Image, ImageDraw, ImageFont
import requests


PATCH_VERSION = "2026-07-01-dual-anchor-v14-live-execution-brake"
STRATEGY_NAME = "DUAL_QLD_TQQQ_FINAL_ANCHOR"
OUTPUT_DIR_DEFAULT = "dual_anchor_v13_ai_mine_results"
DOWNLOAD_PERIOD_DEFAULT = "15y"
FULL_SELL_SCORE = 8
SELL_MA = 100


# ======================================================
# User settings - Telegram
# ======================================================
# 여기에 텔레그램 정보를 한 번만 입력해두면 실행할 때마다 명령어에 토큰/채팅ID를
# 길게 적지 않아도 됩니다.
#
# 사용 방법
# 1) BotFather에서 받은 봇 토큰을 TELEGRAM_BOT_TOKEN_IN_FILE에 입력
# 2) 내 텔레그램 chat_id를 TELEGRAM_CHAT_ID_IN_FILE에 입력
# 3) 매번 자동 전송하려면 TELEGRAM_SEND_DEFAULT = False 로 변경
#
# 보안 주의: 이 파일을 다른 사람에게 공유하거나 깃허브에 올릴 때는
# 반드시 아래 토큰과 chat_id를 지우세요.
TELEGRAM_SEND_DEFAULT = True
TELEGRAM_BOT_TOKEN_IN_FILE = ""  # 보안상 비워둠. 환경변수 TELEGRAM_BOT_TOKEN 또는 실행 인자로 입력
TELEGRAM_CHAT_ID_IN_FILE = ""    # 보안상 비워둠. 환경변수 TELEGRAM_CHAT_ID 또는 실행 인자로 입력




# ======================================================
# Google Colab quick start
# ======================================================
# 1) Colab 왼쪽 파일 영역에 이 .py 파일을 업로드합니다.
# 2) 아래 예시처럼 실행합니다.
#    !python DUAL_QLD_TQQQ_anchor_dashboard_v8_realtime_performance_colab.py --cash 10000 --qld-weight 0.6 --tqqq-weight 0.4
# 3) 실제 체결 기록까지 남길 때 예시:
#    !python DUAL_QLD_TQQQ_anchor_dashboard_v8_realtime_performance_colab.py --cash 5000 --qld-shares 20 --qld-avg-price 80 --record-trade --qld-exec-action BUY --qld-exec-shares 2 --qld-exec-price 85
# 4) 결과물은 /content/dual_anchor_output 폴더에 생성됩니다.

# ======================================================
# Utility
# ======================================================

def ensure_output_dir(output_dir: str) -> str:
    os.makedirs(output_dir, exist_ok=True)
    return output_dir


def get_font_path(bold: bool = False) -> Optional[str]:
    candidates: List[str] = []
    if platform.system() == "Windows":
        candidates += [
            r"C:\Windows\Fonts\malgunbd.ttf" if bold else r"C:\Windows\Fonts\malgun.ttf",
            r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf",
        ]
    else:
        candidates += [
            "/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf" if bold else "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return None


def get_font(size: int, bold: bool = False):
    path = get_font_path(bold)
    if path:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            pass
    return ImageFont.load_default()


def money(x) -> str:
    try:
        return f"${float(x):,.2f}"
    except Exception:
        return str(x)


def num(x, digits: int = 2) -> str:
    try:
        return f"{float(x):,.{digits}f}"
    except Exception:
        return str(x)


def pct_value(x) -> str:
    try:
        return f"{float(x) * 100:.2f}%"
    except Exception:
        return str(x)


def pct_plain(x) -> str:
    try:
        return f"{float(x):.2f}%"
    except Exception:
        return str(x)


def yes_no(v) -> str:
    return "충족" if bool(v) else "미충족"


def draw_rounded_bar(draw: ImageDraw.ImageDraw, xy, radius: int, fill: str, outline=None, width: int = 1):
    draw.rounded_rectangle(xy, radius=radius, fill=fill, outline=outline, width=width)


def wrap_lines(text: str, max_chars: int) -> List[str]:
    text = str(text)
    lines: List[str] = []
    for paragraph in text.split("\n"):
        paragraph = paragraph.strip()
        if not paragraph:
            lines.append("")
            continue
        while len(paragraph) > max_chars:
            cut = paragraph.rfind(" ", 0, max_chars)
            if cut < max_chars // 2:
                cut = max_chars
            lines.append(paragraph[:cut].strip())
            paragraph = paragraph[cut:].strip()
        if paragraph:
            lines.append(paragraph)
    return lines


# ======================================================
# Data and indicators
# ======================================================

def download_ohlcv(ticker: str, period: str) -> pd.DataFrame:
    df = yf.download(ticker, period=period, interval="1d", auto_adjust=True, progress=False, threads=False)
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)
    if df.empty:
        raise RuntimeError(f"{ticker} 데이터 다운로드 실패")
    return df[["Open", "High", "Low", "Close", "Volume"]].dropna()


def rsi(series: pd.Series, n: int = 14) -> pd.Series:
    delta = series.diff()
    gain = delta.clip(lower=0)
    loss = -delta.clip(upper=0)
    avg_gain = gain.rolling(n).mean()
    avg_loss = loss.rolling(n).mean()
    rs = avg_gain / avg_loss.replace(0, np.nan)
    return 100 - (100 / (1 + rs))


def macd(series: pd.Series, fast: int = 12, slow: int = 26, signal: int = 9):
    ema_fast = series.ewm(span=fast, adjust=False).mean()
    ema_slow = series.ewm(span=slow, adjust=False).mean()
    macd_line = ema_fast - ema_slow
    signal_line = macd_line.ewm(span=signal, adjust=False).mean()
    hist = macd_line - signal_line
    return macd_line, signal_line, hist


def add_indicators(data: pd.DataFrame) -> pd.DataFrame:
    for prefix in ["QQQ", "QLD", "TQQQ"]:
        close = data[f"{prefix}_Close"]
        high = data[f"{prefix}_High"]
        low = data[f"{prefix}_Low"]
        vol = data[f"{prefix}_Volume"]

        for n in [5, 10, 20, 50, 100, 120, 150, 200]:
            data[f"{prefix}_MA{n}"] = close.rolling(n).mean()

        data[f"{prefix}_RSI14"] = rsi(close, 14)
        m, s, h = macd(close)
        data[f"{prefix}_MACD"] = m
        data[f"{prefix}_MACD_SIGNAL"] = s
        data[f"{prefix}_MACD_HIST"] = h

        low14 = low.rolling(14).min()
        high14 = high.rolling(14).max()
        den = (high14 - low14).replace(0, np.nan)
        data[f"{prefix}_STOCH_K"] = 100 * (close - low14) / den
        data[f"{prefix}_STOCH_D"] = data[f"{prefix}_STOCH_K"].rolling(3).mean()

        tp = (high + low + close) / 3
        tp_ma = tp.rolling(20).mean()
        tp_md = tp.rolling(20).apply(lambda x: np.mean(np.abs(x - np.mean(x))), raw=True)
        data[f"{prefix}_CCI20"] = (tp - tp_ma) / (0.015 * tp_md.replace(0, np.nan))

        data[f"{prefix}_BB_MID"] = close.rolling(20).mean()
        data[f"{prefix}_BB_STD"] = close.rolling(20).std()
        data[f"{prefix}_BB_UPPER"] = data[f"{prefix}_BB_MID"] + 2 * data[f"{prefix}_BB_STD"]
        data[f"{prefix}_BB_LOWER"] = data[f"{prefix}_BB_MID"] - 2 * data[f"{prefix}_BB_STD"]

        direction = np.sign(close.diff()).fillna(0)
        data[f"{prefix}_OBV"] = (direction * vol).cumsum()
        data[f"{prefix}_OBV_MA20"] = data[f"{prefix}_OBV"].rolling(20).mean()
        data[f"{prefix}_VOL_MA20"] = vol.rolling(20).mean()
        data[f"{prefix}_1Y_HIGH"] = close.rolling(252).max()
        data[f"{prefix}_DD_1Y"] = close / data[f"{prefix}_1Y_HIGH"] - 1

    data["QQQ_MA50_SLOPE_UP"] = data["QQQ_MA50"] > data["QQQ_MA50"].shift(5)
    return data.dropna().copy()


def prepare_data(period: str) -> pd.DataFrame:
    print(f"[data] QQQ/QLD/TQQQ 다운로드 중... period={period}")
    qqq = download_ohlcv("QQQ", period)
    qld = download_ohlcv("QLD", period)
    tqqq = download_ohlcv("TQQQ", period)

    idx = qqq.index.intersection(qld.index).intersection(tqqq.index)
    qqq = qqq.loc[idx]
    qld = qld.loc[idx]
    tqqq = tqqq.loc[idx]

    data = pd.DataFrame(index=idx)
    for col in ["Open", "High", "Low", "Close", "Volume"]:
        data[f"QQQ_{col}"] = qqq[col]
        data[f"QLD_{col}"] = qld[col]
        data[f"TQQQ_{col}"] = tqqq[col]

    data = add_indicators(data)
    print(f"[data] 시작일={data.index[0].date()} 종료일={data.index[-1].date()} rows={len(data):,}")
    return data


def safe_float(value, default: Optional[float] = None) -> Optional[float]:
    try:
        if value is None:
            return default
        v = float(value)
        if math.isnan(v) or math.isinf(v):
            return default
        return v
    except Exception:
        return default


def get_realtime_snapshot(tickers: List[str], close_reference: Dict[str, float]) -> Tuple[Dict[str, Dict[str, object]], str]:
    """Fetch near-real-time quote snapshots for order checking.

    This is intentionally separated from the strategy signal.
    - Strategy signal: daily adjusted Close and daily indicators.
    - Realtime snapshot: last/regular-market price for final order checking.

    Yahoo/yfinance quotes can be delayed depending on market/data conditions,
    so this should be treated as an order-check helper, not an execution feed.
    """
    checked_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    snapshots: Dict[str, Dict[str, object]] = {}

    for ticker in tickers:
        ref_close = safe_float(close_reference.get(ticker), 0.0) or 0.0
        price = None
        source = "unavailable"
        market_state = "unknown"
        currency = "USD"
        previous_close = None

        try:
            info = yf.Ticker(ticker).fast_info
            # Depending on yfinance version, fast_info can behave like a dict or an object.
            def pick(name: str):
                try:
                    if isinstance(info, dict):
                        return info.get(name)
                    return getattr(info, name)
                except Exception:
                    return None

            price = safe_float(pick("last_price"))
            if price is None:
                price = safe_float(pick("lastPrice"))
            previous_close = safe_float(pick("previous_close"))
            if previous_close is None:
                previous_close = safe_float(pick("previousClose"))
            currency = str(pick("currency") or currency)
            if price is not None:
                source = "yfinance fast_info"
        except Exception:
            price = None

        if price is None:
            try:
                hist = yf.Ticker(ticker).history(period="1d", interval="1m", prepost=True, auto_adjust=True)
                if not hist.empty:
                    price = safe_float(hist["Close"].dropna().iloc[-1])
                    source = "yfinance 1m prepost"
            except Exception:
                price = None

        if price is None:
            price = ref_close
            source = "fallback: strategy close"

        diff = price - ref_close if ref_close else 0.0
        diff_pct = diff / ref_close if ref_close else 0.0
        prev_diff = price - previous_close if previous_close else None
        prev_diff_pct = prev_diff / previous_close if previous_close else None

        snapshots[ticker] = {
            "Ticker": ticker,
            "Realtime_Price": price,
            "Strategy_Close": ref_close,
            "Close_Diff": diff,
            "Close_Diff_%": diff_pct * 100,
            "Previous_Close": previous_close,
            "Previous_Close_Diff": prev_diff,
            "Previous_Close_Diff_%": prev_diff_pct * 100 if prev_diff_pct is not None else None,
            "Currency": currency,
            "Market_State": market_state,
            "Source": source,
            "Checked_At_Local": checked_at,
        }

    return snapshots, checked_at


# ======================================================
# Strategy logic
# ======================================================

def calc_risk_score(row: pd.Series, data: pd.DataFrame, asset: str) -> Tuple[int, List[str], pd.DataFrame]:
    rules: List[Dict[str, object]] = []

    def add_rule(name: str, condition: bool, points: int, current, threshold, meaning: str):
        rules.append({
            "자산": asset,
            "구분": "위험점수",
            "항목": name,
            "현재값": current,
            "기준값": threshold,
            "충족여부": yes_no(condition),
            "점수": points if condition else 0,
            "설명": meaning,
        })

    add_rule(f"{asset} 20일선 하회", row[f"{asset}_Close"] < row[f"{asset}_MA20"], 1, row[f"{asset}_Close"], row[f"{asset}_MA20"], "단기 추세 약화")
    add_rule(f"{asset} 50일선 하회", row[f"{asset}_Close"] < row[f"{asset}_MA50"], 2, row[f"{asset}_Close"], row[f"{asset}_MA50"], "중기 추세 약화")
    add_rule("QQQ 50일선 하회", row["QQQ_Close"] < row["QQQ_MA50"], 1, row["QQQ_Close"], row["QQQ_MA50"], "시장 단기/중기 추세 약화")
    add_rule("QQQ 200일선 하회", row["QQQ_Close"] < row["QQQ_MA200"], 3, row["QQQ_Close"], row["QQQ_MA200"], "큰 하락장 방어 신호")
    add_rule(f"{asset} MACD 음전환", row[f"{asset}_MACD_HIST"] < 0, 1, row[f"{asset}_MACD_HIST"], 0, "모멘텀 음수")
    add_rule(f"{asset} RSI 50 하회", row[f"{asset}_RSI14"] < 50, 1, row[f"{asset}_RSI14"], 50, "상승 탄력 부족")
    add_rule("스토캐스틱 약세", row[f"{asset}_STOCH_K"] < row[f"{asset}_STOCH_D"], 1, row[f"{asset}_STOCH_K"], row[f"{asset}_STOCH_D"], "단기 흐름 약세")
    add_rule("CCI 0 하회", row[f"{asset}_CCI20"] < 0, 1, row[f"{asset}_CCI20"], 0, "평균 대비 약한 가격 위치")
    add_rule("볼린저 중단선 하회", row[f"{asset}_Close"] < row[f"{asset}_BB_MID"], 1, row[f"{asset}_Close"], row[f"{asset}_BB_MID"], "20일 평균 아래")
    add_rule("OBV 20일선 하회", row[f"{asset}_OBV"] < row[f"{asset}_OBV_MA20"], 1, row[f"{asset}_OBV"], row[f"{asset}_OBV_MA20"], "거래량 기반 수급 약화")
    add_rule("하락+거래량 증가", row[f"{asset}_Close"] < data[f"{asset}_Close"].iloc[-2] and row[f"{asset}_Volume"] > row[f"{asset}_VOL_MA20"], 1, row[f"{asset}_Volume"], row[f"{asset}_VOL_MA20"], "하락일 거래량 증가")

    df = pd.DataFrame(rules)
    score = int(df["점수"].sum())
    reasons = df.loc[df["점수"] > 0, "항목"].tolist()
    return score, reasons, df


def build_asset_signal(row: pd.Series, data: pd.DataFrame, asset: str, cash: float, shares: float, avg_price: float, target_weight: float, total_equity: float) -> Tuple[Dict[str, object], pd.DataFrame, pd.DataFrame]:
    price = float(row[f"{asset}_Close"])
    value = shares * price
    risk_score, risk_reasons, risk_df = calc_risk_score(row, data, asset)

    sell_rules = [
        [asset, "매도", "QQQ 200일선 이탈", row["QQQ_Close"] < row["QQQ_MA200"], row["QQQ_Close"], "<", row["QQQ_MA200"], "큰 하락장 여부 확인"],
        [asset, "매도", f"{asset} {SELL_MA}일선 이탈", row[f"{asset}_Close"] < row[f"{asset}_MA{SELL_MA}"], row[f"{asset}_Close"], "<", row[f"{asset}_MA{SELL_MA}"], "중기 추세 이탈 확인"],
        [asset, "매도", f"{asset} MACD 음전환", row[f"{asset}_MACD_HIST"] < 0, row[f"{asset}_MACD_HIST"], "<", 0, "모멘텀 약세 확인"],
        [asset, "매도", f"위험점수 {FULL_SELL_SCORE} 이상", risk_score >= FULL_SELL_SCORE, risk_score, ">=", FULL_SELL_SCORE, "위험 신호 누적 확인"],
    ]
    buy_rules = [
        [asset, "매수", f"{asset} 20일선 회복", row[f"{asset}_Close"] > row[f"{asset}_MA20"], row[f"{asset}_Close"], ">", row[f"{asset}_MA20"], "단기 추세 회복 필요"],
        [asset, "매수", f"{asset} MACD 양전환", row[f"{asset}_MACD_HIST"] > 0, row[f"{asset}_MACD_HIST"], ">", 0, "모멘텀 플러스 전환 필요"],
        [asset, "매수", "QQQ 50일선 회복", row["QQQ_Close"] > row["QQQ_MA50"], row["QQQ_Close"], ">", row["QQQ_MA50"], "시장 대표 ETF 회복 필요"],
    ]
    condition_df = pd.DataFrame(sell_rules + buy_rules, columns=["자산", "구분", "조건", "충족", "현재값", "부등호", "기준값", "의미"])
    condition_df["충족여부"] = condition_df["충족"].map(yes_no)
    condition_df["차이"] = condition_df["현재값"].astype(float) - condition_df["기준값"].astype(float)
    condition_df["기준 대비 %"] = np.where(condition_df["기준값"].astype(float) != 0, condition_df["차이"] / condition_df["기준값"].astype(float), np.nan)
    condition_df["해석"] = condition_df.apply(lambda x: "조건 충족" if x["충족"] else f"아직 미충족 - {x['의미']}", axis=1)

    buy_ready = bool(condition_df[condition_df["구분"] == "매수"]["충족"].all())
    sell_ready = bool(condition_df[condition_df["구분"] == "매도"]["충족"].all())
    strong_recovery = bool(buy_ready and row[f"{asset}_RSI14"] >= 50 and row["QQQ_MA50_SLOPE_UP"])

    target_value = total_equity * target_weight
    drift = value - target_value
    drift_pct_of_equity = drift / total_equity if total_equity > 0 else 0.0
    position_return = value / (shares * avg_price) - 1 if shares > 0 and avg_price > 0 else 0.0

    if shares > 0:
        if sell_ready:
            action = f"SELL_ALL_{asset}"
            title = f"{asset} 전량매도"
            order_action = "SELL_ALL"
            recommended_amount = value
            recommended_shares = shares
            reason = f"{asset} 보유 중이며 매도 4조건이 모두 충족되었습니다."
            priority = "방어매도"
        else:
            # 중요: 목표비중을 초과했다는 이유만으로는 매도하지 않는다.
            # 앵커 전략에서 매도는 오직 매도 4조건이 모두 충족될 때만 실행한다.
            action = f"HOLD_{asset}"
            title = f"{asset} 보유유지"
            order_action = "HOLD"
            recommended_amount = 0.0
            recommended_shares = 0.0
            if drift_pct_of_equity > 0.03:
                reason = f"{asset} 보유 중입니다. 목표비중보다 약 {drift_pct_of_equity*100:.2f}%p 높지만, 매도 4조건이 모두 충족되지 않았으므로 팔지 않고 보유합니다."
            else:
                reason = f"{asset} 보유 중이지만 전량매도 조건이 모두 충족되지는 않았습니다."
            priority = "보유"
    else:
        if buy_ready and target_value > 0 and cash > 0:
            buy_amount = min(cash, target_value)
            if not strong_recovery:
                buy_amount *= 0.70
            action = f"BUY_{asset}"
            title = f"{asset} 매수"
            order_action = "BUY"
            recommended_amount = buy_amount
            recommended_shares = buy_amount / price if price > 0 else 0.0
            reason = f"{asset} 미보유 상태이고 매수 3조건이 모두 충족되었습니다."
            priority = "신규매수"
        else:
            action = f"WAIT_{asset}"
            title = f"{asset} 대기"
            order_action = "WAIT"
            recommended_amount = 0.0
            recommended_shares = 0.0
            missing = condition_df[(condition_df["구분"] == "매수") & (~condition_df["충족"])]
            missing_text = ", ".join(missing["조건"].tolist()) if not missing.empty else "현금 또는 목표비중"
            reason = f"{asset} 미보유 상태이며 매수 조건 중 미충족 항목이 있습니다: {missing_text}."
            priority = "대기"

    # If currently holding below target and buy signal is active, optional add-buy
    if shares > 0 and not sell_ready and buy_ready:
        shortage = target_value - value
        if shortage / total_equity > 0.03 if total_equity > 0 else False:
            add_amount = min(cash, shortage)
            if add_amount > 0:
                action = f"ADD_{asset}"
                title = f"{asset} 추가매수"
                order_action = "BUY_MORE"
                recommended_amount = add_amount
                recommended_shares = add_amount / price if price > 0 else 0.0
                reason = f"{asset} 보유 중이고 매수 조건이 충족된 상태에서 목표비중보다 낮습니다. 추가매수 후보입니다."
                priority = "비중확대"

    signal = {
        "Asset": asset,
        "Action": action,
        "Title": title,
        "Order_Action": order_action,
        "Priority": priority,
        "Reason": reason,
        "Price": price,
        "Shares": shares,
        "Avg_Price": avg_price,
        "Position_Value": value,
        "Position_Return_%": position_return * 100,
        "Target_Weight": target_weight,
        "Target_Value": target_value,
        "Current_Weight": value / total_equity if total_equity > 0 else 0.0,
        "Drift_Value": drift,
        "Drift_%p": drift_pct_of_equity * 100,
        "Recommended_Amount": recommended_amount,
        "Recommended_Shares": recommended_shares,
        "Risk_Score": risk_score,
        "Risk_Reasons": ", ".join(risk_reasons) if risk_reasons else "없음",
        "Buy_Ready": buy_ready,
        "Sell_Ready": sell_ready,
        "Strong_Recovery": strong_recovery,
        "MA20": float(row[f"{asset}_MA20"]),
        "MA50": float(row[f"{asset}_MA50"]),
        "MA100": float(row[f"{asset}_MA100"]),
        "RSI14": float(row[f"{asset}_RSI14"]),
        "MACD_HIST": float(row[f"{asset}_MACD_HIST"]),
    }
    return signal, condition_df, risk_df


def allocate_cash(qld_signal: Dict[str, object], tqqq_signal: Dict[str, object], cash: float) -> Tuple[Dict[str, object], Dict[str, object]]:
    """If both assets ask to buy, allocate limited cash by target shortfall proportion."""
    sigs = [qld_signal, tqqq_signal]
    buy_sigs = [s for s in sigs if s["Order_Action"] in ["BUY", "BUY_MORE"] and float(s["Recommended_Amount"]) > 0]
    requested = sum(float(s["Recommended_Amount"]) for s in buy_sigs)
    if requested <= cash or requested <= 0:
        return qld_signal, tqqq_signal
    for s in buy_sigs:
        scaled = cash * float(s["Recommended_Amount"]) / requested
        s["Recommended_Amount"] = scaled
        s["Recommended_Shares"] = scaled / float(s["Price"]) if float(s["Price"]) > 0 else 0.0
        s["Reason"] += " 현금 한도 때문에 매수금액을 비례 조정했습니다."
    return qld_signal, tqqq_signal


# ======================================================
# Dashboard images
# ======================================================

def action_color(order_action: str) -> str:
    if order_action in ["BUY", "BUY_MORE"]:
        return "#0F8A5F"
    if order_action in ["SELL_ALL", "SELL_PARTIAL"]:
        return "#D72638"
    if order_action == "WAIT":
        return "#F59E0B"
    return "#2563EB"


def order_korean(signal: Dict[str, object]) -> str:
    asset = signal["Asset"]
    oa = signal["Order_Action"]
    if oa == "BUY":
        return f"{asset} 사라"
    if oa == "BUY_MORE":
        return f"{asset} 추가매수"
    if oa == "SELL_ALL":
        return f"{asset} 전량 팔아라"
    if oa == "SELL_PARTIAL":
        return f"{asset} 일부 팔아라"
    if oa == "HOLD":
        return f"{asset} 보유"
    return f"{asset} 대기"


def create_dual_dashboard(path: str, today_date: pd.Timestamp, summary: Dict[str, object], qld_signal: Dict[str, object], tqqq_signal: Dict[str, object], condition_df: pd.DataFrame, risk_df: pd.DataFrame) -> str:
    """Create the integrated dashboard.
    v6 layout fix:
    - slightly taller canvas
    - asset decision/reason cards are taller
    - condition boxes moved lower so reason text does not collide
    """
    W, H = 2400, 1850
    bg = "#F3F6FB"
    navy = "#10233F"
    panel = "#FFFFFF"
    line = "#D9E2EF"
    text = "#142033"
    muted = "#64748B"
    green = "#0F8A5F"
    red = "#D72638"
    blue = "#2563EB"
    cream = "#FFF7D6"

    img = Image.new("RGB", (W, H), bg)
    draw = ImageDraw.Draw(img)

    draw_rounded_bar(draw, (50, 40, W - 50, 180), 34, navy)
    draw.text((90, 96), "DUAL QLD + TQQQ ANCHOR DASHBOARD", font=get_font(50, True), fill="white", anchor="lm")
    draw.text((W - 90, 82), today_date.strftime("%Y-%m-%d"), font=get_font(34, True), fill="#C7D7F5", anchor="rm")
    draw.text((W - 90, 126), f"전략 판단 기준: {today_date.strftime('%Y-%m-%d')} 종가", font=get_font(23, True), fill="#FDE68A", anchor="rm")
    draw.text((90, 145), "목표비중 기반 통합 현황판 - 매도는 매도 4조건일 때만, 목표초과는 보유", font=get_font(25), fill="#BFD3F7", anchor="lm")

    # Account summary
    draw_rounded_bar(draw, (50, 220, W - 50, 420), 28, panel, outline=line, width=2)
    acc_items = [
        ("총자산", money(summary["Total_Equity"]), text),
        ("현금", money(summary["Cash"]), text),
        ("QLD 비중", f"{qld_signal['Current_Weight']*100:.1f}% / 목표 {qld_signal['Target_Weight']*100:.1f}%", green),
        ("TQQQ 비중", f"{tqqq_signal['Current_Weight']*100:.1f}% / 목표 {tqqq_signal['Target_Weight']*100:.1f}%", blue),
        ("QQQ 상태", "MA50 위" if summary["QQQ_Above_MA50"] else "MA50 아래", green if summary["QQQ_Above_MA50"] else red),
        ("큰 추세", "MA200 위" if summary["QQQ_Above_MA200"] else "MA200 아래", green if summary["QQQ_Above_MA200"] else red),
    ]
    x = 90
    card_w = 360
    for label, value, color in acc_items:
        draw.text((x, 272), label, font=get_font(24, True), fill=muted, anchor="lm")
        draw.text((x, 340), value, font=get_font(34, True), fill=color, anchor="lm")
        x += card_w

    def asset_card(sig: Dict[str, object], x0: int, y0: int, color_title: str):
        w, h = 1100, 620
        draw_rounded_bar(draw, (x0, y0, x0 + w, y0 + h), 30, panel, outline=line, width=2)
        draw_rounded_bar(draw, (x0, y0, x0 + w, y0 + 90), 30, color_title)
        draw.rectangle((x0, y0 + 48, x0 + w, y0 + 90), fill=color_title)
        draw.text((x0 + 40, y0 + 46), sig["Asset"], font=get_font(42, True), fill="white", anchor="lm")
        draw.text((x0 + w - 40, y0 + 46), str(sig.get("Final_Order_Action_KR", order_korean(sig))), font=get_font(34, True), fill="white", anchor="rm")

        rows = [
            ("종가/현재가", f"{money(sig['Price'])} / {money(sig.get('Live_Price', sig['Price']))}", "보유수량", f"{sig['Shares']:,.6f}"),
            ("종가판단", order_korean(sig), "실시간상태", str(sig.get("Live_Status_KR", "미확인"))),
            ("최종행동", str(sig.get("Final_Order_Action_KR", order_korean(sig))), "현재/목표비중", f"{sig['Current_Weight']*100:.1f}% / {sig['Target_Weight']*100:.1f}%"),
            ("실시간주문", money(sig.get("Live_Order_Amount", 0.0)), "실시간수량", f"{float(sig.get('Live_Order_Shares', 0.0)):,.6f}"),
            ("위험점수", f"{sig['Risk_Score']} / {FULL_SELL_SCORE}", "실시간점수", str(sig.get("Live_Score", 0))),
        ]
        # v14.2 balanced layout fix:
        # Keep the original two-column design, but give the left value area more room.
        # The right-side values such as share counts/status are usually short,
        # so the divider can move closer to the center.
        left_label_x = x0 + 40
        left_value_x = x0 + 500
        right_label_x = x0 + 560
        right_value_x = x0 + w - 40

        yy = y0 + 120
        for a, b, c, d in rows:
            draw.line((x0 + 35, yy + 44, x0 + w - 35, yy + 44), fill=line, width=1)
            draw.text((left_label_x, yy + 22), a, font=get_font(22, True), fill=muted, anchor="lm")
            draw.text((left_value_x, yy + 22), b, font=get_font(24, True), fill=text, anchor="rm")
            draw.text((right_label_x, yy + 22), c, font=get_font(22, True), fill=muted, anchor="lm")
            draw.text((right_value_x, yy + 22), d, font=get_font(24, True), fill=text, anchor="rm")
            yy += 58

        # The reason box starts after the KPI rows and has enough height for 3 wrapped lines.
        reason_box_y = y0 + 440
        draw_rounded_bar(draw, (x0 + 35, reason_box_y, x0 + w - 35, y0 + h - 32), 22, cream, outline="#F3D777", width=1)
        draw.text((x0 + 65, reason_box_y + 32), "판단 이유", font=get_font(25, True), fill=text, anchor="lm")
        for i, ln in enumerate(wrap_lines(str(sig.get("Final_Order_Text", sig.get("Reason", ""))), 52)[:3]):
            draw.text((x0 + 65, reason_box_y + 68 + i * 29), ln, font=get_font(20), fill=text, anchor="la")

    asset_card(qld_signal, 50, 460, green)
    asset_card(tqqq_signal, 1250, 460, blue)

    # Conditions compact table
    def condition_box(asset: str, x0: int, y0: int, color: str):
        sub = condition_df[condition_df["자산"] == asset]
        draw_rounded_bar(draw, (x0, y0, x0 + 1100, y0 + 390), 28, panel, outline=line, width=2)
        draw.text((x0 + 35, y0 + 45), f"{asset} 조건 체크", font=get_font(31, True), fill=text, anchor="lm")
        yy = y0 + 85
        for _, r in sub.iterrows():
            ok = bool(r["충족"])
            c = green if ok else red
            dot_x = x0 + 45
            draw.ellipse((dot_x, yy + 10, dot_x + 22, yy + 32), fill=c)
            draw.text((x0 + 85, yy + 22), f"[{r['구분']}] {r['조건']}", font=get_font(22, True), fill=text, anchor="lm")
            draw.text((x0 + 850, yy + 22), f"{num(r['현재값'],2)} {r['부등호']} {num(r['기준값'],2)}", font=get_font(19), fill=muted, anchor="rm")
            draw.text((x0 + 1060, yy + 22), "충족" if ok else "미충족", font=get_font(22, True), fill=c, anchor="rm")
            yy += 42

    condition_box("QLD", 50, 1135, green)
    condition_box("TQQQ", 1250, 1135, blue)

    # Footer
    draw_rounded_bar(draw, (50, 1615, W - 50, 1765), 30, "#E0F2FE", outline=line, width=2)
    note = "규칙 기반 신호입니다. 전략 전량매도 신호와 비중조절 매도는 성격이 다릅니다. 목표비중 초과만으로는 매도하지 않습니다. 실제 주문 전 현재가, 세금, 수수료, 계좌 상황을 직접 확인하세요."
    for i, ln in enumerate(wrap_lines(note, 95)[:3]):
        draw.text((90, 1660 + i * 34), ln, font=get_font(25, True), fill="#075985", anchor="la")

    img.save(path)
    return path

def create_asset_dashboard(path: str, today_date: pd.Timestamp, summary: Dict[str, object], signal: Dict[str, object], condition_df: pd.DataFrame, risk_df: pd.DataFrame) -> str:
    """Create a clean single-asset dashboard for QLD or TQQQ inside the integrated script."""
    asset = str(signal["Asset"])
    W, H = 2200, 1700
    bg = "#F3F6FB"
    navy = "#10233F"
    panel = "#FFFFFF"
    line = "#D9E2EF"
    text = "#142033"
    muted = "#64748B"
    green = "#0F8A5F"
    red = "#D72638"
    orange = "#F59E0B"
    blue = "#2563EB"
    cream = "#FFF7D6"
    color = action_color(str(signal["Order_Action"]))

    img = Image.new("RGB", (W, H), bg)
    draw = ImageDraw.Draw(img)

    draw_rounded_bar(draw, (50, 40, W - 50, 170), 34, navy)
    draw.text((90, 90), f"{asset} ANCHOR RESEARCH DASHBOARD", font=get_font(48, True), fill="white", anchor="lm")
    draw.text((W - 90, 76), today_date.strftime("%Y-%m-%d"), font=get_font(34, True), fill="#C7D7F5", anchor="rm")
    draw.text((W - 90, 119), f"전략 판단 기준: {today_date.strftime('%Y-%m-%d')} 종가", font=get_font(22, True), fill="#FDE68A", anchor="rm")
    draw.text((90, 138), "통합파일에서 생성된 단독 현황판: 조건, 위험점수, 주문가이드를 한 장으로 요약", font=get_font(23), fill="#BFD3F7", anchor="lm")

    # Action card
    draw_rounded_bar(draw, (50, 210, 520, 440), 30, color)
    draw.text((285, 284), str(signal["Order_Action"]), font=get_font(52, True), fill="white", anchor="mm")
    draw.text((285, 350), str(signal.get("Final_Order_Action_KR", order_korean(signal))), font=get_font(29, True), fill="white", anchor="mm")
    draw.text((285, 402), str(signal["Priority"]), font=get_font(25, True), fill="white", anchor="mm")

    # KPI cards
    kpis = [
        (f"{asset} 종가/현재가", f"{money(signal['Price'])}/{money(signal.get('Live_Price', signal['Price']))}", text),
        ("실시간상태", str(signal.get("Live_Status_KR", "미확인")), color),
        ("위험/실시간", f"{signal['Risk_Score']} / {FULL_SELL_SCORE} · L{signal.get('Live_Score', 0)}", red if int(signal["Risk_Score"]) >= FULL_SELL_SCORE else green),
        ("최종주문금액", money(signal.get("Live_Order_Amount", 0.0)), green if signal.get("Live_Status") in ["LIVE_BUY_OK"] else red if signal.get("Live_Status") in ["LIVE_SELL_ALERT", "EMERGENCY_SELL"] else text),
        ("최종주문수량", f"{float(signal.get('Live_Order_Shares',0.0)):,.6f}", text),
        ("현재비중", pct_value(signal["Current_Weight"]), text),
    ]
    start_x, start_y = 560, 210
    card_w, card_h, gap = 220, 105, 18
    for i, (label, value, c) in enumerate(kpis):
        x = start_x + (i % 3) * (card_w + gap)
        y = start_y + (i // 3) * (card_h + 20)
        draw_rounded_bar(draw, (x, y, x + card_w, y + card_h), 24, panel, outline=line, width=2)
        draw.text((x + 18, y + 32), str(label), font=get_font(20, True), fill=muted, anchor="lm")
        draw.text((x + card_w - 18, y + 72), str(value), font=get_font(24, True), fill=c, anchor="rm")

    # Account/order summary - right side layout fix
    # v11.1: the old box was too short for five rows, so the last row
    # (비중차이) could touch or cross the border. This version gives the
    # summary area more height and uses a stable row rhythm.
    acc_x0, acc_x1 = 1340, W - 50
    acc_y0, acc_y1 = 210, 475
    draw_rounded_bar(draw, (acc_x0, acc_y0, acc_x1, acc_y1), 28, panel, outline=line, width=2)
    draw.text((acc_x0 + 35, acc_y0 + 45), "계좌/주문 요약", font=get_font(31, True), fill=text, anchor="lm")
    rows = [
        ("보유수량", f"{float(signal['Shares']):,.6f}"),
        ("평가금액", money(signal["Position_Value"])),
        ("목표비중", pct_value(signal["Target_Weight"])),
        ("목표금액", money(signal["Target_Value"])),
        ("비중차이", f"{float(signal['Drift_%p']):+.2f}%p"),
    ]
    yy = acc_y0 + 82
    row_h = 36
    for label, value in rows:
        draw.line((acc_x0 + 35, yy + row_h - 7, acc_x1 - 35, yy + row_h - 7), fill=line, width=1)
        draw.text((acc_x0 + 35, yy + 17), label, font=get_font(21, True), fill=muted, anchor="lm")
        draw.text((acc_x1 - 35, yy + 17), value, font=get_font(22, True), fill=text, anchor="rm")
        yy += row_h

    # Reason box - moved slightly lower to preserve spacing after the taller account box
    draw_rounded_bar(draw, (50, 515, W - 50, 675), 28, cream, outline="#F3D777", width=2)
    draw.text((85, 560), "판단 이유", font=get_font(31, True), fill=text, anchor="lm")
    for i, ln in enumerate(wrap_lines(str(signal.get("Final_Order_Text", signal.get("Reason", ""))), 96)[:3]):
        draw.text((85, 605 + i * 32), ln, font=get_font(23), fill=text, anchor="la")

    def table_box(x: int, y: int, w: int, h: int, title: str, rows_df: pd.DataFrame, focus_color: str):
        draw_rounded_bar(draw, (x, y, x + w, y + h), 28, panel, outline=line, width=2)
        draw_rounded_bar(draw, (x, y, x + w, y + 66), 28, focus_color)
        draw.rectangle((x, y + 38, x + w, y + 66), fill=focus_color)
        draw.text((x + 28, y + 34), title, font=get_font(30, True), fill="white", anchor="lm")
        header_y = y + 88
        draw.text((x + 30, header_y), "조건", font=get_font(22, True), fill=muted, anchor="lm")
        draw.text((x + w - 225, header_y), "현재/기준", font=get_font(22, True), fill=muted, anchor="rm")
        draw.text((x + w - 34, header_y), "결과", font=get_font(22, True), fill=muted, anchor="rm")
        yy2 = y + 118
        for _, r in rows_df.iterrows():
            ok = bool(r["충족"])
            c = green if ok else red
            right = f"{num(r['현재값'], 2)} {r['부등호']} {num(r['기준값'], 2)}"
            draw.line((x + 24, yy2 - 12, x + w - 24, yy2 - 12), fill=line, width=1)
            draw.ellipse((x + 28, yy2 + 2, x + 50, yy2 + 24), fill=c)
            draw.text((x + 65, yy2 + 13), str(r["조건"]), font=get_font(23, True), fill=text, anchor="lm")
            draw.text((x + w - 225, yy2 + 13), right, font=get_font(20), fill=muted, anchor="rm")
            draw.text((x + w - 34, yy2 + 13), "충족" if ok else "미충족", font=get_font(23, True), fill=c, anchor="rm")
            yy2 += 62

    asset_cond = condition_df[condition_df["자산"] == asset].copy()
    buy_df = asset_cond[asset_cond["구분"] == "매수"].copy()
    sell_df = asset_cond[asset_cond["구분"] == "매도"].copy()
    table_box(50, 725, 1025, 360, "매수 조건 체크", buy_df, green)
    table_box(1125, 725, 1025, 430, "매도 조건 체크", sell_df, red)

    # Risk box - two columns to avoid clipping
    draw_rounded_bar(draw, (50, 1195, W - 50, 1535), 28, panel, outline=line, width=2)
    draw.text((85, 1241), "위험점수 세부 사유", font=get_font(33, True), fill=text, anchor="lm")
    active = risk_df[(risk_df["자산"] == asset) & (risk_df["점수"] > 0)].copy()
    if active.empty:
        draw.text((85, 1295), "현재 위험점수 부여 사유가 없습니다.", font=get_font(25, True), fill=green, anchor="la")
    else:
        items = list(active.head(10).iterrows())
        for idx, (_, r) in enumerate(items):
            col = idx // 5
            rowi = idx % 5
            x = 90 + col * 1010
            yy3 = 1295 + rowi * 44
            draw.ellipse((x, yy3 + 3, x + 22, yy3 + 25), fill=red)
            draw.text((x + 38, yy3 + 14), f"+{int(r['점수'])}  {r['항목']}", font=get_font(22, True), fill=text, anchor="lm")
            draw.text((x + 390, yy3 + 14), str(r["설명"]), font=get_font(20), fill=muted, anchor="lm")

    footer = "규칙 기반 신호입니다. 목표비중 초과만으로는 매도하지 않으며, 전량매도는 매도 4조건 충족 시에만 표시됩니다. 실제 주문 전 직접 확인하세요."
    for i, ln in enumerate(wrap_lines(footer, 105)[:2]):
        draw.text((W // 2, H - 70 + i * 30), ln, font=get_font(22, True), fill=muted, anchor="mm")

    img.save(path)
    return path



def create_order_summary_dashboard(path: str, today_date: pd.Timestamp, summary: Dict[str, object], qld_signal: Dict[str, object], tqqq_signal: Dict[str, object]) -> str:
    """Create a clean order-summary-only PNG for Telegram and quick checking.
    v6 layout fix:
    - order cards are taller
    - risk-score row no longer overlaps the reason box
    - footer moved lower
    """
    W, H = 1800, 1380
    bg = "#F3F6FB"
    navy = "#10233F"
    panel = "#FFFFFF"
    line = "#D9E2EF"
    text = "#142033"
    muted = "#64748B"
    green = "#0F8A5F"
    red = "#D72638"
    blue = "#2563EB"
    cream = "#FFF7D6"

    img = Image.new("RGB", (W, H), bg)
    draw = ImageDraw.Draw(img)

    draw_rounded_bar(draw, (50, 40, W - 50, 170), 34, navy)
    draw.text((90, 92), "TODAY ORDER SUMMARY", font=get_font(50, True), fill="white", anchor="lm")
    draw.text((W - 90, 78), today_date.strftime("%Y-%m-%d"), font=get_font(34, True), fill="#C7D7F5", anchor="rm")
    draw.text((W - 90, 121), f"전략 판단 기준: {today_date.strftime('%Y-%m-%d')} 종가", font=get_font(22, True), fill="#FDE68A", anchor="rm")
    draw.text((90, 140), "종가 전략 신호 + 실시간 실행 필터 반영 최종 주문 가이드", font=get_font(24), fill="#BFD3F7", anchor="lm")

    # Account overview
    draw_rounded_bar(draw, (50, 205, W - 50, 360), 28, panel, outline=line, width=2)
    overview = [
        ("총자산", money(summary["Total_Equity"]), text),
        ("현재 현금", money(summary["Cash"]), text),
        ("주문 후 예상현금", money(summary["Cash_After_Recommended"]), green if float(summary["Cash_After_Recommended"]) >= 0 else red),
        ("목표비중", f"QLD {summary['QLD_Target_Weight']*100:.0f}% / TQQQ {summary['TQQQ_Target_Weight']*100:.0f}%", blue),
    ]
    x = 90
    for label, value, c in overview:
        draw.text((x, 252), label, font=get_font(23, True), fill=muted, anchor="lm")
        draw.text((x, 315), value, font=get_font(33, True), fill=c, anchor="lm")
        x += 410

    def draw_order_card(sig: Dict[str, object], x0: int, y0: int, accent: str):
        w, h = 810, 650
        oa = str(sig["Order_Action"])
        c = action_color(oa)
        draw_rounded_bar(draw, (x0, y0, x0 + w, y0 + h), 30, panel, outline=line, width=2)
        draw_rounded_bar(draw, (x0, y0, x0 + w, y0 + 95), 30, c)
        draw.rectangle((x0, y0 + 48, x0 + w, y0 + 95), fill=c)
        draw.text((x0 + 35, y0 + 48), str(sig["Asset"]), font=get_font(42, True), fill="white", anchor="lm")
        draw.text((x0 + w - 35, y0 + 48), order_korean(sig), font=get_font(36, True), fill="white", anchor="rm")

        if oa in ["BUY", "BUY_MORE"]:
            order_label = "매수 주문"
            amount_color = green
        elif oa == "SELL_ALL":
            order_label = "전량매도 주문"
            amount_color = red
        else:
            order_label = "주문 없음"
            amount_color = muted

        # Big order numbers
        final_label = str(sig.get("Final_Order_Action_KR", order_korean(sig)))
        live_status = str(sig.get("Live_Status_KR", "실시간 미확인"))
        draw.text((x0 + 45, y0 + 132), "종가 판단", font=get_font(23, True), fill=muted, anchor="lm")
        draw.text((x0 + w - 45, y0 + 132), order_korean(sig), font=get_font(26, True), fill=text, anchor="rm")
        draw.text((x0 + 45, y0 + 176), "실시간 상태", font=get_font(23, True), fill=muted, anchor="lm")
        draw.text((x0 + w - 45, y0 + 176), live_status, font=get_font(27, True), fill=amount_color, anchor="rm")
        draw.text((x0 + 45, y0 + 222), "최종 행동", font=get_font(25, True), fill=muted, anchor="lm")
        draw.text((x0 + w - 45, y0 + 222), final_label, font=get_font(29, True), fill=amount_color, anchor="rm")

        rows = [
            ("종가/현재가", f"{money(sig['Price'])} / {money(sig.get('Live_Price', sig['Price']))}"),
            ("현재가 괴리", f"{float(sig.get('Live_Diff_%', 0.0)):+.2f}%"),
            ("실시간 주문", f"{money(sig.get('Live_Order_Amount', 0.0))} / {float(sig.get('Live_Order_Shares', 0.0)):,.4f}주"),
            ("보유수량", f"{float(sig['Shares']):,.6f} 주"),
            ("위험/실시간점수", f"{sig['Risk_Score']} / {FULL_SELL_SCORE} · L{sig.get('Live_Score', 0)}"),
        ]
        yy = y0 + 265
        for label, value in rows:
            draw.line((x0 + 40, yy + 37, x0 + w - 40, yy + 37), fill=line, width=1)
            draw.text((x0 + 45, yy + 18), label, font=get_font(22, True), fill=muted, anchor="lm")
            draw.text((x0 + w - 45, yy + 18), value, font=get_font(24, True), fill=text, anchor="rm")
            yy += 45

        # Start reason box below all rows so the risk-score row remains fully visible.
        reason_y0 = y0 + 505
        reason_y1 = y0 + h - 32
        draw_rounded_bar(draw, (x0 + 40, reason_y0, x0 + w - 40, reason_y1), 20, cream, outline="#F3D777", width=1)
        reason_text = str(sig.get("Final_Order_Text", sig.get("Reason", "")))
        for i, ln in enumerate(wrap_lines(reason_text, 46)[:3]):
            draw.text((x0 + 65, reason_y0 + 28 + i * 28), ln, font=get_font(20), fill=text, anchor="la")

    draw_order_card(qld_signal, 50, 405, green)
    draw_order_card(tqqq_signal, 940, 405, blue)

    # Bottom rule reminder
    draw_rounded_bar(draw, (50, 1110, W - 50, 1280), 30, "#E0F2FE", outline=line, width=2)
    notes = [
        "종가 판단은 백테스트 기준 앵커, 실시간 상태는 주문 직전 실행 필터/비상 브레이크입니다.",
        "LIVE_SELL_ALERT/EMERGENCY_SELL은 장중 급락과 핵심선 훼손을 반영한 방어 경고입니다.",
    ]
    for i, ln in enumerate(notes):
        draw.text((90, 1160 + i * 42), ln, font=get_font(25, True), fill="#075985", anchor="la")

    draw.text((W // 2, H - 35), "규칙 기반 신호입니다. 실제 주문 전 현재가, 세금, 수수료, 계좌 상황을 직접 확인하세요.", font=get_font(22, True), fill=muted, anchor="mm")
    img.save(path)
    return path


def build_realtime_order_df(signals: List[Dict[str, object]], snapshots: Dict[str, Dict[str, object]]) -> pd.DataFrame:
    """Create a separate order-check table using quote snapshots.

    Strategy logic remains daily-close based. This table only recalculates the
    final order-check price, position value, and buy shares using the latest
    available quote snapshot.
    """
    rows: List[Dict[str, object]] = []
    for sig in signals:
        asset = str(sig["Asset"])
        snap = snapshots.get(asset, {})
        strategy_price = safe_float(sig.get("Price"), 0.0) or 0.0
        realtime_price = safe_float(snap.get("Realtime_Price"), strategy_price) or strategy_price
        shares = safe_float(sig.get("Shares"), 0.0) or 0.0
        rec_amount = safe_float(sig.get("Recommended_Amount"), 0.0) or 0.0
        close_based_rec_shares = safe_float(sig.get("Recommended_Shares"), 0.0) or 0.0
        order_action = str(sig.get("Order_Action", ""))
        if order_action in ["BUY", "BUY_MORE"] and realtime_price > 0:
            realtime_rec_shares = rec_amount / realtime_price
        elif order_action in ["SELL_ALL", "SELL_PARTIAL"]:
            realtime_rec_shares = shares
        else:
            realtime_rec_shares = 0.0
        rows.append({
            "Asset": asset,
            "Order_Action": order_action,
            "Korean_Order": order_korean(sig),
            "Strategy_Close_Basis_Date": snap.get("Strategy_Basis_Date", ""),
            "Strategy_Close": strategy_price,
            "Realtime_Check_Price": realtime_price,
            "Close_Diff": realtime_price - strategy_price,
            "Close_Diff_%": ((realtime_price / strategy_price - 1) * 100) if strategy_price else 0.0,
            "Shares": shares,
            "Position_Value_Strategy_Close": shares * strategy_price,
            "Position_Value_Realtime": shares * realtime_price,
            "Recommended_Amount": rec_amount,
            "Recommended_Shares_Close_Basis": close_based_rec_shares,
            "Recommended_Shares_Realtime_Check": realtime_rec_shares,
            "Quote_Source": snap.get("Source", ""),
            "Checked_At_Local": snap.get("Checked_At_Local", ""),
            "Note": "전략 신호는 종가 기준, 실제 주문 전 가격/수량 확인은 현재가 기준 참고",
        })
    return pd.DataFrame(rows)


# ======================================================
# v14 Live execution layer
# ======================================================

def live_status_korean(status: str) -> str:
    mapping = {
        "LIVE_BUY_OK": "실시간 매수 가능",
        "PULLBACK_WAIT": "추격매수 주의/눌림 대기",
        "LIVE_BUY_BLOCK": "장중 조건 훼손/매수 보류",
        "LIVE_HOLD_OK": "실시간 보유 유지",
        "LIVE_SELL_ALERT": "장중 방어 경고",
        "EMERGENCY_SELL": "비상 방어매도 후보",
        "LIVE_WAIT": "실시간 대기",
    }
    return mapping.get(str(status), str(status))


def final_order_action_korean(sig: Dict[str, object]) -> str:
    status = str(sig.get("Live_Status", ""))
    asset = str(sig.get("Asset", ""))
    if status == "LIVE_BUY_OK":
        return f"{asset} 현재가 기준 매수 가능"
    if status == "PULLBACK_WAIT":
        return f"{asset} 추격매수 금지, 눌림 대기"
    if status == "LIVE_BUY_BLOCK":
        return f"{asset} 장중 조건 훼손, 매수 보류"
    if status == "LIVE_SELL_ALERT":
        return f"{asset} 일부 축소 검토"
    if status == "EMERGENCY_SELL":
        return f"{asset} 전량/대폭 방어매도 검토"
    if status == "LIVE_HOLD_OK":
        return f"{asset} 보유 유지"
    return order_korean(sig)


def build_live_execution_signal(sig: Dict[str, object], row: pd.Series, snapshots: Dict[str, Dict[str, object]]) -> Dict[str, object]:
    """Add a live execution layer on top of the daily-close anchor signal.

    Daily anchor signal is kept intact in Order_Action/Reason.
    Live_Status and Final_Order_Text are the practical execution layer shown to the user.
    This layer is intentionally a filter/brake, not a replacement for backtested daily logic.
    """
    asset = str(sig.get("Asset", ""))
    out = dict(sig)
    out["Daily_Order_Action"] = str(sig.get("Order_Action", ""))
    out["Daily_Korean_Order"] = order_korean(sig)

    strategy_price = safe_float(sig.get("Price"), 0.0) or 0.0
    live_price = safe_float(snapshots.get(asset, {}).get("Realtime_Price"), strategy_price) or strategy_price
    qqq_strategy = safe_float(row.get("QQQ_Close"), 0.0) or 0.0
    qqq_live = safe_float(snapshots.get("QQQ", {}).get("Realtime_Price"), qqq_strategy) or qqq_strategy
    shares = safe_float(sig.get("Shares"), 0.0) or 0.0
    daily_action = str(sig.get("Order_Action", ""))
    risk_score = int(safe_float(sig.get("Risk_Score"), 0.0) or 0.0)
    ma20 = safe_float(sig.get("MA20"), 0.0) or 0.0
    ma100 = safe_float(sig.get("MA100"), 0.0) or 0.0
    macd_hist = safe_float(sig.get("MACD_HIST"), 0.0) or 0.0
    qqq_ma50 = safe_float(row.get("QQQ_MA50"), 0.0) or 0.0
    qqq_ma200 = safe_float(row.get("QQQ_MA200"), 0.0) or 0.0

    live_diff_pct = ((live_price / strategy_price - 1) * 100) if strategy_price else 0.0
    qqq_live_diff_pct = ((qqq_live / qqq_strategy - 1) * 100) if qqq_strategy else 0.0
    chase_limit = 1.5 if asset == "QLD" else 2.5
    asset_drop_alert = -7.0 if asset == "QLD" else -14.0
    asset_drop_emergency = -10.0 if asset == "QLD" else -20.0

    reasons: List[str] = []
    live_score = 0

    def add(cond: bool, points: int, text: str):
        nonlocal live_score
        if bool(cond):
            live_score += int(points)
            reasons.append(text)

    add(live_price < ma20, 1, f"{asset} 현재가 MA20 이탈")
    add(live_price < ma100, 2, f"{asset} 현재가 MA100 이탈")
    add(qqq_live < qqq_ma50, 1, "QQQ 현재가 MA50 이탈")
    add(qqq_live < qqq_ma200, 3, "QQQ 현재가 MA200 이탈")
    add(live_diff_pct <= asset_drop_alert, 1, f"{asset} 장중 하락률 {live_diff_pct:+.2f}%")
    add(qqq_live_diff_pct <= -5.0, 1, f"QQQ 장중 하락률 {qqq_live_diff_pct:+.2f}%")
    add(macd_hist < 0, 1, f"{asset} 일봉 MACD Histogram 음수")
    add(risk_score >= 6, 1, f"전일 기준 위험점수 {risk_score}점")

    status = "LIVE_WAIT"
    final_text = "종가 기준 신호를 확인하고 실시간 가격은 참고합니다."
    live_sell_ratio = 0.0

    if shares > 0:
        emergency = (
            (live_price < ma100 and (qqq_live < qqq_ma200 or live_diff_pct <= asset_drop_emergency or risk_score >= 6))
            or (live_score >= 6 and live_price < ma20)
            or daily_action == "SELL_ALL"
        )
        alert = live_score >= 3 or (live_price < ma20 and qqq_live < qqq_ma50)
        if emergency:
            status = "EMERGENCY_SELL"
            live_sell_ratio = 1.0 if daily_action == "SELL_ALL" or qqq_live < qqq_ma200 or live_diff_pct <= asset_drop_emergency else 0.7
            final_text = f"종가 기준은 {order_korean(sig)}입니다. 하지만 실시간으로 핵심 방어선이 훼손되어 전량 또는 대폭 방어매도를 검토합니다."
        elif alert:
            status = "LIVE_SELL_ALERT"
            live_sell_ratio = 0.5 if live_score >= 4 else 0.3
            final_text = f"종가 기준은 {order_korean(sig)}입니다. 다만 장중 약세 신호가 누적되어 {int(live_sell_ratio*100)}% 축소를 검토합니다."
        else:
            status = "LIVE_HOLD_OK"
            final_text = f"종가 기준 {order_korean(sig)}이며, 실시간 가격도 핵심 방어선을 크게 훼손하지 않아 보유 유지입니다."
    else:
        if daily_action in ["BUY", "BUY_MORE"]:
            if live_price < ma20 or qqq_live < qqq_ma50:
                status = "LIVE_BUY_BLOCK"
                final_text = f"종가 기준은 {order_korean(sig)} 후보지만, 장중 현재가가 핵심 매수 조건을 훼손하여 매수 보류입니다."
            elif live_diff_pct > chase_limit:
                status = "PULLBACK_WAIT"
                reasons.append(f"{asset} 현재가가 종가 대비 {live_diff_pct:+.2f}%로 추격매수 제한 {chase_limit:.1f}% 초과")
                final_text = f"종가 기준은 {order_korean(sig)} 후보지만, 현재가가 종가 대비 높아 추격매수하지 않고 눌림 대기입니다."
            else:
                status = "LIVE_BUY_OK"
                final_text = f"종가 기준 {order_korean(sig)}이며, 실시간 가격도 허용 범위라 현재가 기준 매수 가능합니다."
        else:
            status = "LIVE_WAIT"
            final_text = f"종가 기준 {order_korean(sig)}이며, 실시간 기준으로도 신규 주문은 대기입니다."

    rec_amount = safe_float(sig.get("Recommended_Amount"), 0.0) or 0.0
    if status == "LIVE_BUY_OK" and live_price > 0:
        live_order_amount = rec_amount
        live_order_shares = rec_amount / live_price
    elif status in ["LIVE_SELL_ALERT", "EMERGENCY_SELL"]:
        live_order_shares = shares * live_sell_ratio
        live_order_amount = live_order_shares * live_price
    else:
        live_order_amount = 0.0
        live_order_shares = 0.0

    out.update({
        "Live_Status": status,
        "Live_Status_KR": live_status_korean(status),
        "Final_Order_Action_KR": final_order_action_korean({**out, "Live_Status": status}),
        "Final_Order_Text": final_text,
        "Live_Reason": ", ".join(reasons) if reasons else "실시간 훼손 신호 없음",
        "Live_Score": live_score,
        "Live_Price": live_price,
        "Live_Diff_%": live_diff_pct,
        "QQQ_Live_Price": qqq_live,
        "QQQ_Live_Diff_%": qqq_live_diff_pct,
        "Live_Order_Amount": live_order_amount,
        "Live_Order_Shares": live_order_shares,
        "Live_Sell_Ratio": live_sell_ratio,
        "Live_Checked_At": snapshots.get(asset, {}).get("Checked_At_Local", ""),
        "Live_Quote_Source": snapshots.get(asset, {}).get("Source", ""),
    })
    return out


def apply_live_execution_layer(signals: List[Dict[str, object]], row: pd.Series, snapshots: Dict[str, Dict[str, object]]) -> List[Dict[str, object]]:
    return [build_live_execution_signal(sig, row, snapshots) for sig in signals]


def enrich_realtime_df_with_live(realtime_df: pd.DataFrame, signals: List[Dict[str, object]]) -> pd.DataFrame:
    if realtime_df is None or realtime_df.empty:
        return realtime_df
    live_map = {str(s.get("Asset", "")): s for s in signals}
    rows: List[Dict[str, object]] = []
    for _, r in realtime_df.iterrows():
        row = r.to_dict()
        sig = live_map.get(str(row.get("Asset", "")), {})
        row.update({
            "Daily_Order_Action": sig.get("Daily_Order_Action", row.get("Order_Action", "")),
            "Live_Status": sig.get("Live_Status", ""),
            "Live_Status_KR": sig.get("Live_Status_KR", ""),
            "Final_Order_Action_KR": sig.get("Final_Order_Action_KR", ""),
            "Final_Order_Text": sig.get("Final_Order_Text", ""),
            "Live_Reason": sig.get("Live_Reason", ""),
            "Live_Score": sig.get("Live_Score", 0),
            "Live_Order_Amount": sig.get("Live_Order_Amount", 0.0),
            "Live_Order_Shares": sig.get("Live_Order_Shares", 0.0),
            "Live_Sell_Ratio": sig.get("Live_Sell_Ratio", 0.0),
        })
        rows.append(row)
    return pd.DataFrame(rows)


def create_realtime_order_dashboard(path: str, today_date: pd.Timestamp, checked_at: str, summary: Dict[str, object], realtime_df: pd.DataFrame, qld_signal: Dict[str, object], tqqq_signal: Dict[str, object]) -> str:
    """Create a separate PNG for final order checking with quote snapshots."""
    W, H = 2200, 1450
    bg = "#F3F6FB"
    navy = "#10233F"
    panel = "#FFFFFF"
    line = "#D9E2EF"
    text = "#142033"
    muted = "#64748B"
    green = "#0F8A5F"
    red = "#D72638"
    blue = "#2563EB"
    amber = "#F59E0B"
    cream = "#FFF7D6"

    img = Image.new("RGB", (W, H), bg)
    draw = ImageDraw.Draw(img)

    draw_rounded_bar(draw, (50, 40, W - 50, 180), 34, navy)
    draw.text((90, 88), "REALTIME ORDER CHECK", font=get_font(52, True), fill="white", anchor="lm")
    draw.text((90, 138), "종가 앵커 신호 + 실시간 실행 필터/비상 브레이크로 최종 행동 표시", font=get_font(25), fill="#BFD3F7", anchor="lm")
    draw.text((W - 90, 78), f"전략: {today_date.strftime('%Y-%m-%d')} 종가", font=get_font(30, True), fill="#FDE68A", anchor="rm")
    draw.text((W - 90, 122), f"현재가 확인: {checked_at}", font=get_font(24, True), fill="#C7D7F5", anchor="rm")

    # Summary card
    draw_rounded_bar(draw, (50, 220, W - 50, 380), 28, panel, outline=line, width=2)
    qld_rt = realtime_df.loc[realtime_df["Asset"] == "QLD"].iloc[0].to_dict() if not realtime_df[realtime_df["Asset"] == "QLD"].empty else {}
    tqqq_rt = realtime_df.loc[realtime_df["Asset"] == "TQQQ"].iloc[0].to_dict() if not realtime_df[realtime_df["Asset"] == "TQQQ"].empty else {}
    rt_total = float(summary.get("Cash", 0.0)) + float(qld_rt.get("Position_Value_Realtime", 0.0)) + float(tqqq_rt.get("Position_Value_Realtime", 0.0))
    close_total = float(summary.get("Total_Equity", 0.0))
    items = [
        ("종가 기준 총자산", money(close_total), text),
        ("현재가 기준 총자산", money(rt_total), blue),
        ("총자산 차이", f"{money(rt_total - close_total)} ({(rt_total / close_total - 1) * 100 if close_total else 0:+.2f}%)", green if rt_total >= close_total else red),
        ("현재 현금", money(summary.get("Cash", 0.0)), text),
    ]
    x = 90
    for label, value, c in items:
        draw.text((x, 268), label, font=get_font(24, True), fill=muted, anchor="lm")
        draw.text((x, 332), value, font=get_font(34, True), fill=c, anchor="lm")
        x += 510

    def row_card(row: Dict[str, object], sig: Dict[str, object], x0: int, y0: int, accent: str):
        w, h = 1025, 620
        action = str(row.get("Order_Action", ""))
        act_color = action_color(action)
        close_p = float(row.get("Strategy_Close", 0.0) or 0.0)
        rt_p = float(row.get("Realtime_Check_Price", 0.0) or 0.0)
        diff_pct = float(row.get("Close_Diff_%", 0.0) or 0.0)
        draw_rounded_bar(draw, (x0, y0, x0 + w, y0 + h), 30, panel, outline=line, width=2)
        draw_rounded_bar(draw, (x0, y0, x0 + w, y0 + 90), 30, accent)
        draw.rectangle((x0, y0 + 48, x0 + w, y0 + 90), fill=accent)
        draw.text((x0 + 35, y0 + 45), str(row.get("Asset", "")), font=get_font(42, True), fill="white", anchor="lm")
        draw.text((x0 + w - 35, y0 + 45), str(row.get("Final_Order_Action_KR", row.get("Korean_Order", ""))), font=get_font(30, True), fill="white", anchor="rm")

        draw.text((x0 + 45, y0 + 112), "종가 판단", font=get_font(22, True), fill=muted, anchor="lm")
        draw.text((x0 + 430, y0 + 112), str(row.get("Korean_Order", "")), font=get_font(25, True), fill=text, anchor="rm")
        draw.text((x0 + 450, y0 + 112), "실시간 상태", font=get_font(22, True), fill=muted, anchor="lm")
        draw.text((x0 + w - 45, y0 + 112), str(row.get("Live_Status_KR", "")), font=get_font(25, True), fill=act_color, anchor="rm")

        draw.text((x0 + 45, y0 + 160), "전략 판단가", font=get_font(24, True), fill=muted, anchor="lm")
        draw.text((x0 + 430, y0 + 160), money(close_p), font=get_font(34, True), fill=text, anchor="rm")
        draw.text((x0 + 45, y0 + 220), "주문 전 현재가", font=get_font(24, True), fill=muted, anchor="lm")
        draw.text((x0 + 430, y0 + 220), money(rt_p), font=get_font(42, True), fill=blue, anchor="rm")
        draw.text((x0 + 45, y0 + 280), "종가 대비", font=get_font(24, True), fill=muted, anchor="lm")
        draw.text((x0 + 450, y0 + 280), f"{money(rt_p - close_p)} ({diff_pct:+.2f}%)", font=get_font(28, True), fill=green if diff_pct >= 0 else red, anchor="rm")

        right_rows = [
            ("종가기준 추천", money(row.get("Recommended_Amount", 0.0))),
            ("실시간 최종금액", money(row.get("Live_Order_Amount", 0.0))),
            ("실시간 최종수량", f"{float(row.get('Live_Order_Shares', 0.0) or 0.0):,.6f} 주"),
            ("현재가 평가금액", money(row.get("Position_Value_Realtime", 0.0))),
            ("데이터 출처", str(row.get("Quote_Source", ""))),
        ]
        yy = y0 + 126
        for label, value in right_rows:
            draw.line((x0 + 430, yy + 38, x0 + w - 40, yy + 38), fill=line, width=1)
            draw.text((x0 + 450, yy + 18), label, font=get_font(22, True), fill=muted, anchor="lm")
            draw.text((x0 + w - 45, yy + 18), value, font=get_font(24, True), fill=act_color if label == "현재가기준 주문수량" else text, anchor="rm")
            yy += 54

        draw_rounded_bar(draw, (x0 + 40, y0 + 445, x0 + w - 40, y0 + h - 32), 22, cream, outline="#F3D777", width=1)
        guide = str(row.get("Final_Order_Text", "이 화면은 주문 전 확인용입니다."))
        for i, ln in enumerate(wrap_lines(guide, 54)[:3]):
            draw.text((x0 + 65, y0 + 485 + i * 30), ln, font=get_font(22, True), fill=text, anchor="la")

    row_card(qld_rt, qld_signal, 50, 430, green)
    row_card(tqqq_rt, tqqq_signal, 1125, 430, blue)

    draw_rounded_bar(draw, (50, 1100, W - 50, 1320), 30, "#E0F2FE", outline=line, width=2)
    notes = [
        "① 종가 판단: 백테스트 기준 앵커입니다. 이동평균/MACD/위험점수는 확정 종가로 계산됩니다.",
        "② 실시간 상태: 현재가로 추격매수, 조건 훼손, 장중 방어매도 경고를 판단합니다.",
        "③ 현재가 스냅샷은 지연될 수 있습니다. 최종 주문 가격은 반드시 증권사 앱에서 다시 확인하세요.",
    ]
    for i, ln in enumerate(notes):
        draw.text((90, 1150 + i * 45), ln, font=get_font(25, True), fill="#075985", anchor="la")
    draw.text((W // 2, H - 45), "투자 조언이 아니라 사용자가 정한 규칙 기반 확인 도구입니다.", font=get_font(23, True), fill=muted, anchor="mm")
    img.save(path)
    return path


def append_trade_log_if_requested(args, today_date: pd.Timestamp, out_dir: str) -> pd.DataFrame:
    """Append actual executions to a trade log when --record-trade is used."""
    trade_log_path = args.trade_log if os.path.isabs(args.trade_log) else os.path.join(out_dir, args.trade_log)
    rows: List[Dict[str, object]] = []
    if getattr(args, "record_trade", False):
        trade_date = args.trade_date or today_date.strftime("%Y-%m-%d")
        specs = [
            ("QLD", args.qld_exec_action, args.qld_exec_shares, args.qld_exec_price, args.qld_avg_price),
            ("TQQQ", args.tqqq_exec_action, args.tqqq_exec_shares, args.tqqq_exec_price, args.tqqq_avg_price),
        ]
        for asset, action, shares, price, avg_price in specs:
            action = str(action).upper()
            shares = safe_float(shares, 0.0) or 0.0
            price = safe_float(price, 0.0) or 0.0
            avg_price = safe_float(avg_price, 0.0) or 0.0
            if action in ["BUY", "SELL"] and shares > 0 and price > 0:
                gross = shares * price
                realized_pl = shares * (price - avg_price) if action == "SELL" and avg_price > 0 else 0.0
                rows.append({
                    "Recorded_At": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Trade_Date": trade_date,
                    "Asset": asset,
                    "Action": action,
                    "Shares": shares,
                    "Price": price,
                    "Gross_Amount": gross,
                    "Avg_Price_Used_For_Sell_PL": avg_price,
                    "Approx_Realized_PL": realized_pl,
                    "Approx_Realized_Return_%": ((price / avg_price - 1) * 100) if action == "SELL" and avg_price > 0 else 0.0,
                    "Memo": "SELL 수익은 입력한 평균단가 기준의 단순 추정값입니다.",
                })
    new_df = pd.DataFrame(rows)
    if not new_df.empty:
        if os.path.exists(trade_log_path):
            old = pd.read_csv(trade_log_path)
            all_df = pd.concat([old, new_df], ignore_index=True)
        else:
            all_df = new_df
        all_df.to_csv(trade_log_path, index=False, encoding="utf-8-sig")
        return all_df
    if os.path.exists(trade_log_path):
        return pd.read_csv(trade_log_path)
    return pd.DataFrame(columns=["Recorded_At", "Trade_Date", "Asset", "Action", "Shares", "Price", "Gross_Amount", "Avg_Price_Used_For_Sell_PL", "Approx_Realized_PL", "Approx_Realized_Return_%", "Memo"])


def update_performance_history(args, out_dir: str, today_date: pd.Timestamp, summary: Dict[str, object], realtime_df: pd.DataFrame, qld_signal: Dict[str, object], tqqq_signal: Dict[str, object]) -> pd.DataFrame:
    """Append the current account snapshot and compute cumulative return."""
    perf_path = args.performance_log if os.path.isabs(args.performance_log) else os.path.join(out_dir, args.performance_log)
    def get_row(asset: str) -> Dict[str, object]:
        sub = realtime_df[realtime_df["Asset"] == asset]
        return sub.iloc[0].to_dict() if not sub.empty else {}
    qld = get_row("QLD")
    tqqq = get_row("TQQQ")
    cash = float(summary.get("Cash", 0.0))
    total_close = float(summary.get("Total_Equity", 0.0))
    total_rt = cash + float(qld.get("Position_Value_Realtime", 0.0) or 0.0) + float(tqqq.get("Position_Value_Realtime", 0.0) or 0.0)
    row = {
        "Run_At": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Strategy_Date": today_date.strftime("%Y-%m-%d"),
        "Cash": cash,
        "QLD_Shares": float(qld_signal.get("Shares", 0.0)),
        "QLD_Avg_Price": float(qld_signal.get("Avg_Price", 0.0)),
        "QLD_Strategy_Close": float(qld.get("Strategy_Close", qld_signal.get("Price", 0.0)) or 0.0),
        "QLD_Realtime_Price": float(qld.get("Realtime_Check_Price", qld_signal.get("Price", 0.0)) or 0.0),
        "QLD_Value_Realtime": float(qld.get("Position_Value_Realtime", 0.0) or 0.0),
        "QLD_Unrealized_PL": float(qld_signal.get("Shares", 0.0)) * (float(qld.get("Realtime_Check_Price", qld_signal.get("Price", 0.0)) or 0.0) - float(qld_signal.get("Avg_Price", 0.0))),
        "TQQQ_Shares": float(tqqq_signal.get("Shares", 0.0)),
        "TQQQ_Avg_Price": float(tqqq_signal.get("Avg_Price", 0.0)),
        "TQQQ_Strategy_Close": float(tqqq.get("Strategy_Close", tqqq_signal.get("Price", 0.0)) or 0.0),
        "TQQQ_Realtime_Price": float(tqqq.get("Realtime_Check_Price", tqqq_signal.get("Price", 0.0)) or 0.0),
        "TQQQ_Value_Realtime": float(tqqq.get("Position_Value_Realtime", 0.0) or 0.0),
        "TQQQ_Unrealized_PL": float(tqqq_signal.get("Shares", 0.0)) * (float(tqqq.get("Realtime_Check_Price", tqqq_signal.get("Price", 0.0)) or 0.0) - float(tqqq_signal.get("Avg_Price", 0.0))),
        "Total_Equity_Strategy_Close": total_close,
        "Total_Equity_Realtime": total_rt,
        "QLD_Order": order_korean(qld_signal),
        "TQQQ_Order": order_korean(tqqq_signal),
        "QLD_Live_Status": qld_signal.get("Live_Status_KR", ""),
        "TQQQ_Live_Status": tqqq_signal.get("Live_Status_KR", ""),
        "QLD_Final_Order": qld_signal.get("Final_Order_Action_KR", ""),
        "TQQQ_Final_Order": tqqq_signal.get("Final_Order_Action_KR", ""),
    }
    new_df = pd.DataFrame([row])
    if os.path.exists(perf_path):
        old = pd.read_csv(perf_path)
        hist = pd.concat([old, new_df], ignore_index=True)
    else:
        hist = new_df
    base = float(hist["Total_Equity_Realtime"].iloc[0]) if len(hist) and float(hist["Total_Equity_Realtime"].iloc[0]) != 0 else 0.0
    hist["Cumulative_Return_%"] = (hist["Total_Equity_Realtime"] / base - 1) * 100 if base else 0.0
    hist["Snapshot_Change"] = hist["Total_Equity_Realtime"].diff().fillna(0.0)
    hist["Snapshot_Change_%"] = hist["Total_Equity_Realtime"].pct_change().fillna(0.0) * 100
    hist.to_csv(perf_path, index=False, encoding="utf-8-sig")
    return hist


def create_performance_dashboard(path: str, today_date: pd.Timestamp, performance_df: pd.DataFrame, trade_df: pd.DataFrame) -> str:
    """Create a cumulative performance PNG from saved snapshots and executions."""
    W, H = 2200, 1500
    bg = "#F3F6FB"
    navy = "#10233F"
    panel = "#FFFFFF"
    line = "#D9E2EF"
    text = "#142033"
    muted = "#64748B"
    green = "#0F8A5F"
    red = "#D72638"
    blue = "#2563EB"
    img = Image.new("RGB", (W, H), bg)
    draw = ImageDraw.Draw(img)

    latest = performance_df.iloc[-1].to_dict() if not performance_df.empty else {}
    first_equity = float(performance_df["Total_Equity_Realtime"].iloc[0]) if not performance_df.empty else 0.0
    latest_equity = float(latest.get("Total_Equity_Realtime", 0.0) or 0.0)
    cum_ret = float(latest.get("Cumulative_Return_%", 0.0) or 0.0)
    change = float(latest.get("Snapshot_Change", 0.0) or 0.0)
    change_pct = float(latest.get("Snapshot_Change_%", 0.0) or 0.0)
    qld_upl = float(latest.get("QLD_Unrealized_PL", 0.0) or 0.0)
    tqqq_upl = float(latest.get("TQQQ_Unrealized_PL", 0.0) or 0.0)
    realized = float(trade_df["Approx_Realized_PL"].sum()) if not trade_df.empty and "Approx_Realized_PL" in trade_df.columns else 0.0

    draw_rounded_bar(draw, (50, 40, W - 50, 180), 34, navy)
    draw.text((90, 88), "CUMULATIVE PERFORMANCE TRACKER", font=get_font(50, True), fill="white", anchor="lm")
    draw.text((90, 138), "실제 매매 후 보유수량/현금 입력값을 누적 저장하여 성과를 추적", font=get_font(25), fill="#BFD3F7", anchor="lm")
    draw.text((W - 90, 82), today_date.strftime("%Y-%m-%d"), font=get_font(34, True), fill="#C7D7F5", anchor="rm")
    draw.text((W - 90, 126), "현재가 기준 누적 성과", font=get_font(23, True), fill="#FDE68A", anchor="rm")

    cards = [
        ("첫 기록 총자산", money(first_equity), text),
        ("현재 총자산", money(latest_equity), blue),
        ("누적 수익률", f"{cum_ret:+.2f}%", green if cum_ret >= 0 else red),
        ("최근 변화", f"{money(change)} ({change_pct:+.2f}%)", green if change >= 0 else red),
    ]
    x = 70
    for label, value, c in cards:
        draw_rounded_bar(draw, (x, 230, x + 500, 390), 28, panel, outline=line, width=2)
        draw.text((x + 30, 282), label, font=get_font(25, True), fill=muted, anchor="lm")
        draw.text((x + 470, 340), value, font=get_font(38, True), fill=c, anchor="rm")
        x += 535

    # Asset P/L cards
    asset_rows = [
        ("QLD", latest.get("QLD_Shares", 0.0), latest.get("QLD_Avg_Price", 0.0), latest.get("QLD_Realtime_Price", 0.0), qld_upl),
        ("TQQQ", latest.get("TQQQ_Shares", 0.0), latest.get("TQQQ_Avg_Price", 0.0), latest.get("TQQQ_Realtime_Price", 0.0), tqqq_upl),
    ]
    for i, (asset, shares, avgp, price, upl) in enumerate(asset_rows):
        x0 = 70 + i * 1065
        y0 = 445
        draw_rounded_bar(draw, (x0, y0, x0 + 1000, y0 + 310), 30, panel, outline=line, width=2)
        draw_rounded_bar(draw, (x0, y0, x0 + 1000, y0 + 78), 30, blue if asset == "TQQQ" else green)
        draw.rectangle((x0, y0 + 38, x0 + 1000, y0 + 78), fill=blue if asset == "TQQQ" else green)
        draw.text((x0 + 35, y0 + 40), f"{asset} 보유 성과", font=get_font(34, True), fill="white", anchor="lm")
        rows = [
            ("보유수량", f"{float(shares):,.6f} 주"),
            ("평균단가", money(avgp)),
            ("현재확인가", money(price)),
            ("평가손익", money(upl)),
        ]
        yy = y0 + 115
        for label, value in rows:
            draw.line((x0 + 35, yy + 34, x0 + 965, yy + 34), fill=line, width=1)
            draw.text((x0 + 45, yy + 17), label, font=get_font(23, True), fill=muted, anchor="lm")
            c = green if label == "평가손익" and float(upl) >= 0 else red if label == "평가손익" else text
            draw.text((x0 + 955, yy + 17), value, font=get_font(26, True), fill=c, anchor="rm")
            yy += 46

    # Recent snapshot table
    draw_rounded_bar(draw, (70, 805, W - 70, 1215), 30, panel, outline=line, width=2)
    draw.text((105, 855), "최근 누적 기록", font=get_font(34, True), fill=text, anchor="lm")
    cols = [("기록시간", 120), ("총자산", 620), ("누적수익률", 950), ("QLD 평가손익", 1260), ("TQQQ 평가손익", 1600), ("최근변화", 1970)]
    for label, xcol in cols:
        draw.text((xcol, 920), label, font=get_font(22, True), fill=muted, anchor="lm" if label == "기록시간" else "rm")
    recent = performance_df.tail(6).copy()
    yy = 965
    for _, r in recent.iterrows():
        draw.line((105, yy - 23, W - 105, yy - 23), fill=line, width=1)
        vals = [
            str(r.get("Run_At", ""))[:16],
            money(r.get("Total_Equity_Realtime", 0.0)),
            f"{float(r.get('Cumulative_Return_%', 0.0)):+.2f}%",
            money(r.get("QLD_Unrealized_PL", 0.0)),
            money(r.get("TQQQ_Unrealized_PL", 0.0)),
            money(r.get("Snapshot_Change", 0.0)),
        ]
        for idx, val in enumerate(vals):
            label, xcol = cols[idx]
            fill = text
            if idx in [2, 3, 4, 5]:
                try:
                    numv = float(str(val).replace("$", "").replace(",", "").replace("%", ""))
                    fill = green if numv >= 0 else red
                except Exception:
                    fill = text
            draw.text((xcol, yy), val, font=get_font(22, True), fill=fill, anchor="lm" if idx == 0 else "rm")
        yy += 48

    # Trade summary
    draw_rounded_bar(draw, (70, 1260, W - 70, 1415), 30, "#FFF7D6", outline="#F3D777", width=2)
    trade_count = len(trade_df) if trade_df is not None else 0
    draw.text((105, 1310), "실제 체결 기록 요약", font=get_font(30, True), fill=text, anchor="lm")
    draw.text((105, 1362), f"누적 체결 {trade_count}건 / 매도 실현손익 추정 {money(realized)}", font=get_font(27, True), fill=green if realized >= 0 else red, anchor="lm")
    draw.text((W - 105, 1362), "SELL 손익은 입력 평균단가 기준 단순 추정입니다.", font=get_font(23, True), fill=muted, anchor="rm")
    img.save(path)
    return path


# ======================================================
# Game UI Dashboards - v9
# ======================================================

def _blend_hex(c1: str, c2: str, t: float) -> str:
    t = max(0.0, min(1.0, float(t)))
    a = tuple(int(c1.lstrip('#')[i:i+2], 16) for i in (0, 2, 4))
    b = tuple(int(c2.lstrip('#')[i:i+2], 16) for i in (0, 2, 4))
    m = tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))
    return '#%02x%02x%02x' % m


def _draw_gradient_rect(draw: ImageDraw.ImageDraw, xy, top: str, bottom: str):
    x0, y0, x1, y1 = map(int, xy)
    h = max(1, y1 - y0)
    for y in range(y0, y1):
        draw.line((x0, y, x1, y), fill=_blend_hex(top, bottom, (y - y0) / h))


def _draw_glow_text(draw: ImageDraw.ImageDraw, xy, text_value: str, font, fill: str, anchor: str = 'mm', glow: str = '#00ffbb', radius: int = 2):
    x, y = xy
    for dx in range(-radius, radius + 1):
        for dy in range(-radius, radius + 1):
            if dx or dy:
                draw.text((x + dx, y + dy), text_value, font=font, fill=glow, anchor=anchor)
    draw.text((x, y), text_value, font=font, fill=fill, anchor=anchor)



def _safe_status(sig: Dict[str, object]) -> Tuple[str, str, str]:
    oa = str(sig.get('Order_Action', ''))
    risk = int(sig.get('Risk_Score', 0) or 0)
    if oa in ['BUY', 'BUY_MORE']:
        return 'READY', '진입 가능 · 매수 게이트 열림', '#2EE6A6'
    if oa == 'SELL_ALL':
        return 'BLOCKED', '방어 매도 · 4조건 동시 충족', '#FF5A66'
    if oa == 'HOLD':
        if risk >= FULL_SELL_SCORE - 2:
            return 'GUARD', '보유 중 · 리스크 감시', '#FFD166'
        return 'MINING', '보유 중 · 채굴 가동', '#2EE6A6'
    return 'IDLE', '대기 · 매수 조건 미충족', '#7EA2D6'


def _poly(draw, pts, fill, outline=None, width=1):
    draw.polygon(pts, fill=fill)
    if outline:
        draw.line(pts + [pts[0]], fill=outline, width=width, joint='curve')


def _iso_platform(draw, cx, cy, w, h, depth, top, left, right, outline):
    top_pts = [(cx, cy-h//2), (cx+w//2, cy), (cx, cy+h//2), (cx-w//2, cy)]
    left_pts = [(cx-w//2, cy), (cx, cy+h//2), (cx, cy+h//2+depth), (cx-w//2, cy+depth)]
    right_pts = [(cx+w//2, cy), (cx, cy+h//2), (cx, cy+h//2+depth), (cx+w//2, cy+depth)]
    _poly(draw, left_pts, left, outline, 2)
    _poly(draw, right_pts, right, outline, 2)
    _poly(draw, top_pts, top, outline, 3)
    return top_pts


def _draw_ore_pile(draw, x, y, color, outline, count=7):
    for i in range(count):
        ox = x + (i % 4) * 32 - (i//4)*15
        oy = y - (i//4)*24 + (i % 2) * 5
        pts = [(ox, oy+18), (ox+14, oy), (ox+32, oy+5), (ox+42, oy+22), (ox+22, oy+30)]
        _poly(draw, pts, color, outline, 1)
        draw.line((ox+9, oy+17, ox+28, oy+8), fill=_blend_hex(color, '#FFFFFF', 0.35), width=2)


def _draw_mine_house(draw, cx, cy, accent, base_color, label):
    # pseudo 3D mine building: roof, body, dark tunnel, rails, ores
    roof_top = [(cx-155, cy-115), (cx-40, cy-185), (cx+155, cy-115), (cx+35, cy-50)]
    roof_side = [(cx-155, cy-115), (cx+35, cy-50), (cx+35, cy-18), (cx-155, cy-80)]
    _poly(draw, roof_side, _blend_hex(base_color, '#000000', 0.18), accent, 2)
    _poly(draw, roof_top, _blend_hex(base_color, '#FFFFFF', 0.18), accent, 3)
    body = [(cx-120, cy-80), (cx+120, cy-80), (cx+120, cy+95), (cx-120, cy+95)]
    draw.rounded_rectangle((cx-120, cy-80, cx+120, cy+95), radius=24, fill=base_color, outline=accent, width=3)
    draw.rounded_rectangle((cx-52, cy-35, cx+52, cy+92), radius=32, fill='#050A12', outline=accent, width=3)
    draw.ellipse((cx-42, cy-25, cx+42, cy+65), fill='#081323')
    draw.rectangle((cx-42, cy+20, cx+42, cy+92), fill='#081323')
    # beams
    for bx in [cx-105, cx+105]:
        draw.line((bx, cy-72, bx, cy+88), fill=_blend_hex(accent, '#FFD166', 0.35), width=7)
    draw.line((cx-112, cy-58, cx+112, cy-58), fill=_blend_hex(accent, '#FFD166', 0.35), width=7)
    # rails
    for off in [-18, 18]:
        draw.line((cx+off, cy+88, cx+off*4, cy+205), fill='#B88640', width=5)
    for j in range(5):
        yy = cy+110+j*20
        draw.line((cx-75+j*11, yy, cx+75-j*11, yy), fill='#D9B16A', width=3)
    draw.rounded_rectangle((cx-150, cy+118, cx+150, cy+165), radius=20, fill='#07111E', outline=accent, width=2)
    draw.text((cx, cy+141), label, font=get_font(25, True), fill=accent, anchor='mm')


def _draw_worker(draw, x, y, hardhat='#FFD166', suit='#2EE6A6'):
    draw.ellipse((x-14, y-35, x+14, y-7), fill='#F6C08B', outline='#1B2430', width=2)
    draw.pieslice((x-18, y-43, x+18, y-12), 180, 360, fill=hardhat, outline='#2B1E06', width=2)
    draw.rounded_rectangle((x-20, y-7, x+20, y+40), radius=10, fill=suit, outline='#09111D', width=2)
    draw.line((x-30, y+4, x-58, y+28), fill='#D9B16A', width=5)
    draw.line((x-58, y+28, x-70, y+12), fill='#D9B16A', width=4)
    draw.line((x+30, y+2, x+52, y+24), fill='#D9B16A', width=5)
    draw.line((x-12, y+40, x-22, y+70), fill='#D9B16A', width=6)
    draw.line((x+12, y+40, x+22, y+70), fill='#D9B16A', width=6)


def create_game_mine_dashboard(path: str, today_date: pd.Timestamp, summary: Dict[str, object], qld_signal: Dict[str, object], tqqq_signal: Dict[str, object], condition_df: pd.DataFrame, risk_df: pd.DataFrame) -> str:
    W, H = 2600, 1650
    img = Image.new('RGB', (W, H), '#050912')
    draw = ImageDraw.Draw(img)
    _draw_gradient_rect(draw, (0, 0, W, H), '#050912', '#0B1320')
    for i in range(180):
        x = (i * 173 + 41) % W
        y = (i * 97 + 83) % H
        r = 1 + (i % 4)
        color = '#18375E' if i % 5 else '#916C38'
        draw.ellipse((x, y, x+r, y+r), fill=color)
    qld_status, qld_desc, qld_status_color = _safe_status(qld_signal)
    tqqq_status, tqqq_desc, tqqq_status_color = _safe_status(tqqq_signal)
    qld_pl = float(qld_signal.get('Position_Value', 0.0) or 0.0) - float(qld_signal.get('Shares', 0.0) or 0.0) * float(qld_signal.get('Avg_Price', 0.0) or 0.0)
    tqqq_pl = float(tqqq_signal.get('Position_Value', 0.0) or 0.0) - float(tqqq_signal.get('Shares', 0.0) or 0.0) * float(tqqq_signal.get('Avg_Price', 0.0) or 0.0)
    core_output = qld_pl + tqqq_pl
    total_risk = int(qld_signal.get('Risk_Score', 0) or 0) + int(tqqq_signal.get('Risk_Score', 0) or 0)
    mode = 'RESTING' if total_risk < 8 else 'GUARD MODE'
    mode_color = '#2EE6A6' if total_risk < 8 else '#FF5A66'

    draw.rounded_rectangle((50, 38, 920, 125), radius=26, fill='#0D1828', outline='#1A365D', width=2)
    draw.ellipse((78, 58, 130, 110), fill='#FFD166', outline='#FFF0AA', width=3)
    draw.text((104, 84), 'Q/T', font=get_font(22, True), fill='#1A1A1A', anchor='mm')
    draw.text((160, 83), 'DUAL ANCHOR MINE v11', font=get_font(40, True), fill='#EAF3FF', anchor='lm')
    draw.text((640, 83), today_date.strftime('%Y-%m-%d'), font=get_font(26, True), fill='#8FB2E8', anchor='lm')
    draw.rounded_rectangle((2120, 38, 2525, 125), radius=24, fill='#140D12' if total_risk >= 8 else '#0C1B18', outline=mode_color, width=2)
    draw.text((2160, 76), '●', font=get_font(30, True), fill=mode_color, anchor='lm')
    draw.text((2210, 72), mode, font=get_font(34, True), fill=mode_color, anchor='lm')
    draw.text((2210, 106), '리스크 오프 / 신규 진입 관리', font=get_font(18, True), fill='#B8C7DD', anchor='lm')

    # platforms and connector rails
    _iso_platform(draw, 530, 610, 720, 420, 90, '#173D34', '#0F2B25', '#123127', '#23E0A4')
    _iso_platform(draw, 2070, 610, 720, 420, 90, '#4A2B12', '#2B180C', '#3A210E', '#FF8A1E')
    _iso_platform(draw, 1300, 670, 640, 520, 120, '#19162B', '#100E1C', '#131020', '#FFD166')
    for y in [615, 665, 715]:
        draw.line((890, y, 990, y+30), fill='#B88640', width=8)
        draw.line((1610, y+30, 1710, y), fill='#B88640', width=8)
        draw.line((990, y+30, 1085, y+35), fill='#D2A353', width=4)
        draw.line((1515, y+35, 1610, y+30), fill='#D2A353', width=4)

    _draw_mine_house(draw, 530, 520, '#2EE6A6', '#225B4B', 'QLD SAFE MINE')
    _draw_mine_house(draw, 2070, 520, '#FF9B2F', '#6A3916', 'TQQQ POWER MINE')
    _draw_ore_pile(draw, 380, 765, '#2EE6A6', '#0B5C45', 8)
    _draw_ore_pile(draw, 2075, 765, '#FF9B2F', '#8A430E', 8)
    _draw_worker(draw, 350, 715, hardhat='#FFD166', suit='#2EE6A6')
    _draw_worker(draw, 445, 800, hardhat='#FFD166', suit='#38BDF8')
    _draw_worker(draw, 2175, 785, hardhat='#FFD166', suit='#FF8A1E')
    _draw_worker(draw, 2260, 715, hardhat='#FFD166', suit='#FF5A66')

    # core reactor
    cx, cy = 1300, 610
    draw.ellipse((cx-205, cy-205, cx+205, cy+205), fill='#211A22', outline='#FFE19A', width=6)
    for rr, cc, ww in [(330, '#392620', 5), (250, '#6A4428', 5), (165, '#B47A36', 4), (95, '#FFE08A', 5)]:
        draw.ellipse((cx-rr//2, cy-rr//2, cx+rr//2, cy+rr//2), outline=cc, width=ww)
    for i in range(12):
        ang = i * 30
        # simple radial ticks
        import math as _m
        x1 = cx + int(_m.cos(_m.radians(ang))*118); y1 = cy + int(_m.sin(_m.radians(ang))*118)
        x2 = cx + int(_m.cos(_m.radians(ang))*158); y2 = cy + int(_m.sin(_m.radians(ang))*158)
        draw.line((x1,y1,x2,y2), fill='#77502D', width=3)
    _draw_glow_text(draw, (cx, cy-86), 'MINE CORE · LIVE', get_font(34, True), '#FFE7A8', 'mm', '#7A4B10', 1)
    _draw_glow_text(draw, (cx, cy+8), money(core_output), get_font(76, True), '#2EE6A6' if core_output >= 0 else '#FF5A66', 'mm', '#0C442F' if core_output >= 0 else '#42171B', 2)
    draw.text((cx, cy+95), f"TOTAL EQUITY {money(summary.get('Total_Equity', 0.0))}", font=get_font(28, True), fill='#EAF3FF', anchor='mm')
    draw.text((cx, cy+140), f"CASH {money(summary.get('Cash', 0.0))}", font=get_font(24, True), fill='#92A8C7', anchor='mm')

    def mine_panel(sig, x0, y0, accent, desc, status_color):
        w, h = 930, 275
        draw.rounded_rectangle((x0+14, y0+14, x0+w+14, y0+h+14), radius=30, fill='#03070D')
        draw.rounded_rectangle((x0, y0, x0+w, y0+h), radius=30, fill='#0C1523', outline=accent, width=2)
        draw.text((x0+35, y0+50), str(sig['Asset']), font=get_font(38, True), fill=accent, anchor='lm')
        draw.text((x0+w-35, y0+50), order_korean(sig), font=get_font(34, True), fill='#EAF3FF', anchor='rm')
        draw.rounded_rectangle((x0+35, y0+82, x0+w-35, y0+132), radius=18, fill='#081220', outline=status_color, width=2)
        draw.text((x0+60, y0+107), desc, font=get_font(23, True), fill=status_color, anchor='lm')
        rows = [('평가금액', money(sig.get('Position_Value', 0.0))), ('위험점수', f"{sig.get('Risk_Score', 0)} / {FULL_SELL_SCORE}"), ('추천주문', f"{money(sig.get('Recommended_Amount', 0.0))} · {float(sig.get('Recommended_Shares', 0.0)):,.4f}주")]
        yy = y0 + 170
        for label, value in rows:
            draw.text((x0+45, yy), label, font=get_font(22, True), fill='#8FA5C3', anchor='lm')
            draw.text((x0+w-45, yy), value, font=get_font(24, True), fill='#EAF3FF', anchor='rm')
            yy += 36
    mine_panel(qld_signal, 100, 1035, '#2EE6A6', qld_desc, qld_status_color)
    mine_panel(tqqq_signal, 1570, 1035, '#FF9B2F', tqqq_desc, tqqq_status_color)

    draw.rounded_rectangle((50, 1390, W-50, 1580), radius=28, fill='#09111D', outline='#20314B', width=2)
    controls = [('CORE OUTPUT', money(core_output), '#2EE6A6' if core_output >= 0 else '#FF5A66'), ('QLD DRIFT', f"{float(qld_signal.get('Drift_%p', 0.0)):+.2f}%p", '#2EE6A6'), ('TQQQ DRIFT', f"{float(tqqq_signal.get('Drift_%p', 0.0)):+.2f}%p", '#FF9B2F'), ('TOTAL RISK', f"{total_risk} / {FULL_SELL_SCORE*2}", mode_color), ('SOLVER', 'STANDBY' if total_risk < 8 else 'GUARD', '#B794F4')]
    x = 85
    for label, value, color in controls:
        draw.rounded_rectangle((x, 1430, x+470, 1535), radius=18, fill='#0C1827', outline='#265167', width=2)
        draw.text((x+24, 1462), label, font=get_font(20, True), fill='#8FA5C3', anchor='lm')
        draw.text((x+24, 1508), value, font=get_font(34, True), fill=color, anchor='lm')
        x += 500
    draw.rounded_rectangle((210, 1592, 2390, 1635), radius=20, fill='#101B2B', outline='#1E3351', width=2)
    msg = f"GOLDIE AI: QLD는 {qld_status}, TQQQ는 {tqqq_status}. 매수는 3조건, 매도는 4조건 동시 충족을 기준으로 판단합니다."
    draw.text((245, 1615), msg, font=get_font(22, True), fill='#D9E7FF', anchor='lm')
    img.save(path)
    return path


def create_account_vibe_dashboard(path: str, today_date: pd.Timestamp, data: pd.DataFrame, signal: Dict[str, object], summary: Dict[str, object], theme: str = 'SAFE') -> str:
    asset = str(signal['Asset'])
    safe = theme.upper() == 'SAFE'
    accent = '#2EE6A6' if safe else '#FF8A1E'
    accent2 = '#1BBF87' if safe else '#B45D17'
    bg = '#070C14'; panel = '#0D1522'; line = '#1B2A3F'; muted = '#9BAEC9'; white = '#EAF3FF'; gold = '#FFD166'; red = '#FF5A66'
    W, H = 1160, 2600
    img = Image.new('RGB', (W, H), bg)
    draw = ImageDraw.Draw(img)
    _draw_gradient_rect(draw, (0, 0, W, H), '#070C14', '#0A111C')
    title = 'SAFE 안정형 계좌' if safe else 'POWER 공격형 계좌'
    base = float(signal.get('Shares', 0.0) or 0.0) * float(signal.get('Avg_Price', 0.0) or 0.0)
    value = float(signal.get('Position_Value', 0.0) or 0.0)
    if base > 0:
        pnl = value - base
        ret = pnl / base * 100
    else:
        # 미보유 상태는 -100%처럼 오해되지 않게 0 기준으로 표시
        pnl = 0.0
        ret = 0.0
    draw.rounded_rectangle((45, 35, W-45, 125), radius=28, fill='#0D1828', outline=accent, width=2)
    draw.text((W//2, 80), f"{'POWER' if not safe else 'SAFE'} {asset} · {title}", font=get_font(38, True), fill=accent, anchor='mm')
    draw.text((W//2, 165), '누적 수익률 · 평균단가 대비', font=get_font(27, True), fill=accent, anchor='mm')
    _draw_glow_text(draw, (W//2, 310), f"{ret:+.1f}%", get_font(120, True), accent if ret >= 0 else red, 'mm', '#0E3B2E' if safe else '#4A240C', 2)
    draw.text((W//2, 445), '순이익 · 평가 기준', font=get_font(26, True), fill=muted, anchor='mm')
    draw.text((W//2, 520), money(pnl), font=get_font(70, True), fill=accent if pnl >= 0 else red, anchor='mm')
    y = 630
    kpis = [('수익률', f"{ret:+.1f}%"), ('일평균', f"{ret/21:+.2f}%"), ('위험점수', f"{signal.get('Risk_Score', 0)}/{FULL_SELL_SCORE}")]
    for i, (label, val) in enumerate(kpis):
        x0 = 70 + i * 350
        draw.rounded_rectangle((x0, y, x0+310, y+135), radius=26, fill=panel, outline=line, width=2)
        draw.text((x0+155, y+50), val, font=get_font(38, True), fill=white if i == 2 else accent if ret >= 0 else red, anchor='mm')
        draw.text((x0+155, y+98), label, font=get_font(22, True), fill=muted, anchor='mm')
    draw.text((W//2, 830), f"Verified · {asset} Anchor Engine", font=get_font(23, True), fill=accent, anchor='mm')
    draw.text((W//2, 865), f"{asset} {money(value)} · live · rule based", font=get_font(22, True), fill=white, anchor='mm')
    draw.line((75, 905, W-75, 905), fill=line, width=2)

    close_col = f'{asset}_Close'
    recent = (data[close_col].pct_change().dropna().tail(5) * 100).tolist() if close_col in data.columns else [0, 0, 0, 0, 0]
    draw.text((75, 960), '일별 성과 · 최근 5거래일', font=get_font(28, True), fill=muted, anchor='lm')
    bar_base = 1220
    for i, val in enumerate(recent):
        x0 = 100 + i * 200
        bh = int(35 + min(120, abs(val) * 12))
        c = accent if val >= 0 else red
        draw.rounded_rectangle((x0, bar_base-bh, x0+135, bar_base), radius=16, fill=c, outline=_blend_hex(c, '#FFFFFF', 0.25), width=1)
        draw.text((x0+67, bar_base-bh-25), f"{val:+.1f}%", font=get_font(22, True), fill=c, anchor='mm')
        draw.text((x0+67, bar_base+35), f"D{i+1}", font=get_font(22, True), fill=muted, anchor='mm')
    y2 = 1295
    draw.rounded_rectangle((0, y2, W, y2+190), radius=0, fill='#0A111D', outline=line, width=2)
    draw.line((W//2, y2, W//2, y2+190), fill=line, width=2)
    buy_ready = 'ON' if bool(signal.get('Buy_Ready')) else 'OFF'
    sell_ready = 'ON' if bool(signal.get('Sell_Ready')) else 'OFF'
    draw.text((W//4, y2+70), buy_ready, font=get_font(64, True), fill=accent if buy_ready == 'ON' else muted, anchor='mm')
    draw.text((W//4, y2+130), '매수 게이트', font=get_font(24, True), fill=muted, anchor='mm')
    draw.text((W*3//4, y2+70), sell_ready, font=get_font(64, True), fill=red if sell_ready == 'ON' else gold, anchor='mm')
    draw.text((W*3//4, y2+130), '매도 게이트', font=get_font(24, True), fill=muted, anchor='mm')

    # Calendar: actual month grid, not only 21 cells. Existing trading days only are filled.
    import calendar as _calendar
    cal_y = 1570
    draw.text((70, cal_y-62), today_date.strftime('%Y년 %m월'), font=get_font(36, True), fill=white, anchor='lm')
    weekdays = ['월', '화', '수', '목', '금', '토', '일']
    cell_w, cell_h = 140, 95
    gap_x, gap_y = 13, 15
    for cidx, wd in enumerate(weekdays):
        draw.text((70 + cidx*(cell_w+gap_x) + cell_w//2, cal_y-15), wd, font=get_font(20, True), fill=muted, anchor='mm')
    month_start = pd.Timestamp(year=today_date.year, month=today_date.month, day=1)
    _, last_day = _calendar.monthrange(today_date.year, today_date.month)
    month_days = [pd.Timestamp(year=today_date.year, month=today_date.month, day=d) for d in range(1, last_day+1)]
    returns = {}
    if close_col in data.columns:
        ret_series = data[close_col].pct_change() * 100
        for idx, val in ret_series.items():
            ts = pd.Timestamp(idx).normalize()
            if ts.year == today_date.year and ts.month == today_date.month and not pd.isna(val):
                returns[ts.day] = float(val)
    for dts in month_days:
        row = (dts.day + month_start.weekday() - 1) // 7
        col = dts.weekday()
        x0 = 70 + col * (cell_w + gap_x)
        y0 = cal_y + row * (cell_h + gap_y)
        val = returns.get(dts.day, None)
        if val is None:
            fill = '#0A111D'; outline = '#152235'; txt = '#50627C'; sub = ''
        elif val >= 0:
            fill = _blend_hex('#0D1B21', accent2, min(0.68, 0.18 + abs(val)/7)); outline = accent; txt = accent; sub = f"{val:+.2f}%"
        else:
            fill = _blend_hex('#171018', '#7A1D25', min(0.68, 0.18 + abs(val)/7)); outline = red; txt = '#FF8A8A'; sub = f"{val:+.2f}%"
        draw.rounded_rectangle((x0, y0, x0+cell_w, y0+cell_h), radius=16, fill=fill, outline=outline, width=1)
        draw.text((x0+14, y0+22), str(dts.day), font=get_font(19, True), fill=muted if val is not None else txt, anchor='lm')
        if sub:
            draw.text((x0+15, y0+64), sub, font=get_font(22, True), fill=txt, anchor='lm')
    reason = str(signal.get('Reason', ''))
    draw.rounded_rectangle((45, H-145, W-45, H-45), radius=24, fill='#101B2B', outline=line, width=2)
    for i, ln in enumerate(wrap_lines(reason, 58)[:2] if reason else ['조건 확인 완료']):
        draw.text((75, H-105+i*34), ln, font=get_font(22, True), fill=white, anchor='lm')
    img.save(path)
    return path


def create_risk_gate_dashboard(path: str, today_date: pd.Timestamp, condition_df: pd.DataFrame, risk_df: pd.DataFrame, qld_signal: Dict[str, object], tqqq_signal: Dict[str, object]) -> str:
    W, H = 1800, 2600
    img = Image.new('RGB', (W, H), '#070C14')
    draw = ImageDraw.Draw(img)
    _draw_gradient_rect(draw, (0, 0, W, H), '#070C14', '#0A111C')
    white = '#EAF3FF'; muted = '#9BAEC9'; green = '#2EE6A6'; orange = '#FF8A1E'; red = '#FF5A66'; gold = '#FFD166'; panel = '#0D1522'; line = '#1B2A3F'
    active_count = int(condition_df['충족'].sum()) if '충족' in condition_df.columns else 0
    total_count = len(condition_df)
    draw.text((55, 80), 'BUY / SELL RISK GATES', font=get_font(52, True), fill=white, anchor='lm')
    draw.rounded_rectangle((1240, 35, 1715, 120), radius=24, fill='#101823', outline=gold, width=2)
    draw.text((1478, 77), f'{active_count}/{total_count} 조건 충족', font=get_font(32, True), fill=gold, anchor='mm')
    draw.text((55, 145), '위험점수 세부사유는 활성 항목만 점수에 더해집니다. 아래 표는 전체 위험규칙을 모두 보여줍니다.', font=get_font(25, True), fill=muted, anchor='lm')
    def gate_section(asset: str, sig: Dict[str, object], y0: int, accent: str):
        draw.rounded_rectangle((45, y0, W-45, y0+1120), radius=34, fill=panel, outline=line, width=2)
        draw.text((90, y0+55), f'{asset} CONTROL GATES', font=get_font(38, True), fill=accent, anchor='lm')
        draw.text((W-90, y0+55), order_korean(sig), font=get_font(34, True), fill=white, anchor='rm')
        sub = condition_df[condition_df['자산'] == asset].copy()
        buy = sub[sub['구분'] == '매수'].copy(); sell = sub[sub['구분'] == '매도'].copy()
        risk_all = risk_df[risk_df['자산'] == asset].copy()
        risk_score = int(sig.get('Risk_Score', 0) or 0)
        # buy/sell top boxes
        draw.rounded_rectangle((90, y0+110, W-90, y0+305), radius=28, fill='#0B1F1B', outline=green, width=2)
        draw.text((125, y0+155), 'BUY GATE · 진입 허가 3조건', font=get_font(30, True), fill=green, anchor='lm')
        yy = y0 + 205
        for _, r in buy.iterrows():
            ok = bool(r['충족']); c = green if ok else red
            draw.ellipse((125, yy-14, 155, yy+16), fill=c)
            draw.text((175, yy), str(r['조건']), font=get_font(24, True), fill=white, anchor='lm')
            draw.text((W-140, yy), f"{num(r['현재값'],2)} {r['부등호']} {num(r['기준값'],2)} · {'충족' if ok else '미충족'}", font=get_font(22, True), fill=c, anchor='rm')
            yy += 45
        draw.rounded_rectangle((90, y0+335, W-90, y0+575), radius=28, fill='#241016', outline=red, width=2)
        draw.text((125, y0+380), 'SELL GUARD · 방어 매도 4조건', font=get_font(30, True), fill=red, anchor='lm')
        yy = y0 + 430
        for _, r in sell.iterrows():
            ok = bool(r['충족']); c = red if ok else green
            draw.ellipse((125, yy-14, 155, yy+16), fill=c)
            draw.text((175, yy), str(r['조건']), font=get_font(24, True), fill=white, anchor='lm')
            draw.text((W-140, yy), f"{num(r['현재값'],2)} {r['부등호']} {num(r['기준값'],2)} · {'발동' if ok else '미발동'}", font=get_font(22, True), fill=c, anchor='rm')
            yy += 45
        # risk meter
        mx0, my0, mx1, my1 = 120, y0+645, W-120, y0+705
        draw.text((120, y0+620), f'위험점수 {risk_score}/{FULL_SELL_SCORE}', font=get_font(28, True), fill=white, anchor='lm')
        draw.rounded_rectangle((mx0, my0, mx1, my1), radius=28, fill='#1A2230', outline=line, width=2)
        fill_w = int((mx1 - mx0) * min(1.0, risk_score / FULL_SELL_SCORE))
        meter_color = green if risk_score < 4 else gold if risk_score < FULL_SELL_SCORE else red
        if fill_w > 0:
            draw.rounded_rectangle((mx0, my0, mx0+fill_w, my1), radius=28, fill=meter_color)
        draw.text((mx1, y0+620), 'NORMAL' if risk_score < 4 else 'WATCH' if risk_score < FULL_SELL_SCORE else 'BLOCK', font=get_font(28, True), fill=meter_color, anchor='rm')
        # all risk rules in two columns
        draw.text((90, y0+770), '위험점수 전체 규칙 · 빨간색만 현재 점수에 반영', font=get_font(28, True), fill=white, anchor='lm')
        risk_items = risk_all.to_dict('records')
        for idx, r in enumerate(risk_items[:12]):
            col = idx // 6
            row = idx % 6
            x = 105 + col * 825
            yy = y0 + 825 + row * 44
            pts = int(r.get('점수', 0) or 0)
            active = pts > 0
            c = red if active else green
            draw.ellipse((x, yy-13, x+26, yy+13), fill=c)
            draw.text((x+42, yy), f"{ '+'+str(pts) if active else '+0'}  {r.get('항목','')}", font=get_font(21, True), fill=white if active else muted, anchor='lm')
            draw.text((x+455, yy), str(r.get('설명',''))[:25], font=get_font(19), fill=muted, anchor='lm')
    gate_section('QLD', qld_signal, 220, green)
    gate_section('TQQQ', tqqq_signal, 1370, orange)
    img.save(path)
    return path


# ======================================================
# v11 overrides: cute 3D mine + weekday account calendar
# ======================================================

def _draw_toy_platform(draw, cx, cy, w, h, d, top, left, right, outline):
    top_poly = [(cx, cy-h//2), (cx+w//2, cy), (cx, cy+h//2), (cx-w//2, cy)]
    left_poly = [(cx-w//2, cy), (cx, cy+h//2), (cx, cy+h//2+d), (cx-w//2, cy+d)]
    right_poly = [(cx+w//2, cy), (cx, cy+h//2), (cx, cy+h//2+d), (cx+w//2, cy+d)]
    draw.polygon([(x+26, y+d+34) for x, y in top_poly], fill="#02050B")
    draw.polygon(left_poly, fill=left, outline=outline)
    draw.polygon(right_poly, fill=right, outline=outline)
    draw.polygon(top_poly, fill=top, outline=outline)
    draw.line(top_poly + [top_poly[0]], fill=outline, width=3)
    draw.line((cx-w//2, cy+d, cx, cy+h//2+d, cx+w//2, cy+d), fill=_blend_hex(outline, "#FFFFFF", 0.2), width=2)


def _draw_toy_mine_house(draw, cx, cy, accent, dark, label, status, status_color, side="L"):
    # soft back mountain / rounded mine face
    body = _blend_hex(dark, "#FFFFFF", 0.13)
    roof = _blend_hex(accent, "#FFFFFF", 0.22)
    roof_dark = _blend_hex(accent, "#000000", 0.32)
    draw.rounded_rectangle((cx-185, cy-140, cx+185, cy+95), radius=55, fill=body, outline=accent, width=4)
    draw.polygon([(cx-235, cy-140), (cx, cy-235), (cx+235, cy-140), (cx+165, cy-98), (cx, cy-160), (cx-165, cy-98)], fill=roof_dark, outline=accent)
    draw.polygon([(cx-195, cy-155), (cx, cy-225), (cx+195, cy-155), (cx, cy-88)], fill=roof, outline=_blend_hex(accent, "#FFFFFF", 0.25))
    draw.rounded_rectangle((cx-75, cy-30, cx+75, cy+105), radius=42, fill="#081220", outline=_blend_hex(accent, "#FFFFFF", 0.35), width=4)
    draw.rounded_rectangle((cx-38, cy+0, cx+38, cy+107), radius=26, fill="#03070C")
    for dx in [-31, 31]:
        draw.line((cx+dx, cy+105, cx+dx*2, cy+190), fill="#D0A45D", width=5)
    for k in range(4):
        yy = cy + 128 + k * 18
        draw.line((cx-72+k*12, yy, cx+72-k*12, yy), fill="#E3BB73", width=3)
    draw.rounded_rectangle((cx-210, cy+124, cx+210, cy+178), radius=20, fill="#07111E", outline=accent, width=3)
    draw.text((cx, cy+151), label, font=get_font(22, True), fill=accent, anchor="mm")
    draw.rounded_rectangle((cx-170, cy+194, cx+170, cy+244), radius=18, fill="#081220", outline=status_color, width=3)
    draw.text((cx, cy+219), status, font=get_font(24, True), fill=status_color, anchor="mm")
    ox = cx-260 if side == "L" else cx+240
    oy = cy+220
    for i in range(18):
        px = ox + ((i * 33) % 132) - 64
        py = oy + ((i * 19) % 74) - 22
        draw.rounded_rectangle((px, py, px+42, py+25), radius=8, fill=_blend_hex(accent, "#FFFFFF", 0.10 + (i % 3) * 0.08), outline=_blend_hex(accent, "#000000", 0.18), width=2)
    _draw_worker(draw, ox-62, oy-58, hardhat="#FFD166", suit=accent)
    _draw_worker(draw, ox+74, oy+8, hardhat="#FFD166", suit="#38BDF8" if side == "L" else "#FF5A66")


def _draw_toy_core(draw, cx, cy, core_output, total_equity, cash):
    _draw_toy_platform(draw, cx, cy+260, 560, 360, 120, "#201C35", "#111022", "#151326", "#FFD166")
    for rr, col, width in [(360, "#FFE6A8", 5), (280, "#9B6B36", 5), (205, "#F2C572", 4), (130, "#FFE9A8", 5)]:
        draw.ellipse((cx-rr//2, cy-rr//2, cx+rr//2, cy+rr//2), outline=col, width=width)
    draw.ellipse((cx-92, cy-92, cx+92, cy+92), fill="#112117", outline="#FFE9A8", width=5)
    core_col = "#2EE6A6" if core_output >= 0 else "#FF5A66"
    for r in range(75, 10, -15):
        draw.ellipse((cx-r, cy-r, cx+r, cy+r), outline=_blend_hex(core_col, "#FFFFFF", (80-r)/90), width=4)
    _draw_glow_text(draw, (cx, cy-150), "MINE CORE · LIVE", get_font(34, True), "#FFE7A8", "mm", "#7A4B10", 2)
    _draw_glow_text(draw, (cx, cy-25), money(core_output), get_font(76, True), core_col, "mm", "#0C442F" if core_output >= 0 else "#42171B", 3)
    draw.text((cx, cy+82), f"TOTAL EQUITY {money(total_equity)}", font=get_font(29, True), fill="#EAF3FF", anchor="mm")
    draw.text((cx, cy+128), f"CASH {money(cash)}", font=get_font(25, True), fill="#92A8C7", anchor="mm")


def create_game_mine_dashboard(path: str, today_date: pd.Timestamp, summary: Dict[str, object], qld_signal: Dict[str, object], tqqq_signal: Dict[str, object], condition_df: pd.DataFrame, risk_df: pd.DataFrame) -> str:
    W, H = 2600, 1700
    img = Image.new("RGB", (W, H), "#050912")
    draw = ImageDraw.Draw(img)
    _draw_gradient_rect(draw, (0, 0, W, H), "#050912", "#0B1320")
    for i in range(230):
        x = (i * 173 + 41) % W; y = (i * 97 + 83) % H
        r = 1 + (i % 4)
        draw.ellipse((x, y, x+r, y+r), fill="#163B66" if i % 5 else "#A1722E")
    qld_status, qld_desc, qld_status_color = _safe_status(qld_signal)
    tqqq_status, tqqq_desc, tqqq_status_color = _safe_status(tqqq_signal)
    qld_pl = float(qld_signal.get("Position_Value", 0.0) or 0.0) - float(qld_signal.get("Shares", 0.0) or 0.0) * float(qld_signal.get("Avg_Price", 0.0) or 0.0)
    tqqq_pl = float(tqqq_signal.get("Position_Value", 0.0) or 0.0) - float(tqqq_signal.get("Shares", 0.0) or 0.0) * float(tqqq_signal.get("Avg_Price", 0.0) or 0.0)
    core_output = qld_pl + tqqq_pl
    total_risk = int(qld_signal.get("Risk_Score", 0) or 0) + int(tqqq_signal.get("Risk_Score", 0) or 0)
    mode = "RESTING" if total_risk < 8 else "GUARD MODE"
    mode_color = "#2EE6A6" if total_risk < 8 else "#FF5A66"
    draw.rounded_rectangle((50, 38, 935, 125), radius=26, fill="#0D1828", outline="#1A365D", width=2)
    draw.ellipse((78, 58, 130, 110), fill="#FFD166", outline="#FFF0AA", width=3)
    draw.text((104, 84), "Q/T", font=get_font(22, True), fill="#1A1A1A", anchor="mm")
    draw.text((160, 83), "DUAL ANCHOR MINE v11", font=get_font(40, True), fill="#EAF3FF", anchor="lm")
    draw.text((700, 83), today_date.strftime("%Y-%m-%d"), font=get_font(26, True), fill="#8FB2E8", anchor="lm")
    draw.rounded_rectangle((2120, 38, 2525, 125), radius=24, fill="#140D12" if total_risk >= 8 else "#0C1B18", outline=mode_color, width=2)
    draw.text((2160, 76), "●", font=get_font(30, True), fill=mode_color, anchor="lm")
    draw.text((2210, 72), mode, font=get_font(34, True), fill=mode_color, anchor="lm")
    draw.text((2210, 106), "리스크 오프 / 신규 진입 관리", font=get_font(18, True), fill="#B8C7DD", anchor="lm")
    _draw_toy_platform(draw, 500, 690, 760, 440, 140, "#174B3B", "#0F342A", "#0B261F", "#26E3A6")
    _draw_toy_platform(draw, 2100, 690, 760, 440, 140, "#65360F", "#3D210B", "#2B1607", "#FF8A1E")
    for dy in [0, 35, 70]:
        draw.line((820, 660+dy, 1035, 720+dy), fill="#C49145", width=9)
        draw.line((1780, 660+dy, 1565, 720+dy), fill="#C49145", width=9)
        draw.line((1035, 720+dy, 1115, 720+dy), fill="#E0B35C", width=4)
        draw.line((1485, 720+dy, 1565, 720+dy), fill="#E0B35C", width=4)
    _draw_toy_core(draw, 1300, 620, core_output, summary.get("Total_Equity", 0.0), summary.get("Cash", 0.0))
    _draw_toy_mine_house(draw, 500, 500, "#2EE6A6", "#143B32", "QLD SAFE MINE", qld_status, qld_status_color, side="L")
    _draw_toy_mine_house(draw, 2100, 500, "#FF9B2F", "#5B2C10", "TQQQ POWER MINE", tqqq_status, tqqq_status_color, side="R")
    def mine_panel(sig, x0, y0, accent, desc, status_color):
        w, h = 930, 285
        draw.rounded_rectangle((x0+14, y0+14, x0+w+14, y0+h+14), radius=30, fill="#03070D")
        draw.rounded_rectangle((x0, y0, x0+w, y0+h), radius=30, fill="#0C1523", outline=accent, width=2)
        draw.text((x0+35, y0+50), str(sig["Asset"]), font=get_font(38, True), fill=accent, anchor="lm")
        draw.text((x0+w-35, y0+50), order_korean(sig), font=get_font(34, True), fill="#EAF3FF", anchor="rm")
        draw.rounded_rectangle((x0+35, y0+82, x0+w-35, y0+132), radius=18, fill="#081220", outline=status_color, width=2)
        draw.text((x0+60, y0+107), desc, font=get_font(23, True), fill=status_color, anchor="lm")
        rows = [("평가금액", money(sig.get("Position_Value", 0.0))), ("위험점수", f"{sig.get('Risk_Score', 0)} / {FULL_SELL_SCORE}"), ("추천주문", f"{money(sig.get('Recommended_Amount', 0.0))} · {float(sig.get('Recommended_Shares', 0.0)):,.4f}주")]
        yy = y0 + 175
        for label, value in rows:
            draw.text((x0+45, yy), label, font=get_font(22, True), fill="#8FA5C3", anchor="lm")
            draw.text((x0+w-45, yy), value, font=get_font(24, True), fill="#EAF3FF", anchor="rm")
            yy += 38
    mine_panel(qld_signal, 100, 1055, "#2EE6A6", qld_desc, qld_status_color)
    mine_panel(tqqq_signal, 1570, 1055, "#FF9B2F", tqqq_desc, tqqq_status_color)
    draw.rounded_rectangle((50, 1408, W-50, 1600), radius=28, fill="#09111D", outline="#20314B", width=2)
    controls = [("CORE OUTPUT", money(core_output), "#2EE6A6" if core_output >= 0 else "#FF5A66"), ("QLD DRIFT", f"{float(qld_signal.get('Drift_%p', 0.0)):+.2f}%p", "#2EE6A6"), ("TQQQ DRIFT", f"{float(tqqq_signal.get('Drift_%p', 0.0)):+.2f}%p", "#FF9B2F"), ("TOTAL RISK", f"{total_risk} / {FULL_SELL_SCORE*2}", mode_color), ("SOLVER", "STANDBY" if total_risk < 8 else "GUARD", "#B794F4")]
    x = 85
    for label, value, color in controls:
        draw.rounded_rectangle((x, 1450, x+470, 1555), radius=18, fill="#0C1827", outline="#265167", width=2)
        draw.text((x+24, 1482), label, font=get_font(20, True), fill="#8FA5C3", anchor="lm")
        draw.text((x+24, 1528), value, font=get_font(34, True), fill=color, anchor="lm")
        x += 500
    draw.rounded_rectangle((210, 1612, 2390, 1660), radius=20, fill="#101B2B", outline="#1E3351", width=2)
    msg = f"GOLDIE AI: QLD는 {qld_status}, TQQQ는 {tqqq_status}. 매수는 3조건, 매도는 4조건 동시 충족을 기준으로 판단합니다."
    draw.text((245, 1636), msg, font=get_font(22, True), fill="#D9E7FF", anchor="lm")
    img.save(path)
    return path


def _account_daily_values(data: pd.DataFrame, asset: str, shares: float, idx) -> Tuple[float, float]:
    close_col = f"{asset}_Close"
    if close_col not in data.columns or idx not in data.index or shares <= 0:
        return 0.0, 0.0
    prev = data[close_col].shift(1).loc[idx]
    cur = data[close_col].loc[idx]
    if pd.isna(prev) or pd.isna(cur) or float(prev) == 0:
        return 0.0, 0.0
    pnl = (float(cur) - float(prev)) * shares
    ret = (float(cur) / float(prev) - 1.0) * 100.0
    return pnl, ret


def create_account_vibe_dashboard(path: str, today_date: pd.Timestamp, data: pd.DataFrame, signal: Dict[str, object], summary: Dict[str, object], theme: str = "SAFE") -> str:
    asset = str(signal["Asset"])
    safe = theme.upper() == "SAFE"
    accent = "#2EE6A6" if safe else "#FF8A1E"
    accent2 = "#1BBF87" if safe else "#B45D17"
    bg = "#070C14"; panel = "#0D1522"; line = "#1B2A3F"; muted = "#9BAEC9"; white = "#EAF3FF"; gold = "#FFD166"; red = "#FF5A66"
    W, H = 1160, 2480
    img = Image.new("RGB", (W, H), bg)
    draw = ImageDraw.Draw(img)
    _draw_gradient_rect(draw, (0, 0, W, H), "#070C14", "#0A111C")
    title = "SAFE 안정형 계좌" if safe else "POWER 공격형 계좌"
    shares = float(signal.get("Shares", 0.0) or 0.0)
    base = shares * float(signal.get("Avg_Price", 0.0) or 0.0)
    value = float(signal.get("Position_Value", 0.0) or 0.0)
    pnl = value - base if base > 0 else 0.0
    ret = pnl / base * 100 if base > 0 else 0.0
    draw.rounded_rectangle((45, 35, W-45, 125), radius=28, fill="#0D1828", outline=accent, width=2)
    draw.text((W//2, 80), f"{'POWER' if not safe else 'SAFE'} {asset} · {title}", font=get_font(38, True), fill=accent, anchor="mm")
    draw.text((W//2, 165), "누적 수익률 · 내 보유수량/평균단가 기준", font=get_font(27, True), fill=accent, anchor="mm")
    _draw_glow_text(draw, (W//2, 300), f"{ret:+.1f}%", get_font(112, True), accent if ret >= 0 else red, "mm", "#0E3B2E" if safe else "#4A240C", 2)
    draw.text((W//2, 420), "순이익 · 내 계좌 평가 기준", font=get_font(26, True), fill=muted, anchor="mm")
    draw.text((W//2, 490), money(pnl), font=get_font(64, True), fill=accent if pnl >= 0 else red, anchor="mm")
    y = 590
    close_col = f"{asset}_Close"
    month_data = data[(data.index.year == today_date.year) & (data.index.month == today_date.month)].copy() if close_col in data.columns else pd.DataFrame()
    rets = [_account_daily_values(data, asset, shares, idx)[1] for idx in month_data.tail(21).index]
    avg_daily = sum(rets) / len(rets) if rets else 0.0
    for i, (label, val) in enumerate([("수익률", f"{ret:+.1f}%"), ("일평균", f"{avg_daily:+.2f}%"), ("위험점수", f"{signal.get('Risk_Score', 0)}/{FULL_SELL_SCORE}")]):
        x0 = 70 + i * 350
        draw.rounded_rectangle((x0, y, x0+310, y+132), radius=26, fill=panel, outline=line, width=2)
        draw.text((x0+155, y+48), val, font=get_font(35, True), fill=white if i == 2 else accent if ret >= 0 else red, anchor="mm")
        draw.text((x0+155, y+95), label, font=get_font(22, True), fill=muted, anchor="mm")
    draw.text((W//2, 790), f"Verified · {asset} Anchor Engine", font=get_font(23, True), fill=accent, anchor="mm")
    draw.text((W//2, 825), f"{asset} {money(value)} · 내 계좌 기준 · rule based", font=get_font(22, True), fill=white, anchor="mm")
    draw.line((75, 865, W-75, 865), fill=line, width=2)
    draw.text((75, 925), "일별 성과 · 최근 5거래일 · 내 보유수량 기준", font=get_font(30, True), fill=muted, anchor="lm")
    recent = data[[close_col]].dropna().tail(5) if close_col in data.columns else pd.DataFrame()
    bar_vals = []
    for idx in recent.index:
        pnl_i, ret_i = _account_daily_values(data, asset, shares, idx)
        bar_vals.append((pd.Timestamp(idx).strftime("%m/%d"), pnl_i, ret_i))
    while len(bar_vals) < 5:
        bar_vals.insert(0, ("--", 0.0, 0.0))
    max_abs = max([abs(v[2]) for v in bar_vals] + [1.0])
    bar_base = 1160
    for i, (label_date, pnl_i, val) in enumerate(bar_vals[-5:]):
        x0 = 80 + i * 205
        bh = int(35 + min(130, abs(val) / max_abs * 120))
        c = accent if val >= 0 else red
        draw.rounded_rectangle((x0, bar_base-bh, x0+140, bar_base), radius=16, fill=c, outline=_blend_hex(c, "#FFFFFF", 0.25), width=1)
        draw.text((x0+70, bar_base-bh-50), money(pnl_i), font=get_font(19, True), fill=c, anchor="mm")
        draw.text((x0+70, bar_base-bh-24), f"{val:+.2f}%", font=get_font(20, True), fill=c, anchor="mm")
        draw.text((x0+70, bar_base+35), label_date, font=get_font(19, True), fill=muted, anchor="mm")
    y2 = 1240
    draw.rounded_rectangle((0, y2, W, y2+170), radius=0, fill="#0A111D", outline=line, width=2)
    draw.line((W//2, y2, W//2, y2+170), fill=line, width=2)
    buy_ready = "ON" if bool(signal.get("Buy_Ready")) else "OFF"
    sell_ready = "ON" if bool(signal.get("Sell_Ready")) else "OFF"
    draw.text((W//4, y2+62), buy_ready, font=get_font(58, True), fill=accent if buy_ready == "ON" else muted, anchor="mm")
    draw.text((W//4, y2+120), "매수 게이트", font=get_font(23, True), fill=muted, anchor="mm")
    draw.text((W*3//4, y2+62), sell_ready, font=get_font(58, True), fill=red if sell_ready == "ON" else gold, anchor="mm")
    draw.text((W*3//4, y2+120), "매도 게이트", font=get_font(23, True), fill=muted, anchor="mm")
    import calendar as _calendar
    cal_y = 1515
    draw.text((70, cal_y-58), today_date.strftime("%Y년 %m월 · 장 열린 날만 표시"), font=get_font(32, True), fill=white, anchor="lm")
    weekdays = ["월", "화", "수", "목", "금"]
    cell_w, cell_h = 190, 105
    gap_x, gap_y = 18, 17
    for cidx, wd in enumerate(weekdays):
        draw.text((70 + cidx*(cell_w+gap_x) + cell_w//2, cal_y-14), wd, font=get_font(20, True), fill=muted, anchor="mm")
    _, last_day = _calendar.monthrange(today_date.year, today_date.month)
    close_series = data[close_col].dropna() if close_col in data.columns else pd.Series(dtype=float)
    week_no = 0; last_week_key = None
    for d in range(1, last_day + 1):
        dts = pd.Timestamp(year=today_date.year, month=today_date.month, day=d)
        if dts.weekday() >= 5:
            continue
        week_key = (dts + pd.Timedelta(days=3-dts.weekday())).isocalendar().week
        if last_week_key is None:
            last_week_key = week_key
        elif week_key != last_week_key:
            week_no += 1; last_week_key = week_key
        x0 = 70 + dts.weekday() * (cell_w + gap_x)
        y0 = cal_y + week_no * (cell_h + gap_y)
        matched = [idx for idx in close_series.index if pd.Timestamp(idx).date() == dts.date()]
        if matched:
            pnl_i, ret_i = _account_daily_values(data, asset, shares, matched[-1])
            if ret_i >= 0:
                fill = _blend_hex("#0D1B21", accent2, min(0.62, 0.12 + abs(ret_i)/8)); outline = accent; txt = accent
            else:
                fill = _blend_hex("#171018", "#7A1D25", min(0.62, 0.12 + abs(ret_i)/8)); outline = red; txt = "#FF8A8A"
            pnl_text = money(pnl_i); ret_text = f"{ret_i:+.2f}%"
        else:
            fill = "#0A111D"; outline = "#243249"; txt = "#50627C"; pnl_text = "휴장"; ret_text = ""
        draw.rounded_rectangle((x0, y0, x0+cell_w, y0+cell_h), radius=16, fill=fill, outline=outline, width=1)
        draw.text((x0+14, y0+22), str(d), font=get_font(19, True), fill=muted, anchor="lm")
        draw.text((x0+14, y0+58), pnl_text, font=get_font(19, True), fill=txt, anchor="lm")
        if ret_text:
            draw.text((x0+14, y0+84), ret_text, font=get_font(18, True), fill=txt, anchor="lm")
    reason = str(signal.get("Reason", ""))
    draw.rounded_rectangle((45, H-150, W-45, H-45), radius=24, fill="#101B2B", outline=line, width=2)
    for i, ln in enumerate(wrap_lines(reason, 57)[:2] if reason else ["조건 확인 완료"]):
        draw.text((75, H-110+i*34), ln, font=get_font(22, True), fill=white, anchor="lm")
    img.save(path)
    return path


def create_risk_gate_dashboard(path: str, today_date: pd.Timestamp, condition_df: pd.DataFrame, risk_df: pd.DataFrame, qld_signal: Dict[str, object], tqqq_signal: Dict[str, object]) -> str:
    W, H = 1850, 2850
    img = Image.new("RGB", (W, H), "#070C14")
    draw = ImageDraw.Draw(img)
    _draw_gradient_rect(draw, (0, 0, W, H), "#070C14", "#0A111C")
    white = "#EAF3FF"; muted = "#9BAEC9"; green = "#2EE6A6"; orange = "#FF8A1E"; red = "#FF5A66"; gold = "#FFD166"; panel = "#0D1522"; line = "#1B2A3F"
    active_count = int(condition_df["충족"].sum()) if "충족" in condition_df.columns else 0
    total_count = len(condition_df)
    draw.text((55, 80), "BUY / SELL RISK GATES", font=get_font(52, True), fill=white, anchor="lm")
    draw.rounded_rectangle((1260, 35, 1760, 120), radius=24, fill="#101823", outline=gold, width=2)
    draw.text((1510, 77), f"{active_count}/{total_count} 조건 충족", font=get_font(32, True), fill=gold, anchor="mm")
    draw.text((55, 145), "매수/매도 조건과 전체 위험점수 규칙을 함께 표시합니다. 빨간색 위험규칙만 현재 점수에 반영됩니다.", font=get_font(24, True), fill=muted, anchor="lm")
    def gate_section(asset: str, sig: Dict[str, object], y0: int, accent: str):
        section_h = 1260
        draw.rounded_rectangle((45, y0, W-45, y0+section_h), radius=34, fill=panel, outline=line, width=2)
        draw.text((90, y0+55), f"{asset} CONTROL GATES", font=get_font(38, True), fill=accent, anchor="lm")
        draw.text((W-90, y0+55), order_korean(sig), font=get_font(34, True), fill=white, anchor="rm")
        sub = condition_df[condition_df["자산"] == asset].copy()
        buy = sub[sub["구분"] == "매수"].copy(); sell = sub[sub["구분"] == "매도"].copy()
        risk_all = risk_df[risk_df["자산"] == asset].copy()
        risk_score = int(sig.get("Risk_Score", 0) or 0)
        draw.rounded_rectangle((90, y0+120, W-90, y0+345), radius=28, fill="#0B1F1B", outline=green, width=2)
        draw.text((125, y0+165), "BUY GATE · 진입 허가 3조건", font=get_font(30, True), fill=green, anchor="lm")
        yy = y0 + 220
        for _, r in buy.iterrows():
            ok = bool(r["충족"]); c = green if ok else red
            draw.ellipse((125, yy-14, 155, yy+16), fill=c)
            draw.text((175, yy), str(r["조건"]), font=get_font(24, True), fill=white, anchor="lm")
            draw.text((W-140, yy), f"{num(r['현재값'],2)} {r['부등호']} {num(r['기준값'],2)} · {'충족' if ok else '미충족'}", font=get_font(21, True), fill=c, anchor="rm")
            yy += 50
        draw.rounded_rectangle((90, y0+380, W-90, y0+660), radius=28, fill="#241016", outline=red, width=2)
        draw.text((125, y0+425), "SELL GUARD · 방어 매도 4조건", font=get_font(30, True), fill=red, anchor="lm")
        yy = y0 + 480
        for _, r in sell.iterrows():
            ok = bool(r["충족"]); c = red if ok else green
            draw.ellipse((125, yy-14, 155, yy+16), fill=c)
            draw.text((175, yy), str(r["조건"]), font=get_font(24, True), fill=white, anchor="lm")
            draw.text((W-140, yy), f"{num(r['현재값'],2)} {r['부등호']} {num(r['기준값'],2)} · {'발동' if ok else '미발동'}", font=get_font(21, True), fill=c, anchor="rm")
            yy += 50
        mx0, my0, mx1, my1 = 120, y0+755, W-120, y0+815
        draw.text((120, y0+725), f"위험점수 {risk_score}/{FULL_SELL_SCORE}", font=get_font(28, True), fill=white, anchor="lm")
        draw.rounded_rectangle((mx0, my0, mx1, my1), radius=28, fill="#1A2230", outline=line, width=2)
        fill_w = int((mx1 - mx0) * min(1.0, risk_score / FULL_SELL_SCORE))
        meter_color = green if risk_score < 4 else gold if risk_score < FULL_SELL_SCORE else red
        if fill_w > 0:
            draw.rounded_rectangle((mx0, my0, mx0+fill_w, my1), radius=28, fill=meter_color)
        draw.text((mx1, y0+725), "NORMAL" if risk_score < 4 else "WATCH" if risk_score < FULL_SELL_SCORE else "BLOCK", font=get_font(28, True), fill=meter_color, anchor="rm")
        draw.text((90, y0+900), "위험점수 전체 규칙 · 빨간색만 현재 점수에 반영", font=get_font(28, True), fill=white, anchor="lm")
        for idx, r in enumerate(risk_all.to_dict("records")[:12]):
            col = idx // 6; row = idx % 6
            x = 105 + col * 870; yy = y0 + 960 + row * 48
            pts = int(r.get("점수", 0) or 0)
            active = pts > 0
            c = red if active else green
            draw.ellipse((x, yy-13, x+26, yy+13), fill=c)
            draw.text((x+42, yy), f"{('+'+str(pts)) if active else '+0'}  {r.get('항목','')}", font=get_font(21, True), fill=white if active else muted, anchor="lm")
            draw.text((x+470, yy), str(r.get("설명", ""))[:27], font=get_font(19), fill=muted, anchor="lm")
    gate_section("QLD", qld_signal, 220, green)
    gate_section("TQQQ", tqqq_signal, 1530, orange)
    img.save(path)
    return path

def build_telegram_text(today_date: pd.Timestamp, summary: Dict[str, object], qld_signal: Dict[str, object], tqqq_signal: Dict[str, object]) -> str:
    lines = []
    lines.append("[DUAL QLD + TQQQ 앵커 신호]")
    lines.append(f"날짜: {today_date.strftime('%Y-%m-%d')}")
    lines.append(f"총자산: {money(summary['Total_Equity'])} / 현금: {money(summary['Cash'])}")
    lines.append(f"목표비중: QLD {summary['QLD_Target_Weight']*100:.1f}% / TQQQ {summary['TQQQ_Target_Weight']*100:.1f}% / 현금 {summary['Cash_Target_Weight']*100:.1f}%")
    for sig in [qld_signal, tqqq_signal]:
        lines.append("")
        lines.append(f"[{sig['Asset']}] 종가판단: {order_korean(sig)}")
        lines.append(f"실시간상태: {sig.get('Live_Status_KR', '미확인')} / 최종행동: {sig.get('Final_Order_Action_KR', order_korean(sig))}")
        lines.append(f"종가: {money(sig['Price'])} / 현재가확인: {money(sig.get('Live_Price', sig['Price']))} ({float(sig.get('Live_Diff_%', 0.0)):+.2f}%)")
        lines.append(f"보유: {float(sig['Shares']):,.6f}주 / 평가: {money(sig['Position_Value'])}")
        lines.append(f"현재비중: {float(sig['Current_Weight'])*100:.2f}% / 목표비중: {float(sig['Target_Weight'])*100:.2f}%")
        lines.append(f"종가기준 주문: {money(sig['Recommended_Amount'])} / {float(sig['Recommended_Shares']):,.6f}주")
        lines.append(f"실시간 최종주문: {money(sig.get('Live_Order_Amount', 0.0))} / {float(sig.get('Live_Order_Shares', 0.0)):,.6f}주")
        lines.append(f"위험점수: {sig['Risk_Score']} / {FULL_SELL_SCORE} / 실시간점수: {sig.get('Live_Score', 0)}")
        lines.append(f"사유: {sig.get('Final_Order_Text', sig['Reason'])}")
    lines.append("")
    lines.append(f"실시간 확인: {summary.get('Realtime_Checked_At', '')}")
    lines.append("※ 종가 판단은 백테스트 기준 앵커이고, 실시간 상태는 주문 직전 실행 필터/비상 브레이크입니다. 실제 주문 전 증권사 현재가, 수수료, 세금, 계좌 상황을 직접 확인하세요.")
    return "\n".join(lines)


def send_telegram_message(bot_token: str, chat_id: str, text: str) -> bool:
    try:
        url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
        r = requests.post(url, data={"chat_id": chat_id, "text": text}, timeout=15)
        js = r.json()
        if r.status_code == 200 and js.get("ok"):
            print("[telegram] 메시지 전송 완료")
            return True
        print("[telegram] 메시지 전송 실패:", js)
    except Exception as e:
        print("[telegram] 메시지 오류:", e)
    return False


def send_telegram_photo(bot_token: str, chat_id: str, image_path: str, caption: str = "") -> bool:
    try:
        url = f"https://api.telegram.org/bot{bot_token}/sendPhoto"
        with open(image_path, "rb") as f:
            r = requests.post(url, data={"chat_id": chat_id, "caption": caption}, files={"photo": f}, timeout=30)
        js = r.json()
        if r.status_code == 200 and js.get("ok"):
            print(f"[telegram] 사진 전송 완료: {image_path}")
            return True
        print("[telegram] 사진 전송 실패:", js)
    except Exception as e:
        print("[telegram] 사진 오류:", e)
    return False


def create_graphs(data: pd.DataFrame, out_dir: str) -> Tuple[str, str, str]:
    c = data.tail(260).copy()

    p1 = os.path.join(out_dir, "dual_qld_tqqq_normalized.png")
    norm = pd.DataFrame(index=c.index)
    for asset in ["QQQ", "QLD", "TQQQ"]:
        norm[asset] = c[f"{asset}_Close"] / c[f"{asset}_Close"].iloc[0] * 100
    plt.figure(figsize=(16, 8))
    for asset in ["QQQ", "QLD", "TQQQ"]:
        plt.plot(norm.index, norm[asset], linewidth=2.2, label=f"{asset} indexed to 100")
    plt.axhline(100, linewidth=1.2, linestyle="--")
    plt.title("QQQ / QLD / TQQQ Normalized Trend", fontsize=20, fontweight="bold", pad=18)
    plt.grid(True, alpha=0.25)
    plt.legend(loc="upper left")
    plt.tight_layout()
    plt.savefig(p1, dpi=180, bbox_inches="tight")
    plt.close()

    p2 = os.path.join(out_dir, "dual_asset_ma.png")
    plt.figure(figsize=(16, 8))
    plt.plot(c.index, c["QLD_Close"], linewidth=2.2, label="QLD Close")
    plt.plot(c.index, c["QLD_MA20"], linewidth=1.4, label="QLD MA20")
    plt.plot(c.index, c["QLD_MA100"], linewidth=1.4, label="QLD MA100")
    plt.plot(c.index, c["TQQQ_Close"], linewidth=2.2, label="TQQQ Close")
    plt.plot(c.index, c["TQQQ_MA20"], linewidth=1.4, label="TQQQ MA20")
    plt.plot(c.index, c["TQQQ_MA100"], linewidth=1.4, label="TQQQ MA100")
    plt.title("QLD / TQQQ Price and Moving Averages", fontsize=20, fontweight="bold", pad=18)
    plt.grid(True, alpha=0.25)
    plt.legend(loc="upper left")
    plt.tight_layout()
    plt.savefig(p2, dpi=180, bbox_inches="tight")
    plt.close()

    p3 = os.path.join(out_dir, "dual_momentum.png")
    plt.figure(figsize=(16, 8))
    plt.plot(c.index, c["QLD_MACD_HIST"], linewidth=2.0, label="QLD MACD Hist")
    plt.plot(c.index, c["TQQQ_MACD_HIST"], linewidth=2.0, label="TQQQ MACD Hist")
    plt.axhline(0, linewidth=1.2)
    plt.title("QLD / TQQQ MACD Histogram", fontsize=20, fontweight="bold", pad=18)
    plt.grid(True, alpha=0.25)
    plt.legend(loc="upper left")
    plt.tight_layout()
    plt.savefig(p3, dpi=180, bbox_inches="tight")
    plt.close()

    return p1, p2, p3


# ======================================================
# Excel report
# ======================================================

def build_excel_report(path: str, today_date: pd.Timestamp, summary_df: pd.DataFrame, signals_df: pd.DataFrame, condition_df: pd.DataFrame, risk_df: pd.DataFrame, recent_df: pd.DataFrame, dashboard_path: str, graph_paths: Tuple[str, str, str], realtime_df: Optional[pd.DataFrame] = None, realtime_dashboard_path: str = "", performance_df: Optional[pd.DataFrame] = None, trade_df: Optional[pd.DataFrame] = None, performance_dashboard_path: str = "") -> str:
    with pd.ExcelWriter(path, engine="xlsxwriter", datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as writer:
        wb = writer.book
        fmt_title = wb.add_format({"bold": True, "font_size": 20, "font_color": "white", "bg_color": "#0B2A52", "align": "center", "valign": "vcenter"})
        fmt_section = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#1E3A8A", "border": 1, "align": "center", "valign": "vcenter"})
        fmt_header = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#334155", "border": 1, "align": "center", "valign": "vcenter"})
        fmt_good = wb.add_format({"bg_color": "#DCFCE7", "font_color": "#166534", "bold": True, "border": 1, "align": "center"})
        fmt_bad = wb.add_format({"bg_color": "#FEE2E2", "font_color": "#991B1B", "bold": True, "border": 1, "align": "center"})
        fmt_wait = wb.add_format({"bg_color": "#FEF3C7", "font_color": "#92400E", "bold": True, "border": 1, "align": "center"})
        fmt_cell = wb.add_format({"border": 1, "valign": "vcenter"})
        fmt_wrap = wb.add_format({"border": 1, "text_wrap": True, "valign": "top"})
        fmt_money = wb.add_format({"border": 1, "num_format": "$#,##0.00"})
        fmt_pct = wb.add_format({"border": 1, "num_format": "0.00%"})
        fmt_num = wb.add_format({"border": 1, "num_format": "#,##0.00"})

        ws = wb.add_worksheet("00_Dashboard")
        writer.sheets["00_Dashboard"] = ws
        ws.hide_gridlines(2)
        ws.set_zoom(75)
        ws.set_column("A:A", 3)
        ws.set_column("B:N", 16)
        ws.merge_range("B2:N3", "DUAL QLD + TQQQ 통합 앵커 현황판", fmt_title)
        if os.path.exists(dashboard_path):
            ws.insert_image("B5", dashboard_path, {"x_scale": 0.48, "y_scale": 0.48, "object_position": 1})

        summary_df.to_excel(writer, sheet_name="01_Account_Summary", index=False)
        ws1 = writer.sheets["01_Account_Summary"]
        ws1.hide_gridlines(2)
        ws1.set_column("A:Z", 18)
        for col, header in enumerate(summary_df.columns):
            ws1.write(0, col, header, fmt_header)

        signals_df.to_excel(writer, sheet_name="02_Order_Guide", index=False)
        ws2 = writer.sheets["02_Order_Guide"]
        ws2.hide_gridlines(2)
        ws2.freeze_panes(1, 0)
        ws2.set_column("A:B", 16)
        ws2.set_column("C:C", 22)
        ws2.set_column("D:G", 16)
        ws2.set_column("H:H", 58)
        ws2.set_column("I:Z", 16)
        for col, header in enumerate(signals_df.columns):
            ws2.write(0, col, header, fmt_header)
        if len(signals_df) > 0:
            action_col = list(signals_df.columns).index("Order_Action")
            ws2.conditional_format(1, action_col, len(signals_df), action_col, {"type": "text", "criteria": "containing", "value": "BUY", "format": fmt_good})
            ws2.conditional_format(1, action_col, len(signals_df), action_col, {"type": "text", "criteria": "containing", "value": "SELL", "format": fmt_bad})
            ws2.conditional_format(1, action_col, len(signals_df), action_col, {"type": "text", "criteria": "containing", "value": "WAIT", "format": fmt_wait})

        condition_df.to_excel(writer, sheet_name="03_Condition_Check", index=False)
        ws3 = writer.sheets["03_Condition_Check"]
        ws3.hide_gridlines(2)
        ws3.freeze_panes(1, 0)
        ws3.set_column("A:C", 16)
        ws3.set_column("D:D", 10)
        ws3.set_column("E:H", 16)
        ws3.set_column("I:K", 28)
        for col, header in enumerate(condition_df.columns):
            ws3.write(0, col, header, fmt_header)
        if len(condition_df) > 0:
            yn_col = list(condition_df.columns).index("충족여부")
            ws3.conditional_format(1, yn_col, len(condition_df), yn_col, {"type": "text", "criteria": "containing", "value": "충족", "format": fmt_good})
            ws3.conditional_format(1, yn_col, len(condition_df), yn_col, {"type": "text", "criteria": "containing", "value": "미충족", "format": fmt_bad})

        risk_df.to_excel(writer, sheet_name="04_Risk_Score", index=False)
        ws4 = writer.sheets["04_Risk_Score"]
        ws4.hide_gridlines(2)
        ws4.freeze_panes(1, 0)
        ws4.set_column("A:C", 18)
        ws4.set_column("D:E", 16)
        ws4.set_column("F:G", 12)
        ws4.set_column("H:H", 42)
        for col, header in enumerate(risk_df.columns):
            ws4.write(0, col, header, fmt_header)
        if len(risk_df) > 0:
            score_col = list(risk_df.columns).index("점수")
            ws4.conditional_format(1, score_col, len(risk_df), score_col, {"type": "cell", "criteria": ">", "value": 0, "format": fmt_bad})

        recent_df.to_excel(writer, sheet_name="05_Recent_Data", index=False)
        ws5 = writer.sheets["05_Recent_Data"]
        ws5.hide_gridlines(2)
        ws5.freeze_panes(1, 1)
        ws5.set_column("A:A", 13)
        ws5.set_column("B:Z", 14)
        for col, header in enumerate(recent_df.columns):
            ws5.write(0, col, header, fmt_header)

        ws6 = wb.add_worksheet("06_Graphs")
        writer.sheets["06_Graphs"] = ws6
        ws6.hide_gridlines(2)
        ws6.set_zoom(70)
        ws6.set_column("A:A", 3)
        ws6.set_column("B:N", 16)
        ws6.merge_range("B2:N3", "그래프 이미지", fmt_title)
        titles = ["QQQ/QLD/TQQQ 정규화 추세", "QLD/TQQQ 이동평균", "QLD/TQQQ MACD"]
        rows = [5, 32, 59]
        for i, gp in enumerate(graph_paths):
            ws6.merge_range(rows[i] - 1, 1, rows[i] - 1, 13, titles[i], fmt_section)
            if os.path.exists(gp):
                ws6.insert_image(rows[i], 1, gp, {"x_scale": 0.58, "y_scale": 0.58, "object_position": 1})

        if realtime_df is not None and not realtime_df.empty:
            realtime_df.to_excel(writer, sheet_name="08_Realtime_Order_Check", index=False)
            ws8 = writer.sheets["08_Realtime_Order_Check"]
            ws8.hide_gridlines(2)
            ws8.freeze_panes(1, 0)
            ws8.set_column("A:C", 18)
            ws8.set_column("D:N", 20)
            ws8.set_column("O:Q", 28)
            for col, header in enumerate(realtime_df.columns):
                ws8.write(0, col, header, fmt_header)
            ws8img = wb.add_worksheet("09_Realtime_Dashboard")
            writer.sheets["09_Realtime_Dashboard"] = ws8img
            ws8img.hide_gridlines(2)
            ws8img.set_zoom(70)
            ws8img.merge_range("B2:N3", "실시간 주문 전 확인 현황판", fmt_title)
            if realtime_dashboard_path and os.path.exists(realtime_dashboard_path):
                ws8img.insert_image("B5", realtime_dashboard_path, {"x_scale": 0.52, "y_scale": 0.52, "object_position": 1})

        if performance_df is not None and not performance_df.empty:
            performance_df.to_excel(writer, sheet_name="10_Performance_History", index=False)
            ws10 = writer.sheets["10_Performance_History"]
            ws10.hide_gridlines(2)
            ws10.freeze_panes(1, 0)
            ws10.set_column("A:B", 20)
            ws10.set_column("C:Z", 18)
            for col, header in enumerate(performance_df.columns):
                ws10.write(0, col, header, fmt_header)

        if trade_df is not None and not trade_df.empty:
            trade_df.to_excel(writer, sheet_name="11_Trade_Log", index=False)
            ws11 = writer.sheets["11_Trade_Log"]
            ws11.hide_gridlines(2)
            ws11.freeze_panes(1, 0)
            ws11.set_column("A:K", 20)
            for col, header in enumerate(trade_df.columns):
                ws11.write(0, col, header, fmt_header)

        if performance_dashboard_path and os.path.exists(performance_dashboard_path):
            wsp = wb.add_worksheet("12_Performance_Dashboard")
            writer.sheets["12_Performance_Dashboard"] = wsp
            wsp.hide_gridlines(2)
            wsp.set_zoom(70)
            wsp.merge_range("B2:N3", "누적 성과 현황판", fmt_title)
            wsp.insert_image("B5", performance_dashboard_path, {"x_scale": 0.52, "y_scale": 0.52, "object_position": 1})

        guide = pd.DataFrame([
            ["BUY", "미보유 또는 목표비중 미달 상태에서 매수조건이 충족되어 매수 후보입니다."],
            ["BUY_MORE", "이미 보유 중이지만 목표비중보다 낮고 매수조건이 충족되어 추가매수 후보입니다."],
            ["SELL_ALL", "전략 매도 4조건이 모두 충족되어 전량매도 후보입니다."],
            ["HOLD_OVER_TARGET", "목표비중보다 높아도 매도 4조건이 아니면 팔지 않고 보유합니다."],
            ["HOLD", "보유 중이며 전략 매도조건이 충족되지 않아 보유유지입니다."],
            ["WAIT", "미보유 상태에서 매수조건이 미충족되어 대기입니다."],
            ["목표비중", "예: --qld-weight 0.6 --tqqq-weight 0.4 는 총자산 중 QLD 60%, TQQQ 40%를 목표로 합니다."],
            ["실시간 확인", "08/09 시트와 realtime_order_check PNG는 주문 전 가격/수량 확인용입니다. 전략 신호 자체는 종가 기준입니다."],
            ["누적 성과", "10/12 시트와 performance_dashboard PNG는 실행 후 입력한 현금/보유수량/평균단가를 누적 스냅샷으로 저장해 성과를 보여줍니다."],
            ["체결 기록", "--record-trade 와 --qld-exec-action/price/shares 등을 입력하면 실제 체결 로그를 11_Trade_Log에 누적합니다."],
        ], columns=["항목", "설명"])
        guide.to_excel(writer, sheet_name="07_Guide", index=False)
        ws7 = writer.sheets["07_Guide"]
        ws7.hide_gridlines(2)
        ws7.set_column("A:A", 20)
        ws7.set_column("B:B", 100)
        for col, header in enumerate(guide.columns):
            ws7.write(0, col, header, fmt_header)
        for r in range(1, len(guide) + 1):
            ws7.set_row(r, 42)
            ws7.write(r, 1, guide.iloc[r - 1, 1], fmt_wrap)

    return path





# ======================================================
# v13 AI background overlay dashboards
# ======================================================

AI_MINE_BACKGROUND_FILES = [
    "ai_mine_bg_01_classic_balanced.png",
    "ai_mine_bg_02_classic_dynamic.png",
    "ai_mine_bg_03_classic_floating_core.png",
    "ai_mine_bg_04_classic_clean_future.png",
    "ai_mine_bg_05_cute_ui_hub.png",
    "ai_mine_bg_06_cute_space_islands.png",
    "ai_mine_bg_07_cute_ui_slots.png",
    "ai_mine_bg_08_cute_core_diorama.png",
]


def _list_image_files(folder: str) -> List[str]:
    exts = {".png", ".jpg", ".jpeg", ".webp"}
    if not folder or not os.path.isdir(folder):
        return []
    files = [p.path for p in sorted(os.scandir(folder), key=lambda e: e.name) if os.path.splitext(p.name)[1].lower() in exts and p.is_file()]
    return files


def resolve_ai_mine_backgrounds(bg_dir: Optional[str] = None) -> List[str]:
    """Find AI-generated mine backgrounds for v13 overlays.

    Preferred share layout:
    - DUAL_QLD_TQQQ_anchor_dashboard_v13_ai_mine_8bg_overlay_colab.py
    - ai_mine_backgrounds_v13/*.png

    Behavior
    - If standardized v13 files exist, load them in that fixed order.
    - Otherwise, load every image file found in the candidate folder in filename order.
    - Keeps backward compatibility with the old v12 4-image folder.
    """
    candidates: List[str] = []
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    except Exception:
        script_dir = os.getcwd()

    if bg_dir:
        candidates.append(bg_dir if os.path.isabs(bg_dir) else os.path.join(script_dir, bg_dir))

    candidates += [
        os.path.join(script_dir, "ai_mine_backgrounds_v13"),
        os.path.join(os.getcwd(), "ai_mine_backgrounds_v13"),
        "/mnt/data/ai_mine_backgrounds_v13",
        os.path.join(script_dir, "ai_mine_backgrounds_v12"),
        os.path.join(os.getcwd(), "ai_mine_backgrounds_v12"),
        "/mnt/data/ai_mine_backgrounds_v12",
    ]

    for folder in candidates:
        if not folder or not os.path.isdir(folder):
            continue
        std_paths = [os.path.join(folder, name) for name in AI_MINE_BACKGROUND_FILES]
        if all(os.path.exists(x) for x in std_paths):
            return std_paths
        any_images = _list_image_files(folder)
        if any_images:
            return any_images
    return []


def _cover_resize(img: Image.Image, size: Tuple[int, int]) -> Image.Image:
    """Resize/crop to fill target size without distortion."""
    target_w, target_h = size
    src_w, src_h = img.size
    scale = max(target_w / src_w, target_h / src_h)
    new_w = int(src_w * scale + 0.5)
    new_h = int(src_h * scale + 0.5)
    img = img.resize((new_w, new_h), Image.LANCZOS)
    left = max(0, (new_w - target_w) // 2)
    top = max(0, (new_h - target_h) // 2)
    return img.crop((left, top, left + target_w, top + target_h))


def _rgba_panel(base: Image.Image, xy, radius: int, fill_rgba, outline_rgba=None, width: int = 2):
    overlay = Image.new("RGBA", base.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)
    d.rounded_rectangle(xy, radius=radius, fill=fill_rgba, outline=outline_rgba, width=width)
    return Image.alpha_composite(base, overlay)


def _draw_pill(draw: ImageDraw.ImageDraw, xy, text_value: str, fill: str, outline: str, font_size: int = 26):
    draw.rounded_rectangle(xy, radius=22, fill=fill, outline=outline, width=2)
    draw.text(((xy[0] + xy[2]) // 2, (xy[1] + xy[3]) // 2), text_value, font=get_font(font_size, True), fill="#EAF3FF", anchor="mm")


def _signal_state_label(sig: Dict[str, object]) -> Tuple[str, str, str]:
    oa = str(sig.get("Order_Action", ""))
    risk = int(sig.get("Risk_Score", 0) or 0)
    if oa in ["BUY", "BUY_MORE"]:
        return "BUY READY", "매수 조건 충족", "#2EE6A6"
    if oa == "SELL_ALL":
        return "SELL GUARD", "전량매도 조건 발동", "#FF5A66"
    if oa == "HOLD":
        return ("HOLD WATCH", "보유 · 위험 감시", "#FFD166") if risk >= FULL_SELL_SCORE - 2 else ("HOLD", "보유 유지", "#2EE6A6")
    return "WAIT", "대기 · 조건 미충족", "#8FB2E8"


def create_ai_mine_overlay_dashboard(path: str, background_path: str, variant_no: int, today_date: pd.Timestamp, summary: Dict[str, object], qld_signal: Dict[str, object], tqqq_signal: Dict[str, object]) -> str:
    """Overlay richer account/strategy data on the AI-generated 3D mine background.

    v15-ui: this replaces the old hand-drawn game_mine_control_room image as the
    main mine-style Telegram image. Strategy calculation is not changed.
    """
    W, H = 1920, 1080
    bg = Image.open(background_path).convert("RGB")
    bg = _cover_resize(bg, (W, H)).convert("RGBA")

    # Soft dark overlays so text stays readable on bright AI backgrounds.
    vignette = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    vd = ImageDraw.Draw(vignette)
    vd.rectangle((0, 0, W, 170), fill=(2, 7, 16, 145))
    vd.rectangle((0, 660, W, H), fill=(2, 7, 16, 128))
    vd.rectangle((0, 0, 50, H), fill=(2, 7, 16, 70))
    vd.rectangle((W - 50, 0, W, H), fill=(2, 7, 16, 70))
    bg = Image.alpha_composite(bg, vignette)

    qld_state, qld_sub, qld_color = _signal_state_label(qld_signal)
    tqqq_state, tqqq_sub, tqqq_color = _signal_state_label(tqqq_signal)
    qld_pl = float(qld_signal.get("Position_Value", 0.0) or 0.0) - float(qld_signal.get("Shares", 0.0) or 0.0) * float(qld_signal.get("Avg_Price", 0.0) or 0.0)
    tqqq_pl = float(tqqq_signal.get("Position_Value", 0.0) or 0.0) - float(tqqq_signal.get("Shares", 0.0) or 0.0) * float(tqqq_signal.get("Avg_Price", 0.0) or 0.0)
    total_pl = qld_pl + tqqq_pl
    today_pl = float(summary.get("Today_PnL", 0.0) or 0.0)
    total_risk = int(qld_signal.get("Risk_Score", 0) or 0) + int(tqqq_signal.get("Risk_Score", 0) or 0)
    risk_color = "#2EE6A6" if total_risk < 8 else "#FFD166" if total_risk < 12 else "#FF5A66"
    solver = "STANDBY" if total_risk < 8 else "GUARD"
    solver_color = "#2EE6A6" if solver == "STANDBY" else "#B794F4"

    bg = _rgba_panel(bg, (42, 28, 780, 126), 28, (5, 12, 24, 190), (82, 116, 164, 175), 2)
    bg = _rgba_panel(bg, (1505, 28, 1878, 126), 28, (5, 12, 24, 178), (255, 209, 102, 170), 2)
    draw = ImageDraw.Draw(bg)
    draw.text((78, 68), "DUAL QLD + TQQQ ANCHOR MINE", font=get_font(36, True), fill="#EAF3FF", anchor="lm")
    draw.text((78, 105), "AI 광산 메인 · 계좌/주문/위험 통합", font=get_font(19, True), fill="#9DB4D5", anchor="lm")
    draw.text((1690, 60), today_date.strftime("%Y-%m-%d"), font=get_font(30, True), fill="#FFE7A8", anchor="mm")
    draw.text((1690, 100), "전략 판단 기준: 종가", font=get_font(19, True), fill="#B8C7DD", anchor="mm")

    # Center core box, reusing the information from the old game_mine_control_room.
    bg = _rgba_panel(bg, (690, 628, 1230, 792), 32, (5, 12, 24, 165), (255, 225, 160, 175), 2)
    draw = ImageDraw.Draw(bg)
    draw.text((960, 672), "MINE CORE OUTPUT", font=get_font(24, True), fill="#FFE7A8", anchor="mm")
    draw.text((960, 720), money(total_pl), font=get_font(44, True), fill="#2EE6A6" if total_pl >= 0 else "#FF5A66", anchor="mm")
    draw.text((960, 760), f"TOTAL {money(summary.get('Total_Equity', 0.0))} · CASH {money(summary.get('Cash', 0.0))}", font=get_font(21, True), fill="#D9E7FF", anchor="mm")

    _draw_pill(draw, (70, 710, 410, 762), f"QLD · {qld_state}", "#082119", "#2EE6A6", 22)
    _draw_pill(draw, (1510, 710, 1850, 762), f"TQQQ · {tqqq_state}", "#251407", "#FF9B2F", 22)

    # Bottom control room panel.
    bg = _rgba_panel(bg, (44, 782, 1876, 1042), 34, (3, 8, 17, 220), (72, 92, 130, 165), 2)
    draw = ImageDraw.Draw(bg)
    kpis = [
        ("총자산", money(summary.get("Total_Equity", 0.0)), "#EAF3FF"),
        ("현금", money(summary.get("Cash", 0.0)), "#EAF3FF"),
        ("오늘 손익", money(today_pl), "#2EE6A6" if today_pl >= 0 else "#FF5A66"),
        ("총 위험점수", f"{total_risk} / {FULL_SELL_SCORE * 2}", risk_color),
        ("SOLVER", solver, solver_color),
    ]
    x = 88
    widths = [330, 330, 330, 330, 260]
    for (label, value, color), ww in zip(kpis, widths):
        draw.rounded_rectangle((x, 812, x + ww, 900), radius=24, fill="#0B1625", outline="#263A58", width=2)
        draw.text((x + 24, 840), label, font=get_font(18, True), fill="#8FA5C3", anchor="lm")
        draw.text((x + 24, 878), value, font=get_font(27, True), fill=color, anchor="lm")
        x += ww + 28

    def asset_box(sig: Dict[str, object], x0: int, y0: int, accent: str, state: str, sub: str):
        w, h = 840, 112
        oa_text = order_korean(sig)
        live_text = str(sig.get("Live_Status_KR", "실시간 미확인"))
        draw.rounded_rectangle((x0, y0, x0 + w, y0 + h), radius=24, fill="#0B1625", outline=accent, width=2)
        draw.text((x0 + 30, y0 + 34), str(sig.get("Asset", "")), font=get_font(30, True), fill=accent, anchor="lm")
        draw.text((x0 + 126, y0 + 34), state, font=get_font(25, True), fill="#EAF3FF", anchor="lm")
        draw.text((x0 + w - 28, y0 + 34), oa_text, font=get_font(26, True), fill="#EAF3FF", anchor="rm")
        draw.text((x0 + 30, y0 + 72), sub, font=get_font(19, True), fill="#AFC1DA", anchor="lm")
        draw.text((x0 + 30, y0 + 99), f"평가 {money(sig.get('Position_Value', 0.0))} · 비중 {float(sig.get('Current_Weight',0.0))*100:.1f}%/{float(sig.get('Target_Weight',0.0))*100:.1f}%", font=get_font(18, True), fill="#AFC1DA", anchor="lm")
        draw.text((x0 + w - 28, y0 + 72), f"위험 {sig.get('Risk_Score', 0)}/{FULL_SELL_SCORE} · 주문 {money(sig.get('Recommended_Amount', 0.0))}", font=get_font(19, True), fill="#AFC1DA", anchor="rm")
        draw.text((x0 + w - 28, y0 + 99), f"{live_text} · {float(sig.get('Recommended_Shares',0.0)):,.4f}주", font=get_font(18, True), fill="#AFC1DA", anchor="rm")

    asset_box(qld_signal, 88, 916, "#2EE6A6", qld_state, qld_sub)
    asset_box(tqqq_signal, 992, 916, "#FF9B2F", tqqq_state, tqqq_sub)

    bg.convert("RGB").save(path)
    return path

def create_all_ai_mine_overlay_dashboards(out_dir: str, stamp: str, today_date: pd.Timestamp, summary: Dict[str, object], qld_signal: Dict[str, object], tqqq_signal: Dict[str, object], bg_dir: Optional[str] = None) -> List[str]:
    """Create one main AI mine dashboard.

    v15-ui: use the cute UI hub background (#5) as the main mine dashboard,
    instead of sending the old hand-drawn game_mine_control_room image.
    """
    bg_paths = resolve_ai_mine_backgrounds(bg_dir)
    if not bg_paths:
        print("[v15-ai-bg] AI 배경을 찾지 못했습니다. ai_mine_backgrounds_v13 폴더를 확인하세요.")
        return []

    # Prefer the #5 cute UI hub background when the standard 8-pack exists.
    selected_index = 4 if len(bg_paths) >= 5 else 0
    bg_path = bg_paths[selected_index]
    out_path = os.path.join(out_dir, f"ai_mine_main_dashboard_{stamp}.png")
    create_ai_mine_overlay_dashboard(out_path, bg_path, selected_index + 1, today_date, summary, qld_signal, tqqq_signal)
    return [out_path]


# ======================================================
# Main
# ======================================================


def export_app_json_fragment(path: str, data: pd.DataFrame, signals: List[Dict[str, object]], condition_df: pd.DataFrame, updated_at: str = "") -> str:
    """Export QLD/TQQQ strategy state for AnchorSignalApp v0.2."""
    from pathlib import Path
    names = {"QLD":"ProShares Ultra QQQ", "TQQQ":"ProShares UltraPro QQQ"}
    assets = []
    for signal in signals:
        asset = str(signal.get("Asset", ""))
        chart = []
        for idx, row in data.tail(180).iterrows():
            chart.append({
                "date": pd.Timestamp(idx).strftime("%Y-%m-%d"),
                "price": float(row.get(f"{asset}_Close", 0.0) or 0.0),
                "ma20": float(row.get(f"{asset}_MA20", 0.0) or 0.0),
                "ma50": float(row.get(f"{asset}_MA50", 0.0) or 0.0),
                "ma100": float(row.get(f"{asset}_MA100", 0.0) or 0.0),
                "rsi14": float(row.get(f"{asset}_RSI14", 0.0) or 0.0),
                "macd_hist": float(row.get(f"{asset}_MACD_HIST", 0.0) or 0.0),
                "marker": "",
            })
        daily_action = str(signal.get("Order_Action", "HOLD"))
        live_status = str(signal.get("Live_Status", ""))
        action = {"LIVE_BUY_OK":"BUY", "PULLBACK_WAIT":"WAIT", "LIVE_BUY_BLOCK":"WAIT", "LIVE_SELL_ALERT":"SELL_PARTIAL", "EMERGENCY_SELL":"SELL_ALL", "LIVE_HOLD_OK":"HOLD"}.get(live_status, daily_action)
        if chart and action in {"BUY", "BUY_MORE", "SELL_ALL", "SELL_PARTIAL"}:
            chart[-1]["marker"] = action
        sub = condition_df[condition_df["자산"] == asset] if "자산" in condition_df.columns else condition_df
        conds = [{"label":str(r.get("조건","조건")), "met":bool(r.get("충족",False)), "group":str(r.get("구분",""))} for _, r in sub.iterrows()]
        action_kr = {"BUY":"매수","BUY_MORE":"추가매수","SELL_ALL":"전량매도","SELL_PARTIAL":"일부매도","WAIT":"대기","HOLD":"보유"}.get(action, action)
        cur=float(signal.get("Current_Weight",0.0) or 0.0); tgt=float(signal.get("Target_Weight",0.0) or 0.0)
        assets.append({
            "symbol":asset, "name":names.get(asset,asset),
            "price":float(signal.get("Live_Price",signal.get("Price",0.0)) or 0.0), "change_pct":float(signal.get("Live_Diff_%",0.0) or 0.0),
            "action":action, "action_kr":action_kr, "live_status":str(signal.get("Live_Status_KR","")),
            "risk_score":int(signal.get("Risk_Score",0) or 0), "risk_max":FULL_SELL_SCORE,
            "recommended_amount":float(signal.get("Live_Order_Amount",signal.get("Recommended_Amount",0.0)) or 0.0),
            "recommended_shares":float(signal.get("Live_Order_Shares",signal.get("Recommended_Shares",0.0)) or 0.0),
            "current_weight":cur*100 if cur<=1 else cur, "target_weight":tgt*100 if tgt<=1 else tgt,
            "current_shares":float(signal.get("Shares",0.0) or 0.0), "avg_price":float(signal.get("Avg_Price",0.0) or 0.0), "position_value":float(signal.get("Position_Value",0.0) or 0.0),
            "buy_ready":bool(signal.get("Buy_Ready",False)), "sell_ready":bool(signal.get("Sell_Ready",False)),
            "reason":str(signal.get("Final_Order_Text",signal.get("Reason",""))), "conditions":conds, "chart":chart,
        })
    payload={"updated_at":updated_at or datetime.now().astimezone().strftime("%Y-%m-%d %H:%M %Z"), "assets":assets}
    out=Path(path); out.parent.mkdir(parents=True, exist_ok=True); out.write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding="utf-8")
    print(f"[app-json] {out}"); return str(out)

def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--cash", type=float, default=10000.0, help="현재 현금")
    parser.add_argument("--qld-shares", type=float, default=0.0, help="현재 QLD 보유수량")
    parser.add_argument("--qld-avg-price", type=float, default=0.0, help="현재 QLD 평균단가")
    parser.add_argument("--tqqq-shares", type=float, default=0.0, help="현재 TQQQ 보유수량")
    parser.add_argument("--tqqq-avg-price", type=float, default=0.0, help="현재 TQQQ 평균단가")
    parser.add_argument("--qld-weight", type=float, default=0.6, help="QLD 목표비중, 예: 0.6")
    parser.add_argument("--tqqq-weight", type=float, default=0.4, help="TQQQ 목표비중, 예: 0.4")
    parser.add_argument("--period", type=str, default=DOWNLOAD_PERIOD_DEFAULT, help="yfinance period 예: 5y, 15y, max")
    parser.add_argument("--output-dir", type=str, default=OUTPUT_DIR_DEFAULT)
    parser.add_argument("--ai-bg-dir", type=str, default="", help="v13 AI 광산 배경 폴더. 기본값: 스크립트 옆 ai_mine_backgrounds_v13 (기본 패키지는 8개 배경 포함)")
    parser.add_argument("--send-telegram", action="store_true", default=TELEGRAM_SEND_DEFAULT, help="텔레그램으로 요약 메시지와 PNG 현황판 전송. 파일 상단 TELEGRAM_SEND_DEFAULT=True이면 생략 가능")
    parser.add_argument("--alert-only", action="store_true", help="장중 감시용. 매수/매도 조건 또는 실시간 경고가 있을 때만 텔레그램 전송")
    parser.add_argument("--alert-use-buy-ready", action="store_true", default=True, help="alert-only에서 Buy_Ready 조건도 알림 기준에 포함")
    parser.add_argument("--alert-use-live-warning", action="store_true", default=True, help="alert-only에서 LIVE_SELL_ALERT/EMERGENCY_SELL 같은 실시간 경고도 알림 기준에 포함")
    parser.add_argument("--telegram-bot-token", type=str, default=(TELEGRAM_BOT_TOKEN_IN_FILE or os.getenv("TELEGRAM_BOT_TOKEN", "")), help="텔레그램 봇 토큰. 파일 상단 설정값 또는 환경변수를 기본 사용")
    parser.add_argument("--telegram-chat-id", type=str, default=(TELEGRAM_CHAT_ID_IN_FILE or os.getenv("TELEGRAM_CHAT_ID", "")), help="텔레그램 Chat ID. 파일 상단 설정값 또는 환경변수를 기본 사용")
    parser.add_argument("--performance-log", type=str, default="dual_anchor_performance_history.csv", help="누적 성과 스냅샷 CSV 파일명")
    parser.add_argument("--trade-log", type=str, default="dual_anchor_trade_log.csv", help="실제 체결 기록 CSV 파일명")
    parser.add_argument("--record-trade", action="store_true", help="이번 실행에서 실제 체결 기록을 trade log에 추가")
    parser.add_argument("--trade-date", type=str, default="", help="체결일 YYYY-MM-DD. 미입력 시 전략 기준일 사용")
    parser.add_argument("--qld-exec-action", type=str, default="NONE", choices=["NONE", "BUY", "SELL"], help="실제 QLD 체결 방향")
    parser.add_argument("--qld-exec-shares", type=float, default=0.0, help="실제 QLD 체결 수량")
    parser.add_argument("--qld-exec-price", type=float, default=0.0, help="실제 QLD 체결 단가")
    parser.add_argument("--tqqq-exec-action", type=str, default="NONE", choices=["NONE", "BUY", "SELL"], help="실제 TQQQ 체결 방향")
    parser.add_argument("--tqqq-exec-shares", type=float, default=0.0, help="실제 TQQQ 체결 수량")
    parser.add_argument("--tqqq-exec-price", type=float, default=0.0, help="실제 TQQQ 체결 단가")
    parser.add_argument("--app-json", type=str, default="app_data/dual_latest.json", help="AnchorSignalApp JSON 조각 출력")
    args = parser.parse_args()

    if args.qld_weight < 0 or args.tqqq_weight < 0:
        raise ValueError("목표비중은 0 이상이어야 합니다.")
    if args.qld_weight + args.tqqq_weight > 1.000001:
        raise ValueError("qld-weight + tqqq-weight 합계가 1을 초과하면 안 됩니다. 예: 0.6 + 0.4")

    print("=" * 80)
    print("DUAL QLD + TQQQ Anchor Dashboard")
    print(f"Patch: {PATCH_VERSION}")
    print("=" * 80)

    base_output_dir = args.output_dir
    if not os.path.isabs(base_output_dir):
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
        except Exception:
            script_dir = os.getcwd()
        base_output_dir = os.path.join(script_dir, base_output_dir)

    data = prepare_data(args.period)
    today = data.iloc[-1]
    today_date = data.index[-1]

    qld_price = float(today["QLD_Close"])
    tqqq_price = float(today["TQQQ_Close"])
    qqq_price = float(today["QQQ_Close"])
    qld_value = args.qld_shares * qld_price
    tqqq_value = args.tqqq_shares * tqqq_price
    total_equity = args.cash + qld_value + tqqq_value

    qld_signal, qld_cond, qld_risk = build_asset_signal(today, data, "QLD", args.cash, args.qld_shares, args.qld_avg_price, args.qld_weight, total_equity)
    tqqq_signal, tqqq_cond, tqqq_risk = build_asset_signal(today, data, "TQQQ", args.cash, args.tqqq_shares, args.tqqq_avg_price, args.tqqq_weight, total_equity)
    qld_signal, tqqq_signal = allocate_cash(qld_signal, tqqq_signal, args.cash)

    close_reference = {"QQQ": qqq_price, "QLD": qld_price, "TQQQ": tqqq_price}
    realtime_snapshots, realtime_checked_at = get_realtime_snapshot(["QQQ", "QLD", "TQQQ"], close_reference)
    for _ticker in realtime_snapshots:
        realtime_snapshots[_ticker]["Strategy_Basis_Date"] = today_date.strftime("%Y-%m-%d")
    # v14: keep daily anchor signals intact, then add live execution layer for practical timing.
    live_signals = apply_live_execution_layer([qld_signal, tqqq_signal], today, realtime_snapshots)
    qld_signal, tqqq_signal = live_signals[0], live_signals[1]
    realtime_df = build_realtime_order_df([qld_signal, tqqq_signal], realtime_snapshots)
    realtime_df = enrich_realtime_df_with_live(realtime_df, [qld_signal, tqqq_signal])

    cash_after_recommended = args.cash
    for sig in [qld_signal, tqqq_signal]:
        if sig["Order_Action"] in ["BUY", "BUY_MORE"]:
            cash_after_recommended -= float(sig["Recommended_Amount"])
        elif sig["Order_Action"] in ["SELL_ALL", "SELL_PARTIAL"]:
            cash_after_recommended += float(sig["Recommended_Amount"])

    prev_row = data.iloc[-2] if len(data) >= 2 else today
    today_pnl = args.qld_shares * (qld_price - float(prev_row["QLD_Close"])) + args.tqqq_shares * (tqqq_price - float(prev_row["TQQQ_Close"]))

    summary = {
        "Date": today_date.strftime("%Y-%m-%d"),
        "Strategy": STRATEGY_NAME,
        "Patch_Version": PATCH_VERSION,
        "Cash": args.cash,
        "Today_PnL": today_pnl,
        "Cash_After_Recommended": cash_after_recommended,
        "QQQ_Close": qqq_price,
        "QQQ_MA50": float(today["QQQ_MA50"]),
        "QQQ_MA200": float(today["QQQ_MA200"]),
        "QQQ_Above_MA50": bool(today["QQQ_Close"] > today["QQQ_MA50"]),
        "QQQ_Above_MA200": bool(today["QQQ_Close"] > today["QQQ_MA200"]),
        "QLD_Close": qld_price,
        "TQQQ_Close": tqqq_price,
        "QLD_Value": qld_value,
        "TQQQ_Value": tqqq_value,
        "Total_Equity": total_equity,
        "QLD_Target_Weight": args.qld_weight,
        "TQQQ_Target_Weight": args.tqqq_weight,
        "Cash_Target_Weight": 1 - args.qld_weight - args.tqqq_weight,
        "QLD_Order": order_korean(qld_signal),
        "TQQQ_Order": order_korean(tqqq_signal),
        "QLD_Live_Status": qld_signal.get("Live_Status_KR", ""),
        "TQQQ_Live_Status": tqqq_signal.get("Live_Status_KR", ""),
        "QLD_Final_Order": qld_signal.get("Final_Order_Action_KR", ""),
        "TQQQ_Final_Order": tqqq_signal.get("Final_Order_Action_KR", ""),
        "Realtime_Checked_At": realtime_checked_at,
        "QLD_Realtime_Price": float(realtime_snapshots.get("QLD", {}).get("Realtime_Price", qld_price)),
        "TQQQ_Realtime_Price": float(realtime_snapshots.get("TQQQ", {}).get("Realtime_Price", tqqq_price)),
        "QQQ_Realtime_Price": float(realtime_snapshots.get("QQQ", {}).get("Realtime_Price", qqq_price)),
    }

    summary_df = pd.DataFrame([summary])
    signals_df = pd.DataFrame([qld_signal, tqqq_signal])
    signals_df.insert(1, "Korean_Order", [order_korean(qld_signal), order_korean(tqqq_signal)])
    condition_df = pd.concat([qld_cond, tqqq_cond], ignore_index=True)
    risk_df = pd.concat([qld_risk, tqqq_risk], ignore_index=True)
    export_app_json_fragment(args.app_json, data, [qld_signal, tqqq_signal], condition_df, realtime_checked_at)

    recent_cols = [
        "QLD_Close", "QLD_MA20", "QLD_MA50", "QLD_MA100", "QLD_RSI14", "QLD_MACD_HIST",
        "TQQQ_Close", "TQQQ_MA20", "TQQQ_MA50", "TQQQ_MA100", "TQQQ_RSI14", "TQQQ_MACD_HIST",
        "QQQ_Close", "QQQ_MA50", "QQQ_MA200",
    ]
    recent_df = data.tail(260).copy().reset_index()
    if "Date" not in recent_df.columns:
        recent_df = recent_df.rename(columns={recent_df.columns[0]: "Date"})
    recent_df = recent_df[["Date"] + [c for c in recent_cols if c in recent_df.columns]]

    stamp = today_date.strftime("%Y%m%d")
    run_stamp = datetime.now().strftime("%H%M%S")
    out_dir = ensure_output_dir(os.path.join(base_output_dir, f"ai_mine_v13_{stamp}_{run_stamp}"))
    dashboard_path = os.path.join(out_dir, f"dual_qld_tqqq_dashboard_{stamp}.png")
    qld_dashboard_path = os.path.join(out_dir, f"qld_single_dashboard_from_dual_{stamp}.png")
    tqqq_dashboard_path = os.path.join(out_dir, f"tqqq_single_dashboard_from_dual_{stamp}.png")
    order_summary_path = os.path.join(out_dir, f"today_order_summary_{stamp}.png")
    realtime_dashboard_path = os.path.join(out_dir, f"realtime_order_check_{stamp}.png")
    performance_dashboard_path = os.path.join(out_dir, f"cumulative_performance_dashboard_{stamp}.png")
    game_mine_dashboard_path = os.path.join(out_dir, f"game_mine_control_room_{stamp}.png")
    qld_vibe_dashboard_path = os.path.join(out_dir, f"qld_safe_vibe_account_{stamp}.png")
    tqqq_vibe_dashboard_path = os.path.join(out_dir, f"tqqq_power_vibe_account_{stamp}.png")
    risk_gate_dashboard_path = os.path.join(out_dir, f"buy_sell_risk_gates_{stamp}.png")
    ai_mine_overlay_paths: List[str] = []
    excel_path = os.path.join(out_dir, f"dual_qld_tqqq_anchor_report_{stamp}.xlsx")
    csv_path = os.path.join(out_dir, f"dual_qld_tqqq_summary_{stamp}.csv")

    graph_paths = create_graphs(data, out_dir)
    create_dual_dashboard(dashboard_path, today_date, summary, qld_signal, tqqq_signal, condition_df, risk_df)
    create_asset_dashboard(qld_dashboard_path, today_date, summary, qld_signal, condition_df, risk_df)
    create_asset_dashboard(tqqq_dashboard_path, today_date, summary, tqqq_signal, condition_df, risk_df)
    create_order_summary_dashboard(order_summary_path, today_date, summary, qld_signal, tqqq_signal)
    create_realtime_order_dashboard(realtime_dashboard_path, today_date, realtime_checked_at, summary, realtime_df, qld_signal, tqqq_signal)
    trade_df = append_trade_log_if_requested(args, today_date, out_dir)
    performance_df = update_performance_history(args, out_dir, today_date, summary, realtime_df, qld_signal, tqqq_signal)
    create_performance_dashboard(performance_dashboard_path, today_date, performance_df, trade_df)
    # v15-ui: old hand-drawn game mine image is no longer generated/sent.
    create_account_vibe_dashboard(qld_vibe_dashboard_path, today_date, data, qld_signal, summary, theme='SAFE')
    create_account_vibe_dashboard(tqqq_vibe_dashboard_path, today_date, data, tqqq_signal, summary, theme='POWER')
    create_risk_gate_dashboard(risk_gate_dashboard_path, today_date, condition_df, risk_df, qld_signal, tqqq_signal)
    ai_mine_overlay_paths = create_all_ai_mine_overlay_dashboards(out_dir, stamp, today_date, summary, qld_signal, tqqq_signal, args.ai_bg_dir)
    build_excel_report(excel_path, today_date, summary_df, signals_df, condition_df, risk_df, recent_df, dashboard_path, graph_paths, realtime_df, realtime_dashboard_path, performance_df, trade_df, performance_dashboard_path)
    summary_df.to_csv(csv_path, index=False, encoding="utf-8-sig")

    print("\n==============================")
    print("통합 QLD + TQQQ 앵커 결과")
    print("==============================")
    print(f"날짜: {today_date.strftime('%Y-%m-%d')}")
    print(f"총자산: {money(total_equity)} / 현금: {money(args.cash)}")
    print(f"목표비중: QLD {args.qld_weight*100:.1f}% / TQQQ {args.tqqq_weight*100:.1f}% / 현금 {summary['Cash_Target_Weight']*100:.1f}%")
    print("\n[주문 가이드 - 종가 기준 + 실시간 실행]")
    for sig in [qld_signal, tqqq_signal]:
        print(f"- 종가: {order_korean(sig)} | 실시간: {sig.get('Live_Status_KR','')} | 최종: {sig.get('Final_Order_Action_KR','')}")
        print(f"  종가기준 금액 {money(sig['Recommended_Amount'])} / 수량 {sig['Recommended_Shares']:,.6f}")
        print(f"  실시간기준 금액 {money(sig.get('Live_Order_Amount',0.0))} / 수량 {float(sig.get('Live_Order_Shares',0.0)):,.6f}")
        print(f"  사유: {sig.get('Final_Order_Text', sig.get('Reason',''))}")
    print("\n[생성 파일]")
    print("Excel 리포트:", os.path.abspath(excel_path))
    print("통합 현황판 PNG:", os.path.abspath(dashboard_path))
    print("QLD 단독 현황판 PNG:", os.path.abspath(qld_dashboard_path))
    print("TQQQ 단독 현황판 PNG:", os.path.abspath(tqqq_dashboard_path))
    print("오늘의 주문요약 PNG:", os.path.abspath(order_summary_path))
    print("실시간 주문전 확인 PNG:", os.path.abspath(realtime_dashboard_path))
    print("누적성과 현황판 PNG:", os.path.abspath(performance_dashboard_path))
    print("게임형 광산 메인 PNG: v15-ui에서 AI 광산 메인으로 대체됨")
    print("QLD SAFE 바이브 PNG:", os.path.abspath(qld_vibe_dashboard_path))
    print("TQQQ POWER 바이브 PNG:", os.path.abspath(tqqq_vibe_dashboard_path))
    print("매수매도 리스크 게이트 PNG:", os.path.abspath(risk_gate_dashboard_path))
    for p in ai_mine_overlay_paths:
        print("AI 광산 메인 PNG:", os.path.abspath(p))
    print("누적성과 CSV:", os.path.abspath(args.performance_log if os.path.isabs(args.performance_log) else os.path.join(out_dir, args.performance_log)))
    print("실제체결 로그 CSV:", os.path.abspath(args.trade_log if os.path.isabs(args.trade_log) else os.path.join(out_dir, args.trade_log)))
    for p in graph_paths:
        print("그래프:", os.path.abspath(p))
    print("CSV 요약:", os.path.abspath(csv_path))

    def is_alert_signal(sig: Dict[str, object]) -> bool:
        """Return True when alert-only mode should send Telegram.

        This does not change the strategy calculation. It only decides whether
        the Telegram section should run during frequent monitoring jobs.
        """
        asset = str(sig.get("Asset", ""))
        order_action = str(sig.get("Order_Action", ""))
        live_status = str(sig.get("Live_Status", ""))

        # Final strategy action candidates.
        order_alert = order_action in ["BUY", "BUY_MORE", "SELL_ALL", "SELL_PARTIAL", "ADD_QLD", "ADD_TQQQ"]

        # Raw condition gates. This can alert even when you already hold the asset
        # and the dashboard says HOLD, because the buy gate itself is open.
        condition_alert = False
        if bool(getattr(args, "alert_use_buy_ready", True)):
            condition_alert = condition_alert or bool(sig.get("Buy_Ready"))
        condition_alert = condition_alert or bool(sig.get("Sell_Ready"))

        # Intraday live execution warnings/brakes.
        live_alert = False
        if bool(getattr(args, "alert_use_live_warning", True)):
            live_alert = live_status in ["LIVE_BUY_OK", "LIVE_SELL_ALERT", "EMERGENCY_SELL"]

        triggered = order_alert or condition_alert or live_alert
        if triggered:
            print(f"[alert] {asset} 알림 조건 충족: Order={order_action}, Buy_Ready={sig.get('Buy_Ready')}, Sell_Ready={sig.get('Sell_Ready')}, Live={live_status}")
        else:
            print(f"[alert] {asset} 알림 조건 미충족: Order={order_action}, Buy_Ready={sig.get('Buy_Ready')}, Sell_Ready={sig.get('Sell_Ready')}, Live={live_status}")
        return bool(triggered)

    alert_triggered = any(is_alert_signal(sig) for sig in [qld_signal, tqqq_signal])
    if args.alert_only and not alert_triggered:
        print("[alert] alert-only 모드: QLD/TQQQ 매수·매도 조건 또는 실시간 경고가 없어 텔레그램 전송을 생략합니다.")
        args.send_telegram = False
    elif args.alert_only and alert_triggered:
        print("[alert] alert-only 모드: 알림 조건이 있어 텔레그램을 전송합니다.")

    if args.send_telegram:
        if not args.telegram_bot_token or not args.telegram_chat_id:
            print("[telegram] 전송 설정은 켜져 있지만 bot token/chat id가 없습니다. 파일 상단 TELEGRAM_BOT_TOKEN_IN_FILE / TELEGRAM_CHAT_ID_IN_FILE 값을 입력하세요.")
        else:
            telegram_text = build_telegram_text(today_date, summary, qld_signal, tqqq_signal)
            send_telegram_message(args.telegram_bot_token, args.telegram_chat_id, telegram_text)
            send_telegram_photo(args.telegram_bot_token, args.telegram_chat_id, dashboard_path, caption="DUAL QLD+TQQQ 통합 현황판")
            send_telegram_photo(args.telegram_bot_token, args.telegram_chat_id, qld_dashboard_path, caption="QLD 단독 현황판")
            send_telegram_photo(args.telegram_bot_token, args.telegram_chat_id, tqqq_dashboard_path, caption="TQQQ 단독 현황판")
            send_telegram_photo(args.telegram_bot_token, args.telegram_chat_id, order_summary_path, caption="오늘의 주문요약")
            send_telegram_photo(args.telegram_bot_token, args.telegram_chat_id, realtime_dashboard_path, caption="실시간 주문 전 확인")
            send_telegram_photo(args.telegram_bot_token, args.telegram_chat_id, performance_dashboard_path, caption="누적 성과 현황판")
            send_telegram_photo(args.telegram_bot_token, args.telegram_chat_id, qld_vibe_dashboard_path, caption="QLD SAFE 바이브 계좌")
            send_telegram_photo(args.telegram_bot_token, args.telegram_chat_id, tqqq_vibe_dashboard_path, caption="TQQQ POWER 바이브 계좌")
            send_telegram_photo(args.telegram_bot_token, args.telegram_chat_id, risk_gate_dashboard_path, caption="매수매도 리스크 게이트")
            for i, p in enumerate(ai_mine_overlay_paths, start=1):
                send_telegram_photo(args.telegram_bot_token, args.telegram_chat_id, p, caption="AI 광산 메인 대시보드")

    print("\n완료")


if __name__ == "__main__":
    main()
