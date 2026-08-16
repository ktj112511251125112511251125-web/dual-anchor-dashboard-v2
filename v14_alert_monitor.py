# ======================================================
# V14 99/1 Alert Monitor
# 장중 10분 감시용: 익절 전량매도 / 다운시프트만 텔레그램 발송
# BUY 후보는 알림 보내지 않음
# ======================================================
import os
import json
import argparse
import subprocess
import sys
from datetime import datetime


def ensure_package(package, import_name=None):
    name = import_name or package
    try:
        __import__(name)
    except ModuleNotFoundError:
        print(f"[install] {package} 설치 중...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

for _pkg, _imp in [
    ("requests", "requests"),
    ("pandas", "pandas"),
    ("numpy", "numpy"),
    ("yfinance", "yfinance"),
    ("pytz", "pytz"),
]:
    ensure_package(_pkg, _imp)

import requests
import pandas as pd
import yfinance as yf
import pytz

TARGETS = {
    "TQQQ": 0.15,
    "QLD": 0.10,
    "QQQ": 0.05,
    "TQQQ+QLD": 0.12,
    "TQQQ+QQQ": 0.10,
    "QLD+QQQ": 0.07,
    "QQQ+QLD+TQQQ": 0.07,
}
DOWNSHIFT_MIX = {"QLD": 0.99, "QQQ": 0.01}


def env_float(name, default=0.0):
    try:
        v = os.environ.get(name, "")
        return float(default if str(v).strip() == "" else v)
    except Exception:
        return float(default)


def env_bool(name, default=False):
    v = str(os.environ.get(name, "")).strip().lower()
    if not v:
        return default
    return v in ["1", "true", "yes", "y", "on"]


def parse_args():
    p = argparse.ArgumentParser(description="V14 99/1 intraday alert monitor")
    p.add_argument("--cash", type=float, default=env_float("V14_CASH", 0))
    p.add_argument("--qqq-shares", type=float, default=env_float("V14_QQQ_SHARES", 0))
    p.add_argument("--qld-shares", type=float, default=env_float("V14_QLD_SHARES", 0))
    p.add_argument("--tqqq-shares", type=float, default=env_float("V14_TQQQ_SHARES", 0))
    p.add_argument("--qqq-avg-price", type=float, default=env_float("V14_QQQ_AVG_PRICE", 0))
    p.add_argument("--qld-avg-price", type=float, default=env_float("V14_QLD_AVG_PRICE", 0))
    p.add_argument("--tqqq-avg-price", type=float, default=env_float("V14_TQQQ_AVG_PRICE", 0))
    p.add_argument("--period", default=os.environ.get("V14_PERIOD", "15y"))
    p.add_argument("--send-telegram", action="store_true", default=env_bool("V14_SEND_TELEGRAM", True))
    p.add_argument("--telegram-bot-token", default=os.environ.get("TELEGRAM_BOT_TOKEN", ""))
    p.add_argument("--telegram-chat-id", default=os.environ.get("TELEGRAM_CHAT_ID", ""))
    p.add_argument("--state-file", default=os.environ.get("V14_ALERT_STATE_FILE", ".v14_alert_state/v14_alert_state.json"))
    p.add_argument("--no-dedup", action="store_true", help="중복방지 끄기")
    return p.parse_args()


def send_telegram(text, token, chat_id):
    if not token or not chat_id:
        print("텔레그램 토큰/chat_id 없음. 메시지 출력만 합니다.")
        print(text)
        return False
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    r = requests.post(url, data={"chat_id": chat_id, "text": text}, timeout=20)
    print("telegram status:", r.status_code, r.text[:200])
    return r.ok


def get_prices(period="15y"):
    data = yf.download(["QQQ", "QLD", "TQQQ"], period=period, auto_adjust=True, progress=False)
    if isinstance(data.columns, pd.MultiIndex):
        close = data["Close"].dropna(how="all")
    else:
        close = data[["Close"]].rename(columns={"Close": "QQQ"}).dropna()
    qqq_hist = close["QQQ"].dropna()
    qqq_ma200 = float(qqq_hist.rolling(200).mean().iloc[-1])
    last_close = {t: float(close[t].dropna().iloc[-1]) for t in ["QQQ", "QLD", "TQQQ"] if t in close.columns}
    live = {}
    for t in ["QQQ", "QLD", "TQQQ"]:
        px = None
        try:
            info = yf.Ticker(t).fast_info
            px = getattr(info, "last_price", None) or info.get("last_price")
        except Exception:
            px = None
        live[t] = float(px) if px and px > 0 else float(last_close.get(t, 0))
    return last_close, live, qqq_ma200


def holding_combo(holdings):
    names = [t for t in ["QQQ", "QLD", "TQQQ"] if holdings.get(t, 0) > 0]
    if not names:
        return "NONE"
    # TARGETS 키 순서에 맞게 혼합명 구성
    if set(names) == {"QQQ", "QLD", "TQQQ"}:
        return "QQQ+QLD+TQQQ"
    return "+".join(names)


def load_state(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_state(path, state):
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


def already_sent_today(state, alert_type):
    ny = pytz.timezone("America/New_York")
    today = datetime.now(ny).strftime("%Y-%m-%d")
    return state.get("date") == today and alert_type in state.get("sent", [])


def mark_sent_today(state, alert_type):
    ny = pytz.timezone("America/New_York")
    today = datetime.now(ny).strftime("%Y-%m-%d")
    if state.get("date") != today:
        state = {"date": today, "sent": []}
    if alert_type not in state["sent"]:
        state["sent"].append(alert_type)
    return state


def main():
    args = parse_args()
    holdings = {"QQQ": args.qqq_shares, "QLD": args.qld_shares, "TQQQ": args.tqqq_shares}
    avg = {"QQQ": args.qqq_avg_price, "QLD": args.qld_avg_price, "TQQQ": args.tqqq_avg_price}
    print("V14 ALERT MONITOR 시작")
    print("보유:", holdings, "평단:", avg)

    last_close, live, qqq_ma200 = get_prices(args.period)
    position_value = sum(holdings[t] * live[t] for t in ["QQQ", "QLD", "TQQQ"])
    position_cost = sum(holdings[t] * avg[t] for t in ["QQQ", "QLD", "TQQQ"] if holdings[t] > 0 and avg[t] > 0)
    position_return = (position_value / position_cost - 1.0) if position_cost > 0 else 0.0
    combo = holding_combo(holdings)
    target = TARGETS.get(combo)

    alert_type = None
    text = ""

    # 1순위: 익절 전량매도
    if target is not None and position_cost > 0 and position_return >= target:
        alert_type = "PROFIT_SELL_ALERT"
        lines = [
            "[v14 긴급 익절 전량매도]",
            "",
            f"보유조합: {combo}",
            f"현재 포지션 수익률: {position_return*100:.2f}%",
            f"익절 목표수익률: {target*100:.2f}%",
            f"평가금액: ${position_value:,.2f}",
            f"원금: ${position_cost:,.2f}",
            "",
            "실행 후보:",
        ]
        for t in ["QQQ", "QLD", "TQQQ"]:
            sh = holdings[t]
            if sh > 0:
                amt = sh * live[t]
                lines.append(f"SELL {t}: {sh}주 / 기준가 ${live[t]:,.2f} / 예상 ${amt:,.2f}")
        lines += ["", "주의:", "오늘은 재매수하지 않습니다.", "실제 매도 후 v14_account_config.env의 V14_LAST_PROFIT_SELL_DATE에 오늘 날짜를 입력하세요."]
        text = "\n".join(lines)

    # 2순위: 다운시프트
    elif holdings["TQQQ"] > 0 and live["QQQ"] < qqq_ma200:
        alert_type = "DOWNSHIFT_ALERT"
        sell_amount = holdings["TQQQ"] * live["TQQQ"]
        qld_amt = sell_amount * DOWNSHIFT_MIX["QLD"]
        qqq_amt = sell_amount * DOWNSHIFT_MIX["QQQ"]
        qld_sh = qld_amt / live["QLD"] if live["QLD"] > 0 else 0
        qqq_sh = qqq_amt / live["QQQ"] if live["QQQ"] > 0 else 0
        text = "\n".join([
            "[v14 긴급 다운시프트]",
            "",
            "조건 발생: TQQQ 보유 중 QQQ 현재가 < QQQ MA200",
            f"QQQ 현재가: ${live['QQQ']:,.2f}",
            f"QQQ MA200: ${qqq_ma200:,.2f}",
            "",
            "실행 후보:",
            f"SELL TQQQ: {holdings['TQQQ']}주 / 기준가 ${live['TQQQ']:,.2f} / 예상 ${sell_amount:,.2f}",
            f"BUY QLD 99%: ${qld_amt:,.2f} / 예상 {qld_sh:.6f}주",
            f"BUY QQQ 1%: ${qqq_amt:,.2f} / 예상 {qqq_sh:.6f}주",
        ])

    if not alert_type:
        print("긴급조건 없음: 무음 종료")
        print(f"QQQ live={live['QQQ']:.2f}, MA200={qqq_ma200:.2f}, combo={combo}, return={position_return*100:.2f}%")
        return

    state = load_state(args.state_file)
    if not args.no_dedup and already_sent_today(state, alert_type):
        print(f"오늘 이미 {alert_type} 발송함: 중복 방지로 무음 종료")
        return

    print(text)
    if args.send_telegram:
        send_telegram(text, args.telegram_bot_token, args.telegram_chat_id)
    state = mark_sent_today(state, alert_type)
    save_state(args.state_file, state)


if __name__ == "__main__":
    main()
