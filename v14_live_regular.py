# 실행 전 필요 패키지:
#   pip install yfinance pillow requests matplotlib pandas numpy
# 예시:
#   python QQQ_QLD_TQQQ_v14_99_1_downshift_live_judgement_cmd.py --cash 10000 --qld-shares 5 --qld-avg-price 90


# ======================================================
# QQQ / QLD / TQQQ 최종 전략 - Windows CMD / 터미널 실행용
# v14-99/1-live GitHub Actions 지원: 종가 판단 + 실시간 현재가 판단 + MA200 TQQQ 다운시프트
# 그래프 제외 모든 카드 하단 노란 멘트바 포함
# ======================================================


import os
import re
import argparse
import sys
import subprocess
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
    ("numpy", "numpy"),
    ("pandas", "pandas"),
    ("yfinance", "yfinance"),
    ("matplotlib", "matplotlib"),
    ("pillow", "PIL"),
    ("pandas_market_calendars", "pandas_market_calendars"),
    ("pytz", "pytz"),
]:
    ensure_package(_pkg, _imp)

import requests
import numpy as np
import pandas as pd
import yfinance as yf
import pytz
import pandas_market_calendars as mcal
import matplotlib.pyplot as plt
from PIL import Image, ImageDraw, ImageFont, ImageFilter

IN_COLAB = False
files = None

# ======================================================
# 1. 설정값
# ======================================================

TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")
SEND_TELEGRAM = False

AUTO_IMAGE_OR_MANUAL = False

CASH_MANUAL = 10000
HOLDINGS_MANUAL = {"QQQ": 0.0, "QLD": 0.0, "TQQQ": 0.0}
AVG_PRICE_MANUAL = {"QQQ": 0.0, "QLD": 0.0, "TQQQ": 0.0}

MANUAL_SOLD_TODAY = False
COOLDOWN_DAYS_LEFT = 0
SELL_COOLDOWN_DAYS_AFTER_PROFIT = 7

STRATEGY_NAME = "DD쿨차등_C_17_19_21_24_C20_DS99_1"
DOWNSHIFT_TRIGGER = "TQQQ 보유 중 QQQ < MA200"
DOWNSHIFT_MIX = {"QLD": 0.99, "QQQ": 0.01}
DOWNSHIFT_TQQQ_TO_ZERO = True

DD_STEP_DIV = {"DD_0_5": 17, "DD_5_10": 19, "DD_10_15": 21, "DD_15_MORE": 24}
COOLDOWN_TQQQ_DIV = 20
QQQ_DIV = 60
QLD_DIV = 60
MIX_DIV = 60

TARGETS = {
    "TQQQ": 0.15,
    "QLD": 0.10,
    "QQQ": 0.05,
    "TQQQ_QLD": 0.12,
    "TQQQ_QQQ": 0.10,
    "QLD_QQQ": 0.07,
    "ALL": 0.07,
}

CASH = CASH_MANUAL
HOLDINGS = HOLDINGS_MANUAL.copy()
AVG_PRICE = AVG_PRICE_MANUAL.copy()

OUTPUT_DIR = os.path.abspath(os.environ.get("V14_OUTPUT_DIR", "qqq_qld_tqqq_v14_99_1_live_results"))

def env_float(name, default):
    try:
        value = os.environ.get(name, None)
        if value is None or str(value).strip() == "":
            return float(default)
        return float(value)
    except Exception:
        return float(default)

def env_int(name, default):
    try:
        value = os.environ.get(name, None)
        if value is None or str(value).strip() == "":
            return int(default)
        return int(float(value))
    except Exception:
        return int(default)

def env_bool(name, default=False):
    value = os.environ.get(name, None)
    if value is None or str(value).strip() == "":
        return bool(default)
    return str(value).strip().lower() in ["1", "true", "yes", "y", "on"]


def calculate_cooldown_days_from_last_sell(last_sell_date_str, total_days=7):
    """마지막 익절 전량매도일 기준 NYSE 거래일 쿨다운 잔여일 계산.
    - 매도 당일: 7
    - 다음 NYSE 거래일: 7
    - 그 다음 NYSE 거래일: 6 ...
    """
    s = str(last_sell_date_str or "").strip()
    if not s:
        return 0, ""
    try:
        sell_date = pd.Timestamp(s).date()
    except Exception:
        print(f"[WARN] V14_LAST_PROFIT_SELL_DATE 형식 오류: {s}. YYYY-MM-DD 형식으로 입력하세요.")
        return 0, ""
    ny = pytz.timezone("America/New_York")
    today_ny = datetime.now(ny).date()
    if today_ny <= sell_date:
        return int(total_days), f"마지막 익절매도일 {sell_date}, 매도 당일 또는 이전 날짜로 인식"
    try:
        nyse = mcal.get_calendar("NYSE")
        sched = nyse.schedule(start_date=str(sell_date), end_date=str(today_ny))
        days = [d.date() for d in sched.index]
        # sell_date 다음 거래일부터 쿨다운 7거래일 시작.
        # 첫 거래일에는 7일 남음으로 표시하고, 그 다음 거래일부터 6,5...로 감소.
        elapsed_after_first = max(0, len([d for d in days if d > sell_date]) - 1)
        remaining = max(0, int(total_days) - elapsed_after_first)
        return remaining, f"마지막 익절매도일 {sell_date} 기준 NYSE 거래일 자동계산"
    except Exception as e:
        print(f"[WARN] NYSE 거래일 계산 실패: {e}. 쿨다운 0으로 처리합니다.")
        return 0, ""

def parse_args():
    parser = argparse.ArgumentParser(description="QQQ/QLD/TQQQ v14 live judgement - CMD executable script")
    parser.add_argument("--cash", type=float, default=env_float("V14_CASH", CASH_MANUAL), help="현금")
    parser.add_argument("--qqq-shares", type=float, default=env_float("V14_QQQ_SHARES", HOLDINGS_MANUAL["QQQ"]), help="QQQ 보유수량")
    parser.add_argument("--qld-shares", type=float, default=env_float("V14_QLD_SHARES", HOLDINGS_MANUAL["QLD"]), help="QLD 보유수량")
    parser.add_argument("--tqqq-shares", type=float, default=env_float("V14_TQQQ_SHARES", HOLDINGS_MANUAL["TQQQ"]), help="TQQQ 보유수량")
    parser.add_argument("--qqq-avg-price", type=float, default=env_float("V14_QQQ_AVG_PRICE", AVG_PRICE_MANUAL["QQQ"]), help="QQQ 평균단가")
    parser.add_argument("--qld-avg-price", type=float, default=env_float("V14_QLD_AVG_PRICE", AVG_PRICE_MANUAL["QLD"]), help="QLD 평균단가")
    parser.add_argument("--tqqq-avg-price", type=float, default=env_float("V14_TQQQ_AVG_PRICE", AVG_PRICE_MANUAL["TQQQ"]), help="TQQQ 평균단가")
    parser.add_argument("--sold-today", action="store_true", default=env_bool("V14_SOLD_TODAY", False), help="오늘 이미 매도한 것으로 처리")
    parser.add_argument("--cooldown-days", type=int, default=env_int("V14_COOLDOWN_DAYS", COOLDOWN_DAYS_LEFT), help="남은 쿨다운 거래일. 보통은 직접 쓰지 않고 last-profit-sell-date를 사용")
    parser.add_argument("--last-profit-sell-date", default=os.environ.get("V14_LAST_PROFIT_SELL_DATE", ""), help="마지막 익절 전량매도일 YYYY-MM-DD. 이 값이 있으면 쿨다운 남은 거래일을 자동 계산")
    parser.add_argument("--output-dir", default=os.environ.get("V14_OUTPUT_DIR", OUTPUT_DIR), help="결과 파일 저장 폴더")
    parser.add_argument("--period", default=os.environ.get("V14_PERIOD", "2y"), help="yfinance 데이터 기간. 예: 2y, 15y")
    parser.add_argument("--send-telegram", action="store_true", default=env_bool("V14_SEND_TELEGRAM", False), help="텔레그램 발송")
    parser.add_argument("--telegram-bot-token", default=os.environ.get("TELEGRAM_BOT_TOKEN", ""), help="텔레그램 봇 토큰")
    parser.add_argument("--telegram-chat-id", default=os.environ.get("TELEGRAM_CHAT_ID", ""), help="텔레그램 chat_id")
    return parser.parse_args()

_args = parse_args()
CASH_MANUAL = float(_args.cash)
HOLDINGS_MANUAL = {"QQQ": float(_args.qqq_shares), "QLD": float(_args.qld_shares), "TQQQ": float(_args.tqqq_shares)}
AVG_PRICE_MANUAL = {"QQQ": float(_args.qqq_avg_price), "QLD": float(_args.qld_avg_price), "TQQQ": float(_args.tqqq_avg_price)}
MANUAL_SOLD_TODAY = bool(_args.sold_today)
COOLDOWN_DAYS_LEFT = int(_args.cooldown_days)
COOLDOWN_AUTO_NOTE = "수동/기본값 기준"
if str(getattr(_args, "last_profit_sell_date", "")).strip():
    COOLDOWN_DAYS_LEFT, COOLDOWN_AUTO_NOTE = calculate_cooldown_days_from_last_sell(_args.last_profit_sell_date, SELL_COOLDOWN_DAYS_AFTER_PROFIT)
OUTPUT_DIR = os.path.abspath(_args.output_dir)
V14_PERIOD = str(_args.period)
os.makedirs(OUTPUT_DIR, exist_ok=True)
SEND_TELEGRAM = bool(_args.send_telegram)
TELEGRAM_BOT_TOKEN = _args.telegram_bot_token
TELEGRAM_CHAT_ID = _args.telegram_chat_id
CASH = CASH_MANUAL
HOLDINGS = HOLDINGS_MANUAL.copy()
AVG_PRICE = AVG_PRICE_MANUAL.copy()

# ======================================================
# 2. 디자인 유틸
# ======================================================

def _find_font(bold=True):
    candidates = []
    if os.name == "nt":
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

def get_font(size, bold=True):
    path = _find_font(bold)
    if path:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            pass
    return ImageFont.load_default()

CARD = {
    "bg_top": "#06142E", "bg_bottom": "#05313A", "navy": "#04152D", "navy2": "#0A2145",
    "white": "#FFFFFF", "card_bg": "#F2F6FF", "text": "#06142E", "subtext": "#20324D", "muted": "#BFD3F7",
    "buy": "#008A4E", "buy_dark": "#005F36", "buy_soft": "#DFFBEF",
    "sell": "#D72638", "sell_dark": "#9E1422", "sell_soft": "#FFE3E7",
    "wait": "#E88C00", "wait_dark": "#A76300", "wait_soft": "#FFF1D6",
    "blue": "#1D4ED8", "blue_soft": "#E5EDFF", "purple": "#6D28D9", "pink": "#DB2777",
    "cyan": "#0098B8", "yellow": "#FFD84D",
}

def money(x):
    try: return f"${float(x):,.2f}"
    except Exception: return str(x)

def percent(x):
    try: return f"{float(x):.2f}%"
    except Exception: return str(x)

def num2(x):
    try: return f"{float(x):.2f}"
    except Exception: return str(x)

def hex_to_rgb(h):
    return tuple(int(h[i:i+2], 16) for i in (1,3,5))

def draw_gradient(img, top, bottom):
    w,h = img.size
    top = hex_to_rgb(top); bottom = hex_to_rgb(bottom)
    px = img.load()
    for y in range(h):
        rto = y/h
        r = int(top[0]*(1-rto)+bottom[0]*rto)
        g = int(top[1]*(1-rto)+bottom[1]*rto)
        b = int(top[2]*(1-rto)+bottom[2]*rto)
        for x in range(w):
            px[x,y] = (r,g,b)

def rounded(draw, xy, radius=30, fill="#FFFFFF", outline=None, width=1):
    draw.rounded_rectangle(xy, radius=radius, fill=fill, outline=outline, width=width)

def card(base, xy, radius=40, fill="#FFFFFF"):
    x1,y1,x2,y2 = xy
    shadow = Image.new("RGBA", base.size, (0,0,0,0))
    sd = ImageDraw.Draw(shadow)
    sd.rounded_rectangle((x1+8,y1+10,x2+8,y2+10), radius=radius, fill=(0,0,0,55))
    shadow = shadow.filter(ImageFilter.GaussianBlur(12))
    base.alpha_composite(shadow)
    ImageDraw.Draw(base).rounded_rectangle(xy, radius=radius, fill=fill)

def text_lines(draw, text, font_obj, max_width, max_lines=None):
    text = str(text)
    lines = []
    cur = ""
    for ch in text:
        test = cur + ch
        bb = draw.textbbox((0,0), test, font=font_obj)
        if bb[2]-bb[0] <= max_width:
            cur = test
        else:
            if cur: lines.append(cur)
            cur = ch
    if cur: lines.append(cur)
    if max_lines and len(lines) > max_lines:
        lines = lines[:max_lines]
        last = lines[-1]
        while len(last) > 1:
            test = last + "..."
            bb = draw.textbbox((0,0), test, font=font_obj)
            if bb[2]-bb[0] <= max_width:
                lines[-1] = test
                break
            last = last[:-1]
    return lines

def draw_text_box(draw, x, y, text, font_obj, fill, max_width, line_gap=10, max_lines=None):
    lines = text_lines(draw, text, font_obj, max_width, max_lines)
    step = font_obj.size + line_gap
    for i,line in enumerate(lines):
        draw.text((x, y+i*step), line, font=font_obj, fill=fill)
    return y + len(lines)*step

def chip(draw, x, y, text, fill, text_color, h=48, font_size=22):
    font = get_font(font_size)
    bb = draw.textbbox((0,0), text, font=font)
    w = bb[2]-bb[0] + 42
    rounded(draw, (x,y,x+w,y+h), radius=h//2, fill=fill)
    draw.text((x+w//2,y+h//2), text, font=font, fill=text_color, anchor="mm")
    return x+w+12

def kv(draw, x, y, label, value, value_color=None, size=26, value_x=960):
    if value_color is None: value_color = CARD["text"]
    draw.text((x,y), str(label), font=get_font(size), fill=CARD["subtext"])
    draw.text((value_x,y), str(value), font=get_font(size), fill=value_color, anchor="ra")

def draw_yellow_footer(draw, W, y, line1="상승장=TQQQ · 조정장=QLD", line2="TQQQ 보유+MA200 이탈=QLD99/QQQ1 다운시프트"):
    rounded(draw, (55,y,W-55,y+112), radius=48, fill=CARD["yellow"])
    draw.text((W//2,y+40), line1, font=get_font(29), fill=CARD["text"], anchor="mm")
    draw.text((W//2,y+78), line2, font=get_font(29), fill=CARD["text"], anchor="mm")

# ======================================================
# 3. 텔레그램
# ======================================================

def send_telegram_message(text):
    if not SEND_TELEGRAM: return False
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("텔레그램 토큰/chat_id가 비어 있어 메시지 발송을 건너뜁니다.")
        return False
    try:
        url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
        r = requests.post(url, data={"chat_id": TELEGRAM_CHAT_ID, "text": text}, timeout=10)
        js = r.json()
        if r.status_code == 200 and js.get("ok"):
            print("텔레그램 메시지 발송 완료")
            return True
        print("텔레그램 메시지 발송 실패", js)
    except Exception as e:
        print("텔레그램 메시지 오류:", e)
    return False

def send_telegram_photo(image_path, caption=""):
    if not SEND_TELEGRAM: return False
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("텔레그램 토큰/chat_id가 비어 있어 사진 발송을 건너뜁니다.")
        return False
    try:
        url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendPhoto"
        with open(image_path, "rb") as f:
            r = requests.post(url, data={"chat_id": TELEGRAM_CHAT_ID, "caption": caption}, files={"photo": f}, timeout=25)
        js = r.json()
        if r.status_code == 200 and js.get("ok"):
            print("텔레그램 사진 발송 완료:", image_path)
            return True
        print("텔레그램 사진 발송 실패", js)
    except Exception as e:
        print("텔레그램 사진 오류:", e)
    return False

# ======================================================
# 4. 보유현황 사진 입력/OCR
# ======================================================

def maybe_upload_holdings_image():
    if not IN_COLAB:
        print("CMD/터미널 실행에서는 사진 업로드를 생략합니다")
        return None
    print("보유현황 캡쳐 사진이 있으면 업로드하세요. 없으면 취소하세요.")
    try:
        uploaded = files.upload()
        if not uploaded:
            print("사진 업로드 없음 -> 수동 입력값 사용")
            return None
        image_path = list(uploaded.keys())[0]
        print("업로드된 사진:", image_path)
        return image_path
    except Exception as e:
        print("사진 업로드 취소/실패 -> 수동 입력값 사용", e)
        return None

def normalize_ocr_number(text):
    s = str(text).replace(",", "").replace("$", "").replace(" ", "")
    s = re.sub(r"[^0-9.\-]", "", s)
    if s in ["", ".", "-", "-."]:
        return None
    try: return float(s)
    except Exception: return None

def parse_holdings_from_ocr_text(ocr_lines):
    result = {"QQQ":{"shares":None,"avg_price":None}, "QLD":{"shares":None,"avg_price":None}, "TQQQ":{"shares":None,"avg_price":None}}
    print("="*80); print("OCR 전체 텍스트"); print("="*80); print("\n".join(ocr_lines))
    for ticker in ["TQQQ", "QLD", "QQQ"]:
        for i,line in enumerate(ocr_lines):
            if ticker in line.upper():
                near = ocr_lines[max(0,i-2):min(len(ocr_lines),i+5)]
                nums = []
                for x in near:
                    for token in re.split(r"\s+", x):
                        n = normalize_ocr_number(token)
                        if n is not None: nums.append(n)
                possible = [n for n in nums if 0 < n < 10000]
                if len(possible) >= 1: result[ticker]["shares"] = possible[0]
                if len(possible) >= 2: result[ticker]["avg_price"] = possible[1]
                break
    return result

def has_valid_ocr_result(parsed):
    for ticker in ["QQQ","QLD","TQQQ"]:
        s = parsed[ticker]["shares"]; a = parsed[ticker]["avg_price"]
        if s is not None and a is not None and s > 0 and a > 0:
            return True
    return False

def apply_image_or_manual_input():
    global CASH, HOLDINGS, AVG_PRICE
    CASH = CASH_MANUAL
    HOLDINGS = HOLDINGS_MANUAL.copy()
    AVG_PRICE = AVG_PRICE_MANUAL.copy()
    if not AUTO_IMAGE_OR_MANUAL:
        print("사진 입력 비활성화 -> 수동 입력값 사용")
        return
    image_path = maybe_upload_holdings_image()
    if image_path is None:
        return
    try:
        import easyocr
        reader = easyocr.Reader(["ko", "en"], gpu=False)
        ocr_lines = reader.readtext(image_path, detail=0)
        parsed = parse_holdings_from_ocr_text(ocr_lines)
        print("자동 추출 후보:", parsed)
        if not has_valid_ocr_result(parsed):
            print("OCR 유효 결과 없음 -> 수동 입력값 사용")
            return
        for ticker in ["QQQ","QLD","TQQQ"]:
            if parsed[ticker]["shares"] is not None and parsed[ticker]["shares"] > 0:
                HOLDINGS[ticker] = parsed[ticker]["shares"]
            if parsed[ticker]["avg_price"] is not None and parsed[ticker]["avg_price"] > 0:
                AVG_PRICE[ticker] = parsed[ticker]["avg_price"]
        print("사진 OCR 결과 반영")
        print("CASH=", CASH, "HOLDINGS=", HOLDINGS, "AVG_PRICE=", AVG_PRICE)
    except Exception as e:
        print("사진 분석 실패 -> 수동 입력값 사용", e)

apply_image_or_manual_input()

# ======================================================
# 5. 데이터 다운로드 / 지표 계산
# ======================================================

def download_ohlcv(ticker, period="2y"):
    df = yf.download(ticker, period=period, interval="1d", auto_adjust=True, progress=False, threads=False)
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)
    if df.empty:
        raise RuntimeError(f"{ticker} 데이터 다운로드 실패")
    return df[["Open", "High", "Low", "Close", "Volume"]].dropna()

print("데이터 다운로드 중...")
qqq = download_ohlcv("QQQ", period=V14_PERIOD)
qld = download_ohlcv("QLD", period=V14_PERIOD)
tqqq = download_ohlcv("TQQQ", period=V14_PERIOD)

data = pd.DataFrame(index=qqq.index)
data["QQQ_Open"] = qqq["Open"]
data["QQQ_High"] = qqq["High"]
data["QQQ_Low"] = qqq["Low"]
data["QQQ_Close"] = qqq["Close"]
data["QQQ_Volume"] = qqq["Volume"]
data["QLD_Close"] = qld["Close"]
data["TQQQ_Close"] = tqqq["Close"]
data = data.dropna()
print("데이터 다운로드 완료")

# 이동평균/추세/낙폭
data["QQQ_MA20"] = data["QQQ_Close"].rolling(20).mean()
data["QQQ_MA50"] = data["QQQ_Close"].rolling(50).mean()
data["QQQ_MA100"] = data["QQQ_Close"].rolling(100).mean()
data["QQQ_MA200"] = data["QQQ_Close"].rolling(200).mean()
data["Above_MA20"] = data["QQQ_Close"] > data["QQQ_MA20"]
data["Above_MA200"] = data["QQQ_Close"] > data["QQQ_MA200"]
data["MA50_Slope_Up"] = data["QQQ_MA50"] > data["QQQ_MA50"].shift(5)
data["Rolling_1Y_High"] = data["QQQ_Close"].rolling(252).max()
data["Drawdown_From_1Y_High"] = data["QQQ_Close"] / data["Rolling_1Y_High"] - 1

# 볼린저밴드
data["BB_MID"] = data["QQQ_Close"].rolling(20).mean()
data["BB_STD"] = data["QQQ_Close"].rolling(20).std()
data["BB_UPPER"] = data["BB_MID"] + 2 * data["BB_STD"]
data["BB_LOWER"] = data["BB_MID"] - 2 * data["BB_STD"]
data["BB_WIDTH"] = (data["BB_UPPER"] - data["BB_LOWER"]) / data["BB_MID"]

# RSI
delta = data["QQQ_Close"].diff()
gain = delta.clip(lower=0)
loss = -delta.clip(upper=0)
avg_gain = gain.rolling(14).mean()
avg_loss = loss.rolling(14).mean()
rs = avg_gain / avg_loss.replace(0, np.nan)
data["RSI14"] = 100 - (100 / (1 + rs))

# Stochastic
low14 = data["QQQ_Low"].rolling(14).min()
high14 = data["QQQ_High"].rolling(14).max()
stoch_den = (high14 - low14).replace(0, np.nan)
data["STOCH_K"] = 100 * (data["QQQ_Close"] - low14) / stoch_den
data["STOCH_D"] = data["STOCH_K"].rolling(3).mean()

# CCI
tp = (data["QQQ_High"] + data["QQQ_Low"] + data["QQQ_Close"]) / 3
tp_ma = tp.rolling(20).mean()
tp_md = tp.rolling(20).apply(lambda x: np.mean(np.abs(x - np.mean(x))), raw=True)
data["CCI20"] = (tp - tp_ma) / (0.015 * tp_md.replace(0, np.nan))

# ADX
high = data["QQQ_High"]; low = data["QQQ_Low"]; close = data["QQQ_Close"]
up_move = high.diff(); down_move = -low.diff()
plus_dm_arr = np.where((up_move > down_move) & (up_move > 0), up_move, 0.0)
minus_dm_arr = np.where((down_move > up_move) & (down_move > 0), down_move, 0.0)
plus_dm = pd.Series(plus_dm_arr, index=data.index)
minus_dm = pd.Series(minus_dm_arr, index=data.index)
tr = pd.concat([(high-low), (high-close.shift(1)).abs(), (low-close.shift(1)).abs()], axis=1).max(axis=1)
atr14 = tr.rolling(14).mean()
plus_di14 = 100 * plus_dm.rolling(14).mean() / atr14.replace(0, np.nan)
minus_di14 = 100 * minus_dm.rolling(14).mean() / atr14.replace(0, np.nan)
dx = 100 * (plus_di14 - minus_di14).abs() / (plus_di14 + minus_di14).replace(0, np.nan)
data["ADX14"] = dx.rolling(14).mean()
data["PLUS_DI14"] = plus_di14
data["MINUS_DI14"] = minus_di14

# OBV
obv_direction = np.sign(data["QQQ_Close"].diff()).fillna(0)
data["OBV"] = (obv_direction * data["QQQ_Volume"]).cumsum()
data["OBV_MA20"] = data["OBV"].rolling(20).mean()
data["OBV_Trend_Up"] = data["OBV"] > data["OBV_MA20"]

data = data.dropna()
today = data.iloc[-1]
today_date = data.index[-1]
prices = {"QQQ": float(today["QQQ_Close"]), "QLD": float(today["QLD_Close"]), "TQQQ": float(today["TQQQ_Close"])}

# ======================================================
# 5-1. v14-live: 실시간 현재가 스냅샷
# ======================================================
# 기존 백테스트/전략 판단용 가격은 일봉 종가(prices_close)로 고정합니다.
# 실시간 판단용 가격은 별도 prices_live에 저장합니다.
prices_close = prices.copy()


def safe_float(value, default=None):
    try:
        if value is None:
            return default
        v = float(value)
        if np.isnan(v) or np.isinf(v):
            return default
        return v
    except Exception:
        return default


def get_realtime_snapshot(tickers, fallback_prices):
    checked_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = {}
    for ticker in tickers:
        fallback = safe_float(fallback_prices.get(ticker), 0.0) or 0.0
        live_price = None
        source = "unavailable"
        previous_close = None
        try:
            info = yf.Ticker(ticker).fast_info
            def pick(name):
                try:
                    if isinstance(info, dict):
                        return info.get(name)
                    return getattr(info, name)
                except Exception:
                    return None
            live_price = safe_float(pick("last_price"))
            if live_price is None:
                live_price = safe_float(pick("lastPrice"))
            previous_close = safe_float(pick("previous_close"))
            if previous_close is None:
                previous_close = safe_float(pick("previousClose"))
            if live_price is not None:
                source = "yfinance fast_info"
        except Exception:
            live_price = None
        if live_price is None:
            try:
                hist = yf.Ticker(ticker).history(period="1d", interval="1m", prepost=True, auto_adjust=True)
                if not hist.empty:
                    live_price = safe_float(hist["Close"].dropna().iloc[-1])
                    source = "yfinance 1m prepost"
            except Exception:
                live_price = None
        if live_price is None:
            live_price = fallback
            source = "fallback: daily close"
        rows[ticker] = {
            "Ticker": ticker,
            "Close_Price": fallback,
            "Live_Price": live_price,
            "Diff": live_price - fallback,
            "Diff_%": ((live_price / fallback - 1) * 100) if fallback else 0.0,
            "Previous_Close": previous_close,
            "Source": source,
            "Checked_At": checked_at,
        }
    return rows, checked_at


realtime_snapshot, live_checked_at = get_realtime_snapshot(["QQQ", "QLD", "TQQQ"], prices_close)
prices_live = {ticker: float(realtime_snapshot[ticker]["Live_Price"]) for ticker in ["QQQ", "QLD", "TQQQ"]}
print("실시간 현재가 확인:", live_checked_at)
for t in ["QQQ", "QLD", "TQQQ"]:
    s = realtime_snapshot[t]
    print(f"{t}: 종가 {s['Close_Price']:.2f} / 현재가 {s['Live_Price']:.2f} / 괴리 {s['Diff_%']:+.2f}% / {s['Source']}")


def make_live_today_row(base_row, live_prices):
    row = base_row.copy()
    row["QQQ_Close"] = float(live_prices["QQQ"])
    row["QLD_Close"] = float(live_prices["QLD"])
    row["TQQQ_Close"] = float(live_prices["TQQQ"])
    row["Above_MA20"] = row["QQQ_Close"] > row["QQQ_MA20"]
    row["Above_MA200"] = row["QQQ_Close"] > row["QQQ_MA200"]
    high_1y = max(float(row["Rolling_1Y_High"]), float(row["QQQ_Close"])) if not pd.isna(row["Rolling_1Y_High"]) else float(row["QQQ_Close"])
    row["Rolling_1Y_High"] = high_1y
    row["Drawdown_From_1Y_High"] = row["QQQ_Close"] / high_1y - 1 if high_1y else 0.0
    # MA50_Slope_Up, RSI, Stoch, CCI, ADX, OBV는 확정 일봉 기반 보조지표를 그대로 유지합니다.
    return row


live_today = make_live_today_row(today, prices_live)

# ======================================================
# 6. 계좌 평가 / 전략 판단
# ======================================================

position_value = HOLDINGS["QQQ"]*prices["QQQ"] + HOLDINGS["QLD"]*prices["QLD"] + HOLDINGS["TQQQ"]*prices["TQQQ"]
position_cost = HOLDINGS["QQQ"]*AVG_PRICE["QQQ"] + HOLDINGS["QLD"]*AVG_PRICE["QLD"] + HOLDINGS["TQQQ"]*AVG_PRICE["TQQQ"]
total_equity = CASH + position_value
position_return = position_value / position_cost - 1 if position_cost > 0 else 0.0
live_position_value = HOLDINGS["QQQ"]*prices_live["QQQ"] + HOLDINGS["QLD"]*prices_live["QLD"] + HOLDINGS["TQQQ"]*prices_live["TQQQ"]
live_total_equity = CASH + live_position_value
live_position_return = live_position_value / position_cost - 1 if position_cost > 0 else 0.0

def get_holding_combo_and_target():
    assets = [t for t in ["TQQQ", "QLD", "QQQ"] if HOLDINGS[t] > 0 and AVG_PRICE[t] > 0]
    s = set(assets)
    if s == {"TQQQ"}: return "TQQQ 단독", TARGETS["TQQQ"]
    if s == {"QLD"}: return "QLD 단독", TARGETS["QLD"]
    if s == {"QQQ"}: return "QQQ 단독", TARGETS["QQQ"]
    if s == {"TQQQ","QLD"}: return "TQQQ + QLD", TARGETS["TQQQ_QLD"]
    if s == {"TQQQ","QQQ"}: return "TQQQ + QQQ", TARGETS["TQQQ_QQQ"]
    if s == {"QLD","QQQ"}: return "QLD + QQQ", TARGETS["QLD_QQQ"]
    if s == {"TQQQ","QLD","QQQ"}: return "TQQQ + QLD + QQQ", TARGETS["ALL"]
    return "보유 없음", None

holding_combo, target_return = get_holding_combo_and_target()
auto_sell_signal = target_return is not None and position_cost > 0 and position_return >= target_return
sold_today = MANUAL_SOLD_TODAY or auto_sell_signal

def get_tqqq_dd_div(drawdown):
    if drawdown > -0.05: return DD_STEP_DIV["DD_0_5"], "0% ~ -5%"
    if drawdown > -0.10: return DD_STEP_DIV["DD_5_10"], "-5% ~ -10%"
    if drawdown > -0.15: return DD_STEP_DIV["DD_10_15"], "-10% ~ -15%"
    return DD_STEP_DIV["DD_15_MORE"], "-15% 이하"

def decide_buy_signal_context(row, sold_today_flag, auto_sell_signal_flag, position_return_value):
    above_ma200 = bool(row["Above_MA200"])
    above_ma20 = bool(row["Above_MA20"])
    ma50_up = bool(row["MA50_Slope_Up"])
    dd = float(row["Drawdown_From_1Y_High"])

    # 우선순위 1: 기존 v14 익절/당일매도 규칙.
    # 감사 테스트와 동일하게 익절이 먼저면 다운시프트보다 익절이 우선입니다.
    if sold_today_flag:
        if auto_sell_signal_flag:
            reason = f"익절 조건 달성. {holding_combo} 현재 수익률 {position_return_value*100:.2f}% / 목표 {target_return*100:.2f}%. 오늘 전량매도하세요. 오늘은 재매수하지 않습니다. 다음 거래일부터 쿨다운 7거래일 적용."
        else:
            reason = "오늘 이미 매도 처리됨. 오늘은 재매수하지 않습니다. 남은 쿨다운/재진입 기준을 확인하세요."
        return {"action":"SELL_AND_NO_BUY", "market_type":"익절", "reason":reason, "divisions":None, "mix":{}, "dd_band":None}

    # 우선순위 2: 최종 확정 다운시프트 규칙.
    # 조건: TQQQ 보유 중 QQQ가 MA200 아래이면 TQQQ 전량매도 후 QLD99/QQQ1로 전환.
    if HOLDINGS.get("TQQQ", 0.0) > 0 and not above_ma200:
        reason = "TQQQ 보유 중 QQQ < MA200. TQQQ 전량매도 후 매도대금으로 QLD 99% / QQQ 1% 다운시프트."
        return {"action":"DOWNSHIFT", "market_type":"MA200 다운시프트", "reason":reason, "divisions":None, "mix":DOWNSHIFT_MIX.copy(), "dd_band":None}

    if above_ma200 and ma50_up:
        base_div, dd_band = get_tqqq_dd_div(dd)
        div = max(COOLDOWN_TQQQ_DIV, base_div) if COOLDOWN_DAYS_LEFT > 0 else base_div
        reason = f"강한 상승장 / TQQQ {div}분할"
        if COOLDOWN_DAYS_LEFT > 0:
            if div != base_div:
                reason = f"강한 상승장 / 쿨다운 진행 중 {COOLDOWN_DAYS_LEFT}거래일 남음 / TQQQ 최소 {COOLDOWN_TQQQ_DIV}분할 적용: 기본 {base_div}분할 -> 적용 {div}분할"
            else:
                reason = f"강한 상승장 / 쿨다운 진행 중 {COOLDOWN_DAYS_LEFT}거래일 남음 / 기본 {base_div}분할이 이미 보수 기준 이상이라 TQQQ {div}분할 적용"
        return {"action":"BUY", "market_type":"강한 상승장", "reason":reason, "divisions":div, "base_divisions":base_div, "cooldown_applied_to_tqqq": bool(COOLDOWN_DAYS_LEFT > 0), "mix":{"TQQQ":1.0}, "dd_band":dd_band}

    if above_ma200 and not ma50_up:
        return {"action":"BUY", "market_type":"약한 상승장", "reason":"약한 상승장 / QLD 60분할", "divisions":QLD_DIV, "mix":{"QLD":1.0}, "dd_band":None}

    if not above_ma200 and dd > -0.20:
        return {"action":"BUY", "market_type":"하락 초입", "reason":"200일선 아래 / 낙폭 -20% 이내 / QQQ 60분할", "divisions":QQQ_DIV, "mix":{"QQQ":1.0}, "dd_band":None}

    if not above_ma200 and -0.30 < dd <= -0.20:
        if above_ma20:
            return {"action":"BUY", "market_type":"중간 하락 회복", "reason":"200일선 아래 / -20%~-30% / 20일선 위 / QQQ 50 + QLD 50", "divisions":MIX_DIV, "mix":{"QQQ":0.5,"QLD":0.5}, "dd_band":None}
        return {"action":"BUY", "market_type":"중간 하락", "reason":"200일선 아래 / -20%~-30% / 20일선 아래 / QQQ 70 + QLD 30", "divisions":MIX_DIV, "mix":{"QQQ":0.7,"QLD":0.3}, "dd_band":None}

    if not above_ma200 and dd <= -0.30:
        if above_ma20:
            return {"action":"BUY", "market_type":"깊은 하락 회복", "reason":"200일선 아래 / -30% 이하 / 20일선 위 / QQQ 30 + QLD 70", "divisions":MIX_DIV, "mix":{"QQQ":0.3,"QLD":0.7}, "dd_band":None}
        return {"action":"BUY", "market_type":"깊은 하락", "reason":"200일선 아래 / -30% 이하 / 20일선 아래 / QQQ 50 + QLD 50", "divisions":MIX_DIV, "mix":{"QQQ":0.5,"QLD":0.5}, "dd_band":None}

    return {"action":"NO_BUY", "market_type":"대기", "reason":"조건 없음", "divisions":None, "mix":{}, "dd_band":None}


def decide_buy_signal(row):
    return decide_buy_signal_context(row, sold_today, auto_sell_signal, position_return)

signal = decide_buy_signal(today)

live_auto_sell_signal = target_return is not None and position_cost > 0 and live_position_return >= target_return
live_sold_today = MANUAL_SOLD_TODAY or live_auto_sell_signal
live_signal = decide_buy_signal_context(live_today, live_sold_today, live_auto_sell_signal, live_position_return)



def mix_label(sig):
    mix = sig.get("mix", {}) or {}
    if not mix:
        return "-"
    parts = []
    for ticker in ["TQQQ", "QLD", "QQQ"]:
        if ticker in mix and float(mix[ticker]) > 0:
            parts.append(f"{ticker} {float(mix[ticker])*100:.0f}%")
    for ticker, weight in mix.items():
        if ticker not in ["TQQQ", "QLD", "QQQ"] and float(weight) > 0:
            parts.append(f"{ticker} {float(weight)*100:.0f}%")
    return " / ".join(parts) if parts else "-"


def buy_plan_text(sig):
    if sig.get("action") != "BUY" or sig.get("divisions") is None:
        return "-"
    return f"{int(sig['divisions'])}분할 / 비율 {mix_label(sig)}"


def signal_has_tqqq_buy(sig):
    return sig.get("action") == "BUY" and float(sig.get("mix", {}).get("TQQQ", 0.0)) > 0


def cooldown_order_note(sig, ticker):
    if COOLDOWN_DAYS_LEFT <= 0:
        return ""
    if ticker == "TQQQ" and signal_has_tqqq_buy(sig):
        base = sig.get("base_divisions")
        div = sig.get("divisions")
        if base is not None and div is not None and int(div) != int(base):
            return f" / 쿨다운 {COOLDOWN_DAYS_LEFT}거래일 남음: 기본 {int(base)}분할 -> {int(div)}분할"
        return f" / 쿨다운 {COOLDOWN_DAYS_LEFT}거래일 남음: TQQQ {div}분할"
    if sig.get("action") == "BUY":
        return f" / 쿨다운 {COOLDOWN_DAYS_LEFT}거래일 남음: TQQQ 주문 아님"
    return ""

def build_order_df(sig, price_map, equity_value):
    rows = []
    if sig["action"] == "SELL_AND_NO_BUY":
        # 익절/당일매도: 보유 중인 QQQ/QLD/TQQQ를 각각 몇 주 팔지 명확히 표시합니다.
        # 현금은 매도 대상이 아니며, 매도 후 오늘은 재매수하지 않습니다.
        for ticker in ["TQQQ", "QLD", "QQQ"]:
            shares = float(HOLDINGS.get(ticker, 0.0))
            if shares <= 0:
                continue
            price = float(price_map[ticker])
            amount = shares * price
            rows.append({
                "Order": "SELL", "Ticker": ticker, "Weight": 1.0,
                "Sell_Shares": round(shares, 6), "Sell_Amount": round(amount, 2),
                "Buy_Amount": 0.0, "Price": round(price, 2), "Estimated_Shares": 0.0,
                "Divisions": None, "Mix_Label": "전량매도",
                "Note": f"전량매도하세요 / 오늘 재매수 금지 / 다음 거래일부터 쿨다운 {SELL_COOLDOWN_DAYS_AFTER_PROFIT}거래일"
            })
    elif sig["action"] == "BUY" and sig["divisions"] is not None and CASH > 0:
        total_buy = min(equity_value / sig["divisions"], CASH)
        for ticker, weight in sig["mix"].items():
            amount = total_buy * weight
            price = float(price_map[ticker])
            rows.append({
                "Order": "BUY", "Ticker": ticker, "Weight": weight,
                "Sell_Shares": 0.0, "Sell_Amount": 0.0,
                "Buy_Amount": round(amount,2), "Price": round(price,2),
                "Estimated_Shares": round(amount/price,6) if price > 0 else 0.0,
                "Divisions": int(sig["divisions"]), "Mix_Label": mix_label(sig),
                "Note": f"{int(sig['divisions'])}분할 신규매수 / 비율 {mix_label(sig)}" + cooldown_order_note(sig, ticker)
            })
    elif sig["action"] == "DOWNSHIFT" and HOLDINGS.get("TQQQ", 0.0) > 0:
        tqqq_price = float(price_map["TQQQ"])
        tqqq_shares = float(HOLDINGS.get("TQQQ", 0.0))
        proceeds = tqqq_shares * tqqq_price
        rows.append({
            "Order": "SELL", "Ticker": "TQQQ", "Weight": 1.0,
            "Sell_Shares": round(tqqq_shares, 6), "Sell_Amount": round(proceeds, 2),
            "Buy_Amount": 0.0, "Price": round(tqqq_price, 2), "Estimated_Shares": 0.0,
            "Divisions": None, "Mix_Label": "TQQQ 전량매도 -> QLD 99% / QQQ 1%",
            "Note": "MA200 다운시프트: TQQQ 전량매도"
        })
        for ticker, weight in sig["mix"].items():
            amount = proceeds * float(weight)
            price = float(price_map[ticker])
            rows.append({
                "Order": "BUY", "Ticker": ticker, "Weight": float(weight),
                "Sell_Shares": 0.0, "Sell_Amount": 0.0,
                "Buy_Amount": round(amount, 2), "Price": round(price, 2),
                "Estimated_Shares": round(amount/price, 6) if price > 0 else 0.0,
                "Divisions": None, "Mix_Label": "QLD 99% / QQQ 1%",
                "Note": "TQQQ 매도대금 재투입: QLD99/QQQ1"
            })
    return pd.DataFrame(rows)


buy_df = build_order_df(signal, prices, total_equity)
live_buy_df = build_order_df(live_signal, prices_live, live_total_equity)


def signal_has_tqqq_buy(sig):
    return sig.get("action") == "BUY" and float(sig.get("mix", {}).get("TQQQ", 0.0)) > 0


def cooldown_status_text(sig):
    if sig.get("action") == "SELL_AND_NO_BUY":
        return f"익절/당일매도: 오늘 재매수 금지. 다음 거래일부터 {SELL_COOLDOWN_DAYS_AFTER_PROFIT}거래일 쿨다운 시작."
    if COOLDOWN_DAYS_LEFT > 0:
        if signal_has_tqqq_buy(sig):
            base = sig.get("base_divisions")
            div = sig.get("divisions")
            if base is not None and div is not None and int(div) != int(base):
                return f"쿨다운 진행 중: {COOLDOWN_DAYS_LEFT}거래일 남음. TQQQ 최소 {COOLDOWN_TQQQ_DIV}분할 적용: 기본 {int(base)}분할 -> 적용 {int(div)}분할."
            return f"쿨다운 진행 중: {COOLDOWN_DAYS_LEFT}거래일 남음. 오늘 TQQQ 매수는 {div}분할 기준."
        if sig.get("action") == "BUY":
            return f"쿨다운 진행 중: {COOLDOWN_DAYS_LEFT}거래일 남음. 오늘 주문은 TQQQ가 아니므로 TQQQ 최소분할 제한 영향 없음."
        return f"쿨다운 진행 중: {COOLDOWN_DAYS_LEFT}거래일 남음. 오늘은 매수 조건 없음."
    return "쿨다운 없음/종료. 정상 분할 기준 적용."


def cooldown_order_note(sig, ticker):
    if COOLDOWN_DAYS_LEFT <= 0:
        return ""
    if ticker == "TQQQ" and signal_has_tqqq_buy(sig):
        base = sig.get("base_divisions")
        div = sig.get("divisions")
        if base is not None and div is not None and int(div) != int(base):
            return f" / 쿨다운 {COOLDOWN_DAYS_LEFT}거래일 남음: 기본 {int(base)}분할 -> {int(div)}분할"
        return f" / 쿨다운 {COOLDOWN_DAYS_LEFT}거래일 남음: TQQQ {div}분할"
    if sig.get("action") == "BUY":
        return f" / 쿨다운 {COOLDOWN_DAYS_LEFT}거래일 남음: TQQQ 주문 아님"
    return ""


cooldown_status = cooldown_status_text(signal)
live_cooldown_status = cooldown_status_text(live_signal)


def make_live_final_status(daily_sig, live_sig):
    qqq_diff = float(realtime_snapshot["QQQ"]["Diff_%"])
    biggest_abs_diff = max(abs(float(realtime_snapshot[t]["Diff_%"])) for t in ["QQQ", "QLD", "TQQQ"])
    if live_sig["action"] == "DOWNSHIFT":
        return {
            "status": "LIVE_DOWNSHIFT",
            "title": "실시간 다운시프트",
            "final": "실시간 현재가 기준으로 QQQ < MA200이며 TQQQ를 보유 중입니다. TQQQ 전량매도 후 매도대금으로 QLD 99% / QQQ 1% 전환 수량을 확인하세요.",
            "color": CARD["sell"], "soft": CARD["sell_soft"], "badge": "DS 99/1"
        }
    if live_sig["action"] == "SELL_AND_NO_BUY":
        return {
            "status": "LIVE_SELL",
            "title": "실시간 매도",
            "final": "실시간 현재가 기준으로 익절/매도 조건이 충족되었습니다. 전량매도하세요. 아래 종목별 매도수량을 확인하고, 오늘은 재매수하지 않습니다. 다음 거래일부터 쿨다운 7거래일 기준을 적용하세요.",
            "color": CARD["sell"], "soft": CARD["sell_soft"], "badge": "LIVE SELL"
        }
    if daily_sig["action"] == "BUY" and live_sig["action"] == "BUY":
        if biggest_abs_diff >= 2.5:
            return {
                "status": "PULLBACK_WAIT",
                "title": "추격매수 주의",
                "final": f"종가와 실시간 현재가 괴리가 최대 {biggest_abs_diff:.2f}%입니다. 신호는 유지되지만 주문 전 가격을 다시 확인하세요.",
                "color": CARD["wait"], "soft": CARD["wait_soft"], "badge": "LIVE CHECK"
            }
        return {
            "status": "LIVE_BUY_OK",
            "title": "실시간 매수 가능",
            "final": "종가 기준 매수 판단이 실시간 현재가 기준에서도 유지됩니다. 실시간 가격 기준 주문수량을 확인하세요.",
            "color": CARD["buy"], "soft": CARD["buy_soft"], "badge": "LIVE BUY"
        }
    if daily_sig["action"] == "BUY" and live_sig["action"] != "BUY":
        return {
            "status": "LIVE_BUY_BLOCK",
            "title": "매수 보류",
            "final": "종가 기준은 매수였지만, 실시간 현재가 기준으로 같은 판단식이 매수 조건을 통과하지 못했습니다. 주문 전 재확인이 필요합니다.",
            "color": CARD["sell"], "soft": CARD["sell_soft"], "badge": "BLOCK"
        }
    if daily_sig["action"] != "BUY" and live_sig["action"] == "BUY":
        return {
            "status": "LIVE_BUY_WATCH",
            "title": "실시간 회복 감지",
            "final": "종가 기준은 매수가 아니지만, 실시간 현재가 기준으로는 매수 조건이 보입니다. 백테스트 기준을 유지하려면 종가 재확인이 우선입니다.",
            "color": CARD["wait"], "soft": CARD["wait_soft"], "badge": "WATCH"
        }
    if qqq_diff <= -2.0:
        return {
            "status": "LIVE_RISK_ALERT",
            "title": "장중 약세 경고",
            "final": f"종가 기준 판단은 유지되지만 QQQ 현재가가 종가 대비 {qqq_diff:.2f}% 약세입니다. 보유/추가매수 전 재확인하세요.",
            "color": CARD["wait"], "soft": CARD["wait_soft"], "badge": "RISK"
        }
    return {
        "status": "LIVE_HOLD_OK",
        "title": "실시간 유지",
        "final": "종가 기준 판단과 실시간 현재가 기준 판단이 크게 충돌하지 않습니다.",
        "color": CARD["blue"], "soft": CARD["blue_soft"], "badge": "LIVE OK"
    }


live_final = make_live_final_status(signal, live_signal)

summary = pd.DataFrame([{
    "Date": today_date.strftime("%Y-%m-%d"), "Strategy": STRATEGY_NAME,
    "QQQ_Close": round(prices["QQQ"],2), "QLD_Close": round(prices["QLD"],2), "TQQQ_Close": round(prices["TQQQ"],2),
    "QQQ_MA20": round(float(today["QQQ_MA20"]),2), "QQQ_MA50": round(float(today["QQQ_MA50"]),2), "QQQ_MA100": round(float(today["QQQ_MA100"]),2), "QQQ_MA200": round(float(today["QQQ_MA200"]),2),
    "RSI14": round(float(today["RSI14"]),2), "STOCH_K": round(float(today["STOCH_K"]),2), "CCI20": round(float(today["CCI20"]),2), "ADX14": round(float(today["ADX14"]),2),
    "Drawdown_From_1Y_High_%": round(float(today["Drawdown_From_1Y_High"])*100,2),
    "Cash": round(CASH,2), "Position_Value": round(position_value,2), "Position_Cost": round(position_cost,2), "Position_Return_%": round(position_return*100,2), "Total_Equity": round(total_equity,2),
    "Action": signal["action"], "Market_Type": signal["market_type"], "Reason": signal["reason"],
    "Buy_Divisions": signal.get("divisions"), "Buy_Mix": mix_label(signal), "Buy_Plan_Text": buy_plan_text(signal),
    "Cooldown_Status": cooldown_status, "Cooldown_Days_Left_Input": COOLDOWN_DAYS_LEFT, "Cooldown_Auto_Note": COOLDOWN_AUTO_NOTE, "Last_Profit_Sell_Date": str(getattr(_args, "last_profit_sell_date", "")),
    "TQQQ_Cooldown_Applied_To_Order": bool(signal_has_tqqq_buy(signal) and COOLDOWN_DAYS_LEFT > 0),
    "Sell_Cooldown_Days_After_Profit": SELL_COOLDOWN_DAYS_AFTER_PROFIT if signal["action"] == "SELL_AND_NO_BUY" else 0,
    "Downshift_Trigger": DOWNSHIFT_TRIGGER, "Downshift_Mix": "QLD 99% / QQQ 1%", "TQQQ_Downshift_To_Zero": DOWNSHIFT_TQQQ_TO_ZERO,
    "Live_Checked_At": live_checked_at, "Live_Action": live_signal["action"], "Live_Market_Type": live_signal["market_type"],
    "Live_Buy_Divisions": live_signal.get("divisions"), "Live_Buy_Mix": mix_label(live_signal), "Live_Buy_Plan_Text": buy_plan_text(live_signal),
    "Live_Cooldown_Status": live_cooldown_status,
    "Live_TQQQ_Cooldown_Applied_To_Order": bool(signal_has_tqqq_buy(live_signal) and COOLDOWN_DAYS_LEFT > 0),
    "Live_Sell_Cooldown_Days_After_Profit": SELL_COOLDOWN_DAYS_AFTER_PROFIT if live_signal["action"] == "SELL_AND_NO_BUY" else 0,
    "Live_Final_Status": live_final["status"], "Live_Final_Text": live_final["final"],
    "QQQ_Live": round(prices_live["QQQ"],2), "QQQ_Live_Diff_%": round(float(realtime_snapshot["QQQ"]["Diff_%"]),2),
    "Live_Position_Value": round(live_position_value,2), "Live_Position_Return_%": round(live_position_return*100,2), "Live_Total_Equity": round(live_total_equity,2),
}])

print(summary.to_string(index=False))
if buy_df.empty:
    print("오늘 매수 없음")
else:
    print(buy_df.to_string(index=False))
print("="*80)
print("실시간 판단 요약")
print("="*80)
print("종가 판단:", signal["action"], "/", signal["market_type"])
print("실시간 판단:", live_signal["action"], "/", live_signal["market_type"])
print("최종 실행:", live_final["status"], "/", live_final["final"])
print("쿨다운 상태(종가):", cooldown_status)
print("매수 기준(종가):", buy_plan_text(signal))
print("쿨다운 상태(실시간):", live_cooldown_status)
print("매수 기준(실시간):", buy_plan_text(live_signal))
if live_buy_df.empty:
    print("실시간 기준 매수 주문 없음")
else:
    print(live_buy_df.to_string(index=False))

# ======================================================
# 6-1. 실행용 CSV 저장
# ======================================================
summary_csv_path = os.path.join(OUTPUT_DIR, "today_summary_v14_99_1.csv")
order_csv_path = os.path.join(OUTPUT_DIR, "today_order_plan_v14_99_1.csv")
live_order_csv_path = os.path.join(OUTPUT_DIR, "today_live_order_plan_v14_99_1.csv")

def safe_to_csv(df, path):
    try:
        df.to_csv(path, index=False, encoding="utf-8-sig")
        return path
    except PermissionError:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base, ext = os.path.splitext(path)
        alt = f"{base}_{stamp}{ext}"
        df.to_csv(alt, index=False, encoding="utf-8-sig")
        print(f"[warning] 기존 CSV가 열려 있어 대체 파일로 저장: {alt}")
        return alt

summary_csv_path = safe_to_csv(summary, summary_csv_path)
order_csv_path = safe_to_csv(buy_df, order_csv_path)
live_order_csv_path = safe_to_csv(live_buy_df, live_order_csv_path)
print("CSV 저장:", summary_csv_path)
print("CSV 저장:", order_csv_path)
print("CSV 저장:", live_order_csv_path)

def action_info():
    if signal["action"] == "SELL_AND_NO_BUY":
        return {"title":"전량매도", "verb":"매도하세요", "badge":"SELL", "main":CARD["sell"], "dark":CARD["sell_dark"], "soft":CARD["sell_soft"]}
    if signal["action"] == "DOWNSHIFT" and not buy_df.empty:
        return {"title":"TQQQ→QLD99/QQQ1", "verb":"다운시프트", "badge":"DS 99/1", "main":CARD["sell"], "dark":CARD["sell_dark"], "soft":CARD["sell_soft"]}
    if signal["action"] == "BUY" and not buy_df.empty:
        tickers = buy_df["Ticker"].tolist()
        return {"title": tickers[0] if len(tickers)==1 else "분할매수", "verb":"매수하세요", "badge":"BUY", "main":CARD["buy"], "dark":CARD["buy_dark"], "soft":CARD["buy_soft"]}
    return {"title":"대기", "verb":"쉬세요", "badge":"WAIT", "main":CARD["wait"], "dark":CARD["wait_dark"], "soft":CARD["wait_soft"]}

# ======================================================
# 7. 그래프 생성
# ======================================================

def create_qqq_price_graph(path=os.path.join(OUTPUT_DIR, "qqq_price_ma_graph.png")):
    c = data.tail(260).copy()
    plt.figure(figsize=(14,7.5))
    ax = plt.gca()
    ax.plot(c.index, c["QQQ_Close"], linewidth=2.8, label="QQQ")
    ax.plot(c.index, c["QQQ_MA20"], linewidth=2.0, label="MA20")
    ax.plot(c.index, c["QQQ_MA50"], linewidth=2.0, label="MA50")
    ax.plot(c.index, c["QQQ_MA100"], linewidth=2.0, label="MA100")
    ax.plot(c.index, c["QQQ_MA200"], linewidth=2.5, label="MA200")
    ax.plot(c.index, c["BB_UPPER"], linewidth=1.2, linestyle="--", label="BB Upper")
    ax.plot(c.index, c["BB_LOWER"], linewidth=1.2, linestyle="--", label="BB Lower")
    ax.set_title("QQQ Price · MA20/50/100/200 · Bollinger Bands", fontsize=19, fontweight="bold", pad=18)
    ax.set_xlabel("Date", fontsize=13, fontweight="bold")
    ax.set_ylabel("Price", fontsize=13, fontweight="bold")
    ax.grid(True, alpha=0.25)
    ax.legend(loc="upper left", frameon=True, prop={"weight":"bold"})
    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
    plt.tight_layout(); plt.savefig(path, dpi=180, bbox_inches="tight"); plt.close()
    return path

def create_drawdown_graph(path=os.path.join(OUTPUT_DIR, "qqq_drawdown_graph.png")):
    c = data.tail(260).copy()
    plt.figure(figsize=(14,6))
    ax = plt.gca()
    ax.plot(c.index, c["Drawdown_From_1Y_High"]*100, linewidth=2.8, label="Drawdown")
    for level in [-5,-10,-15,-20,-30]:
        ax.axhline(level, linestyle="--", linewidth=1.4, alpha=0.65)
        ax.text(c.index[3], level+0.4, f"{level}%", fontsize=11, fontweight="bold")
    ax.set_title("QQQ Drawdown from 1-Year High", fontsize=21, fontweight="bold", pad=18)
    ax.set_xlabel("Date", fontsize=13, fontweight="bold")
    ax.set_ylabel("Drawdown (%)", fontsize=13, fontweight="bold")
    ax.grid(True, alpha=0.25); ax.legend(loc="lower left", frameon=True, prop={"weight":"bold"})
    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
    plt.tight_layout(); plt.savefig(path, dpi=180, bbox_inches="tight"); plt.close()
    return path

def create_buy_order_graph(path=os.path.join(OUTPUT_DIR, "today_buy_order_graph.png")):
    plt.figure(figsize=(9,6))
    ax = plt.gca()
    if buy_df.empty:
        ax.bar(["No Buy"], [0]); ax.set_title("Today Order · No Buy", fontsize=21, fontweight="bold", pad=18)
    else:
        labels = buy_df["Ticker"].tolist(); values = buy_df["Buy_Amount"].tolist()
        ax.bar(labels, values); ax.set_title("Today Buy Order Amount", fontsize=21, fontweight="bold", pad=18)
        for i,v in enumerate(values): ax.text(i, v, f"${v:,.0f}", ha="center", va="bottom", fontsize=14, fontweight="bold")
    ax.set_ylabel("Buy Amount ($)", fontsize=13, fontweight="bold")
    ax.grid(True, axis="y", alpha=0.25); ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
    plt.tight_layout(); plt.savefig(path, dpi=180, bbox_inches="tight"); plt.close()
    return path

qqq_price_graph_path = create_qqq_price_graph()
drawdown_graph_path = create_drawdown_graph()
buy_order_graph_path = create_buy_order_graph()

# ======================================================
# 8. 카드 생성
# ======================================================

def create_today_action_card(path=os.path.join(OUTPUT_DIR, "today_action_card.png")):
    W,H = 1080,1920
    img = Image.new("RGB", (W,H), CARD["bg_top"]); draw_gradient(img, CARD["bg_top"], CARD["bg_bottom"]); img = img.convert("RGBA"); d = ImageDraw.Draw(img)
    info = action_info()
    d.text((60,55), "QQQ · QLD · TQQQ", font=get_font(30), fill=CARD["muted"])
    d.text((60,105), "오늘 매수·매도 판단", font=get_font(56), fill=CARD["white"])
    d.text((60,172), today_date.strftime("%Y-%m-%d"), font=get_font(30), fill=CARD["muted"])
    rounded(d, (805,70,1015,132), radius=31, fill=info["soft"]); d.text((910,101), info["badge"], font=get_font(31), fill=info["main"], anchor="mm")
    card(img, (55,245,1025,455), radius=44, fill=CARD["white"]); d = ImageDraw.Draw(img)
    rounded(d, (95,290,260,343), radius=26, fill=info["soft"]); d.text((177,316), "오늘 결론", font=get_font(23), fill=info["main"], anchor="mm")
    d.text((95,392), info["title"], font=get_font(48), fill=info["dark"]); d.text((365,394), info["verb"], font=get_font(45), fill=info["main"])
    card(img, (55,500,1025,960), radius=44, fill=CARD["white"]); d = ImageDraw.Draw(img)
    d.text((95,558), "오늘의 주문", font=get_font(40), fill=CARD["text"])
    header_y=635; rounded(d,(90,header_y,990,header_y+60),radius=22,fill=CARD["navy2"])
    for x,txt in [(125,"주문"),(260,"종목"),(410,"현재가"),(610,"매수가")]: d.text((x,header_y+30), txt, font=get_font(22), fill=CARD["white"], anchor="lm")
    d.text((905,header_y+30), "수량", font=get_font(22), fill=CARD["white"], anchor="mm")
    row_y = header_y + 82
    if signal["action"] in ["SELL_AND_NO_BUY", "BUY", "DOWNSHIFT"] and not buy_df.empty:
        for _, row in buy_df.iterrows():
            is_sell = str(row.get("Order", "BUY")) == "SELL"
            row_soft = CARD["sell_soft"] if is_sell else CARD["buy_soft"]
            row_main = CARD["sell"] if is_sell else CARD["buy"]
            order_label = "매도" if is_sell else "매수"
            amount = row.get("Sell_Amount", 0.0) if is_sell else row.get("Buy_Amount", 0.0)
            shares = row.get("Sell_Shares", 0.0) if is_sell else row.get("Estimated_Shares", 0.0)
            rounded(d,(90,row_y,990,row_y+125),radius=28,fill=row_soft)
            d.text((125,row_y+43),order_label,font=get_font(30),fill=row_main,anchor="lm"); d.text((260,row_y+43),row["Ticker"],font=get_font(31),fill=CARD["text"],anchor="lm")
            d.text((410,row_y+43),money(row["Price"]),font=get_font(24),fill=CARD["text"],anchor="lm"); d.text((610,row_y+43),money(amount),font=get_font(24),fill=row_main,anchor="lm"); d.text((905,row_y+43),str(shares),font=get_font(21),fill=CARD["text"],anchor="mm")
            
            if signal["action"] == "SELL_AND_NO_BUY":
                note = str(row.get("Note", f"전량매도 / 오늘 재매수 금지 / 쿨다운 {SELL_COOLDOWN_DAYS_AFTER_PROFIT}거래일"))
            elif signal["action"] == "DOWNSHIFT":
                note = str(row.get("Note", "TQQQ 다운시프트"))
            else:
                note = f"{signal['divisions']}분할 기준 · 총자산 기준 분할매수"
            
            d.text((125,row_y+92),note,font=get_font(22),fill=CARD["subtext"])
            row_y += 140
    else:
        rounded(d,(90,row_y,990,row_y+125),radius=28,fill=CARD["wait_soft"])
        d.text((125,row_y+45),"대기",font=get_font(30),fill=CARD["wait"],anchor="lm"); d.text((260,row_y+45),"-",font=get_font(30),fill=CARD["text"],anchor="lm")
        d.text((410,row_y+45),"-",font=get_font(26),fill=CARD["text"],anchor="lm"); d.text((610,row_y+45),"$0",font=get_font(26),fill=CARD["wait"],anchor="lm"); d.text((905,row_y+45),"0",font=get_font(26),fill=CARD["text"],anchor="mm")
        d.text((125,row_y+92),"오늘은 매수 조건이 없습니다.",font=get_font(23),fill=CARD["subtext"])
    card(img, (55,1010,1025,1325), radius=44, fill=CARD["white"]); d = ImageDraw.Draw(img)
    d.text((95,1068),"판단 사유",font=get_font(40),fill=CARD["purple"])
    cx=95; cy=1130; cx=chip(d,cx,cy,signal["market_type"],CARD["blue_soft"],CARD["blue"],font_size=22)
    if signal["action"]=="BUY": cx=chip(d,cx,cy,f"{signal['divisions']}분할",CARD["buy_soft"],CARD["buy"],font_size=22)
    if signal["dd_band"] is not None: cx=chip(d,cx,cy,signal["dd_band"],CARD["wait_soft"],CARD["wait"],font_size=22)
    if COOLDOWN_DAYS_LEFT > 0 and signal["action"] != "SELL_AND_NO_BUY":
        cx=chip(d,cx,cy,f"쿨다운 {COOLDOWN_DAYS_LEFT}일",CARD["wait_soft"],CARD["wait"],font_size=22)
        if signal_has_tqqq_buy(signal):
            cx=chip(d,cx,cy,"TQQQ 최소분할 적용",CARD["sell_soft"],CARD["sell"],font_size=22)
    if signal["action"]=="SELL_AND_NO_BUY":
        cx=chip(d,cx,cy,"전량매도",CARD["sell_soft"],CARD["sell"],font_size=22)
        cx=chip(d,cx,cy,"오늘 재매수 금지",CARD["sell_soft"],CARD["sell"],font_size=22)
        cx=chip(d,cx,cy,f"쿨다운 {SELL_COOLDOWN_DAYS_AFTER_PROFIT}거래일",CARD["wait_soft"],CARD["wait"],font_size=22)
    draw_text_box(d,95,1205,signal["reason"],get_font(27),CARD["text"],850,line_gap=10,max_lines=3)
    card(img,(55,1375,1025,1625),radius=44,fill=CARD["white"]); d=ImageDraw.Draw(img)
    d.text((95,1435),"적용 정보",font=get_font(39),fill=CARD["text"])
    kv(d,95,1510,"시장 상태",signal["market_type"],value_color=CARD["purple"],size=28)
    cooldown_display = f"매도 후 {SELL_COOLDOWN_DAYS_AFTER_PROFIT}거래일 예정" if signal["action"] == "SELL_AND_NO_BUY" else (f"진행 중 {COOLDOWN_DAYS_LEFT}일" if COOLDOWN_DAYS_LEFT > 0 else "없음/종료")
    kv(d,95,1570,"쿨다운",cooldown_display,value_color=CARD["wait"],size=28)
    # 상세 쿨다운 문구는 카드 주문/판단 사유와 CSV/텔레그램에 표시합니다.
    draw_yellow_footer(d,W,1690)
    img = img.convert("RGB"); img.save(path, quality=95); return path

def create_analysis_card(path=os.path.join(OUTPUT_DIR, "today_analysis_card.png")):
    W,H = 2320,1080
    img = Image.new("RGB", (W,H), CARD["card_bg"]); img=img.convert("RGBA"); d=ImageDraw.Draw(img); info=action_info()
    rounded(d,(0,0,W,220),radius=0,fill=CARD["navy"])
    d.text((70,42),"TODAY REPORT",font=get_font(31),fill=CARD["cyan"]); d.text((70,92),"오늘 분석 리포트",font=get_font(58),fill=CARD["white"]); d.text((70,160),today_date.strftime("%Y-%m-%d"),font=get_font(29),fill=CARD["muted"])
    rounded(d,(1510,70,2250,145),radius=38,fill=info["soft"]); d.text((1880,107),f"{info['title']} · {info['verb']}",font=get_font(32),fill=info["main"],anchor="mm")
    d.text((70,285),"시장 핵심 지표",font=get_font(41),fill=CARD["purple"])
    def metric(x,y,title,value,color,soft):
        rounded(d,(x,y,x+500,y+135),radius=30,fill=soft); d.text((x+28,y+25),title,font=get_font(23),fill=color); d.text((x+28,y+75),value,font=get_font(34),fill=CARD["text"])
    metric(70,350,"QQQ 현재가",money(prices["QQQ"]),CARD["blue"],CARD["blue_soft"])
    metric(610,350,"1년 고점 대비 낙폭",percent(today["Drawdown_From_1Y_High"]*100),CARD["wait"],CARD["wait_soft"])
    metric(1150,350,"200일선 위","YES" if bool(today["Above_MA200"]) else "NO",CARD["buy"] if bool(today["Above_MA200"]) else CARD["sell"],CARD["buy_soft"] if bool(today["Above_MA200"]) else CARD["sell_soft"])
    metric(1690,350,"MA50 상승중","YES" if bool(today["MA50_Slope_Up"]) else "NO",CARD["buy"] if bool(today["MA50_Slope_Up"]) else CARD["sell"],CARD["buy_soft"] if bool(today["MA50_Slope_Up"]) else CARD["sell_soft"])
    card(img,(70,535,740,790),radius=42,fill=CARD["white"]); d=ImageDraw.Draw(img); d.text((115,590),"이동평균선",font=get_font(38),fill=CARD["blue"])
    kv(d,115,655,"QQQ 20일선",money(today["QQQ_MA20"]),value_color=CARD["cyan"],size=24,value_x=680); kv(d,115,700,"QQQ 50일선",money(today["QQQ_MA50"]),value_color=CARD["purple"],size=24,value_x=680); kv(d,115,745,"QQQ 100일선",money(today["QQQ_MA100"]),value_color=CARD["pink"],size=24,value_x=680)
    card(img,(790,535,1510,790),radius=42,fill=CARD["white"]); d=ImageDraw.Draw(img); d.text((835,590),"계좌 상태",font=get_font(38),fill=CARD["buy"])
    kv(d,835,655,"현금",money(CASH),value_color=CARD["blue"],size=24,value_x=1440); kv(d,835,700,"포지션 수익률",percent(position_return*100),value_color=CARD["buy"] if position_return>=0 else CARD["sell"],size=24,value_x=1440); kv(d,835,745,"총자산",money(total_equity),value_color=CARD["buy"],size=24,value_x=1440)
    card(img,(1560,535,2250,790),radius=42,fill=CARD["white"]); d=ImageDraw.Draw(img); d.text((1605,590),"보조지표 현재값",font=get_font(38),fill=CARD["pink"])
    kv(d,1605,655,"RSI 14",num2(today["RSI14"]),value_color=CARD["purple"],size=24,value_x=2185); kv(d,1605,700,"Stoch K/D",f"{num2(today['STOCH_K'])} / {num2(today['STOCH_D'])}",value_color=CARD["blue"],size=24,value_x=2185); kv(d,1605,745,"ADX 14",num2(today["ADX14"]),value_color=CARD["sell"],size=24,value_x=2185)
    draw_yellow_footer(d,W,900)
    img = img.convert("RGB"); img.save(path, quality=95); return path

def create_strategy_card(path=os.path.join(OUTPUT_DIR, "strategy_card_simple.png")):
    W,H = 1880,1750
    img = Image.new("RGB", (W,H), CARD["bg_top"]); draw_gradient(img, CARD["bg_top"], CARD["bg_bottom"]); img=img.convert("RGBA"); d=ImageDraw.Draw(img)
    d.text((70,55),"최종 전략 요약",font=get_font(58),fill=CARD["white"]); d.text((70,125),"QQQ · QLD · TQQQ 자동 판단 시스템",font=get_font(30),fill=CARD["muted"])
    rounded(d,(70,185,450,245),radius=30,fill=CARD["yellow"]); d.text((260,215),"DD쿨차등_C 전략",font=get_font(25),fill=CARD["text"],anchor="mm")
    sections=[("1. 강한 상승장","QQQ 200일선 위\nMA50 상승","TQQQ 17·19·21·24분할",CARD["buy"],CARD["buy_soft"]),("2. 약한 상승장","QQQ 200일선 위\nMA50 상승 아님","QLD 60분할",CARD["wait"],CARD["wait_soft"]),("3. 다운시프트","TQQQ 보유 + QQQ 200일선 아래\nMA200 위험전환","TQQQ 전량매도 → QLD99 + QQQ1",CARD["sell"],CARD["sell_soft"]),("4. 익절","목표수익 도달 시\n전량매도","그날 재매수 없음 · 7거래일 쿨다운",CARD["blue"],CARD["blue_soft"])]
    for x,sec in zip([70,515,960,1405],sections):
        title,cond,point,color,soft=sec; y=305; card(img,(x,y,x+395,y+300),radius=36,fill=CARD["white"]); d=ImageDraw.Draw(img)
        rounded(d,(x+25,y+30,x+365,y+82),radius=26,fill=soft); d.text((x+195,y+56),title,font=get_font(24),fill=color,anchor="mm")
        lines=cond.split("\n"); d.text((x+35,y+125),lines[0],font=get_font(25),fill=CARD["text"]); d.text((x+35,y+160),lines[1],font=get_font(25),fill=CARD["text"])
        rounded(d,(x+30,y+225,x+365,y+270),radius=22,fill=soft); d.text((x+197,y+247),point,font=get_font(20),fill=color,anchor="mm")
    card(img,(70,665,900,1295),radius=42,fill=CARD["white"]); d=ImageDraw.Draw(img)
    d.text((115,720),"사용 데이터 / 보조지표",font=get_font(36),fill=CARD["purple"])
    d.text((115,785),"가격 데이터",font=get_font(25),fill=CARD["blue"]); d.text((115,825),"QQQ · QLD · TQQQ 종가 / QQQ OHLCV",font=get_font(22),fill=CARD["text"])
    d.text((115,885),"추세 지표",font=get_font(25),fill=CARD["buy"]); d.text((115,925),"20일선 · 50일선 · 100일선 · 200일선",font=get_font(21),fill=CARD["text"]); d.text((115,960),"MA50 기울기",font=get_font(21),fill=CARD["text"])
    d.text((115,1020),"변동성 / 낙폭",font=get_font(25),fill=CARD["wait"]); d.text((115,1060),"볼린저밴드 20,2 · BB 폭",font=get_font(21),fill=CARD["text"]); d.text((115,1095),"1년 고점 대비 낙폭",font=get_font(21),fill=CARD["text"])
    d.text((115,1155),"보조지표",font=get_font(25),fill=CARD["pink"]); d.text((115,1195),"RSI · Stochastic · CCI · ADX",font=get_font(21),fill=CARD["text"]); d.text((115,1230),"+DI/-DI · OBV",font=get_font(21),fill=CARD["text"])
    card(img,(950,665,1810,1050),radius=42,fill=CARD["white"]); d=ImageDraw.Draw(img); d.text((995,720),"하락장 비중표",font=get_font(36),fill=CARD["purple"])
    rows=[("-20% 이내","QQQ 100%"),("-20%~-30% / 20일선 아래","QQQ 70 + QLD 30"),("-20%~-30% / 20일선 위","QQQ 50 + QLD 50"),("-30% 이하 / 20일선 위","QQQ 30 + QLD 70"),("-30% 이하 / 20일선 아래","QQQ 50 + QLD 50")]
    yy=785
    for state,buy in rows:
        d.text((995,yy),state,font=get_font(22),fill=CARD["subtext"]); d.text((1745,yy),buy,font=get_font(22),fill=CARD["text"],anchor="ra"); yy+=47
    card(img,(950,1100,1810,1450),radius=42,fill=CARD["white"]); d=ImageDraw.Draw(img); d.text((995,1155),"익절 기준",font=get_font(34),fill=CARD["sell"])
    targets=[("TQQQ 단독","+15%"),("QLD 단독","+10%"),("QQQ 단독","+5%"),("TQQQ + QLD","+12%"),("TQQQ + QQQ","+10%"),("QLD + QQQ / 전체혼합","+7%")]
    yy=1215
    for name,target in targets:
        d.text((995,yy),name,font=get_font(23),fill=CARD["subtext"]); d.text((1745,yy),target,font=get_font(23),fill=CARD["buy"],anchor="ra"); yy+=38
    draw_yellow_footer(d,W,1545)
    img = img.convert("RGB"); img.save(path, quality=95); return path

# ======================================================
# POWER TQQQ BLUE 계좌 현황판 - v15 account vibe 기반 / v14 CMD 전용
# ======================================================

def _ptb_blend_hex(c1, c2, t):
    t = max(0.0, min(1.0, float(t)))
    a = tuple(int(c1.lstrip("#")[i:i+2], 16) for i in (0, 2, 4))
    b = tuple(int(c2.lstrip("#")[i:i+2], 16) for i in (0, 2, 4))
    m = tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))
    return "#%02x%02x%02x" % m


def _ptb_draw_gradient_rect(draw, xy, top, bottom):
    x0, y0, x1, y1 = map(int, xy)
    h = max(1, y1 - y0)
    for y in range(y0, y1):
        draw.line((x0, y, x1, y), fill=_ptb_blend_hex(top, bottom, (y - y0) / h))


def _ptb_draw_glow_text(draw, xy, text_value, font, fill, anchor="mm", glow="#0B2A4A", radius=2):
    x, y = xy
    for dx in range(-radius, radius + 1):
        for dy in range(-radius, radius + 1):
            if dx or dy:
                draw.text((x + dx, y + dy), str(text_value), font=font, fill=glow, anchor=anchor)
    draw.text((x, y), str(text_value), font=font, fill=fill, anchor=anchor)


def _ptb_wrap_lines(text, max_chars):
    text = str(text or "")
    lines = []
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


def _ptb_account_daily_values(dataframe, asset, shares, idx):
    close_col = f"{asset}_Close"
    if close_col not in dataframe.columns or idx not in dataframe.index or shares <= 0:
        return 0.0, 0.0
    prev = dataframe[close_col].shift(1).loc[idx]
    cur = dataframe[close_col].loc[idx]
    if pd.isna(prev) or pd.isna(cur) or float(prev) == 0:
        return 0.0, 0.0
    pnl = (float(cur) - float(prev)) * float(shares)
    ret = (float(cur) / float(prev) - 1.0) * 100.0
    return pnl, ret


def _ptb_tqqq_risk_score():
    latest = data.iloc[-1]
    score = 0
    try:
        if not bool(latest.get("Above_MA200", False)):
            score += 3
        if not bool(latest.get("Above_MA20", False)):
            score += 1
        if not bool(latest.get("MA50_Slope_Up", False)):
            score += 1
        if float(latest.get("RSI14", 50)) < 50:
            score += 1
        if float(latest.get("CCI20", 0)) < 0:
            score += 1
        if float(latest.get("STOCH_K", 50)) < float(latest.get("STOCH_D", 50)):
            score += 1
    except Exception:
        pass
    return min(int(score), 8)


def create_power_tqqq_blue_vibe_dashboard(path=os.path.join(OUTPUT_DIR, "power_tqqq_blue_dashboard.png")):
    asset = "TQQQ"
    accent = "#38BDF8"
    accent2 = "#2563EB"
    bg = "#070C14"
    panel = "#0D1522"
    line = "#1B2A3F"
    muted = "#9BAEC9"
    white = "#EAF3FF"
    gold = "#93C5FD"
    red = "#FF5A66"
    W, H = 1160, 2480

    img = Image.new("RGB", (W, H), bg)
    draw = ImageDraw.Draw(img)
    _ptb_draw_gradient_rect(draw, (0, 0, W, H), "#070C14", "#0A111C")

    shares = float(HOLDINGS.get("TQQQ", 0.0) or 0.0)
    avg_price = float(AVG_PRICE.get("TQQQ", 0.0) or 0.0)
    live_price = float(prices_live.get("TQQQ", prices.get("TQQQ", 0.0)) or 0.0)
    base = shares * avg_price
    value = shares * live_price
    pnl = value - base if base > 0 else 0.0
    ret = pnl / base * 100 if base > 0 else 0.0
    risk_score = _ptb_tqqq_risk_score()
    buy_ready = bool(live_signal.get("action") == "BUY" and float(live_signal.get("mix", {}).get("TQQQ", 0.0)) > 0)
    sell_ready = bool(live_signal.get("action") in ["SELL_AND_NO_BUY", "DOWNSHIFT"])
    reason = str(live_signal.get("reason", signal.get("reason", "")) or "조건 확인 완료")

    draw.rounded_rectangle((45, 35, W-45, 125), radius=28, fill="#0D1828", outline=accent, width=2)
    draw.text((W//2, 80), "POWER TQQQ · POWER 공격형 계좌", font=get_font(38, True), fill=accent, anchor="mm")
    draw.text((W//2, 165), "누적 수익률 · 내 보유수량/평균단가 기준", font=get_font(27, True), fill=accent, anchor="mm")
    _ptb_draw_glow_text(draw, (W//2, 300), f"{ret:+.1f}%", get_font(112, True), accent if ret >= 0 else red, "mm", "#0B2A4A", 2)
    draw.text((W//2, 420), "순이익 · 내 계좌 평가 기준", font=get_font(26, True), fill=muted, anchor="mm")
    draw.text((W//2, 490), money(pnl), font=get_font(64, True), fill=accent if pnl >= 0 else red, anchor="mm")

    close_col = "TQQQ_Close"
    month_data = data[(data.index.year == today_date.year) & (data.index.month == today_date.month)].copy() if close_col in data.columns else pd.DataFrame()
    rets = [_ptb_account_daily_values(data, asset, shares, idx)[1] for idx in month_data.tail(21).index]
    avg_daily = sum(rets) / len(rets) if rets else 0.0
    y = 590
    for i, (label, val) in enumerate([("수익률", f"{ret:+.1f}%"), ("일평균", f"{avg_daily:+.2f}%"), ("위험점수", f"{risk_score}/8")]):
        x0 = 70 + i * 350
        draw.rounded_rectangle((x0, y, x0+310, y+132), radius=26, fill=panel, outline=line, width=2)
        draw.text((x0+155, y+48), val, font=get_font(35, True), fill=white if i == 2 else accent if ret >= 0 else red, anchor="mm")
        draw.text((x0+155, y+95), label, font=get_font(22, True), fill=muted, anchor="mm")

    draw.text((W//2, 790), "Verified · TQQQ Anchor Engine", font=get_font(23, True), fill=accent, anchor="mm")
    draw.text((W//2, 825), f"TQQQ {money(value)} · 내 계좌 기준 · rule based", font=get_font(22, True), fill=white, anchor="mm")
    draw.line((75, 865, W-75, 865), fill=line, width=2)

    draw.text((75, 925), "일별 성과 · 최근 5거래일 · 내 보유수량 기준", font=get_font(30, True), fill=muted, anchor="lm")
    recent = data[[close_col]].dropna().tail(5) if close_col in data.columns else pd.DataFrame()
    bar_vals = []
    for idx in recent.index:
        pnl_i, ret_i = _ptb_account_daily_values(data, asset, shares, idx)
        bar_vals.append((pd.Timestamp(idx).strftime("%m/%d"), pnl_i, ret_i))
    while len(bar_vals) < 5:
        bar_vals.insert(0, ("--", 0.0, 0.0))
    max_abs = max([abs(v[2]) for v in bar_vals] + [1.0])
    bar_base = 1160
    for i, (label_date, pnl_i, val) in enumerate(bar_vals[-5:]):
        x0 = 80 + i * 205
        bh = int(35 + min(130, abs(val) / max_abs * 120))
        c = accent if val >= 0 else red
        draw.rounded_rectangle((x0, bar_base-bh, x0+140, bar_base), radius=16, fill=c, outline=_ptb_blend_hex(c, "#FFFFFF", 0.25), width=1)
        draw.text((x0+70, bar_base-bh-50), money(pnl_i), font=get_font(19, True), fill=c, anchor="mm")
        draw.text((x0+70, bar_base-bh-24), f"{val:+.2f}%", font=get_font(20, True), fill=c, anchor="mm")
        draw.text((x0+70, bar_base+35), label_date, font=get_font(19, True), fill=muted, anchor="mm")

    y2 = 1240
    draw.rounded_rectangle((0, y2, W, y2+170), radius=0, fill="#0A111D", outline=line, width=2)
    draw.line((W//2, y2, W//2, y2+170), fill=line, width=2)
    draw.text((W//4, y2+62), "ON" if buy_ready else "OFF", font=get_font(58, True), fill=accent if buy_ready else muted, anchor="mm")
    draw.text((W//4, y2+120), "매수 게이트", font=get_font(23, True), fill=muted, anchor="mm")
    draw.text((W*3//4, y2+62), "ON" if sell_ready else "OFF", font=get_font(58, True), fill=red if sell_ready else gold, anchor="mm")
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
    week_no = 0
    last_week_key = None
    for d in range(1, last_day + 1):
        dts = pd.Timestamp(year=today_date.year, month=today_date.month, day=d)
        if dts.weekday() >= 5:
            continue
        week_key = (dts + pd.Timedelta(days=3-dts.weekday())).isocalendar().week
        if last_week_key is None:
            last_week_key = week_key
        elif week_key != last_week_key:
            week_no += 1
            last_week_key = week_key
        x0 = 70 + dts.weekday() * (cell_w + gap_x)
        y0 = cal_y + week_no * (cell_h + gap_y)
        matched = [idx for idx in close_series.index if pd.Timestamp(idx).date() == dts.date()]
        if matched:
            pnl_i, ret_i = _ptb_account_daily_values(data, asset, shares, matched[-1])
            if ret_i >= 0:
                fill = _ptb_blend_hex("#0D1B21", accent2, min(0.62, 0.12 + abs(ret_i)/8)); outline = accent; txt = accent
            else:
                fill = _ptb_blend_hex("#171018", "#7A1D25", min(0.62, 0.12 + abs(ret_i)/8)); outline = red; txt = "#FF8A8A"
            pnl_text = money(pnl_i); ret_text = f"{ret_i:+.2f}%"
        else:
            fill = "#0A111D"; outline = "#243249"; txt = "#50627C"; pnl_text = "휴장"; ret_text = ""
        draw.rounded_rectangle((x0, y0, x0+cell_w, y0+cell_h), radius=16, fill=fill, outline=outline, width=1)
        draw.text((x0+14, y0+22), str(d), font=get_font(19, True), fill=muted, anchor="lm")
        draw.text((x0+14, y0+58), pnl_text, font=get_font(19, True), fill=txt, anchor="lm")
        if ret_text:
            draw.text((x0+14, y0+84), ret_text, font=get_font(18, True), fill=txt, anchor="lm")

    draw.rounded_rectangle((45, H-150, W-45, H-45), radius=24, fill="#101B2B", outline=line, width=2)
    for i, ln in enumerate(_ptb_wrap_lines(reason, 57)[:2]):
        draw.text((75, H-110+i*34), ln, font=get_font(22, True), fill=white, anchor="lm")
    img.save(path, quality=95)
    return path


def create_live_check_card(path=os.path.join(OUTPUT_DIR, "live_execution_check_card.png")):
    W,H = 1080,1920
    img = Image.new("RGB", (W,H), CARD["bg_top"])
    draw_gradient(img, CARD["bg_top"], CARD["bg_bottom"])
    img = img.convert("RGBA")
    d = ImageDraw.Draw(img)
    d.text((60,55), "LIVE EXECUTION", font=get_font(30), fill=CARD["cyan"])
    d.text((60,105), "실시간 실행 판단", font=get_font(56), fill=CARD["white"])
    d.text((60,172), f"종가 {today_date.strftime('%Y-%m-%d')} · 확인 {live_checked_at}", font=get_font(25), fill=CARD["muted"])
    rounded(d, (725,70,1015,132), radius=31, fill=live_final["soft"])
    d.text((870,101), live_final["badge"], font=get_font(26), fill=live_final["color"], anchor="mm")

    card(img, (55,245,1025,455), radius=44, fill=CARD["white"]); d = ImageDraw.Draw(img)
    rounded(d, (95,290,310,343), radius=26, fill=live_final["soft"])
    d.text((202,316), "최종 실행", font=get_font(23), fill=live_final["color"], anchor="mm")
    d.text((95,392), live_final["title"], font=get_font(43), fill=live_final["color"])

    card(img, (55,500,1025,735), radius=44, fill=CARD["white"]); d = ImageDraw.Draw(img)
    d.text((95,558), "종가 판단 vs 실시간 판단", font=get_font(36), fill=CARD["purple"])
    kv(d,95,625,"종가 판단", f"{signal['action']} / {signal['market_type']}", value_color=CARD["blue"], size=24, value_x=960)
    kv(d,95,675,"실시간 판단", f"{live_signal['action']} / {live_signal['market_type']}", value_color=live_final["color"], size=24, value_x=960)

    card(img, (55,780,1025,1115), radius=44, fill=CARD["white"]); d = ImageDraw.Draw(img)
    d.text((95,838), "가격 괴리", font=get_font(36), fill=CARD["text"])
    y=900
    for ticker in ["QQQ", "QLD", "TQQQ"]:
        s = realtime_snapshot[ticker]
        c = CARD["buy"] if float(s["Diff_%"]) >= 0 else CARD["sell"]
        d.text((95,y), ticker, font=get_font(26), fill=CARD["subtext"])
        d.text((245,y), money(s["Close_Price"]), font=get_font(24), fill=CARD["text"])
        d.text((500,y), money(s["Live_Price"]), font=get_font(24), fill=CARD["blue"])
        d.text((950,y), f"{float(s['Diff_%']):+.2f}%", font=get_font(24), fill=c, anchor="ra")
        y += 58
    d.text((95,1080), "표시: 종가 / 현재가 / 괴리율", font=get_font(20), fill=CARD["subtext"])

    card(img, (55,1160,1025,1415), radius=44, fill=CARD["white"]); d = ImageDraw.Draw(img)
    d.text((95,1218), "최종 실행 문구", font=get_font(36), fill=live_final["color"])
    draw_text_box(d,95,1285,live_final["final"],get_font(25),CARD["text"],850,line_gap=10,max_lines=3)
    draw_text_box(d,95,1360,live_cooldown_status,get_font(22),CARD["wait"],850,line_gap=8,max_lines=2)

    card(img, (55,1460,1025,1725), radius=44, fill=CARD["white"]); d = ImageDraw.Draw(img)
    d.text((95,1518), "실시간 주문 기준", font=get_font(36), fill=CARD["text"])
    y=1585
    if not live_buy_df.empty:
        for _, row in live_buy_df.iterrows():
            is_sell = str(row.get("Order", "BUY")) == "SELL"
            row_soft = CARD["sell_soft"] if is_sell else CARD["buy_soft"]
            row_main = CARD["sell"] if is_sell else CARD["buy_dark"]
            amount = row.get("Sell_Amount", 0.0) if is_sell else row.get("Buy_Amount", 0.0)
            shares = row.get("Sell_Shares", 0.0) if is_sell else row.get("Estimated_Shares", 0.0)
            label = "매도" if is_sell else "매수"
            rounded(d,(95,y,965,y+82),radius=26,fill=row_soft)
            d.text((130,y+26),f"{label} {row['Ticker']}",font=get_font(25),fill=CARD["text"],anchor="lm")
            d.text((360,y+26),money(row["Price"]),font=get_font(22),fill=CARD["blue"],anchor="lm")
            d.text((555,y+26),money(amount),font=get_font(22),fill=row_main,anchor="lm")
            d.text((930,y+26),str(shares),font=get_font(18),fill=CARD["text"],anchor="rm")
            d.text((130,y+60),str(row.get("Note", "현재가 기준 예상 수량")),font=get_font(17),fill=CARD["subtext"],anchor="lm")
            y += 92
    else:
        rounded(d,(95,y,965,y+92),radius=26,fill=CARD["wait_soft"])
        d.text((130,y+46),"주문 없음",font=get_font(30),fill=CARD["wait"],anchor="lm")
        d.text((930,y+46),"WAIT",font=get_font(30),fill=CARD["wait"],anchor="rm")

    draw_yellow_footer(d,W,1780,line1="종가 판단은 백테스트 기준 그대로",line2="실시간 판단은 현재가 스냅샷으로 주문 직전 확인")
    img = img.convert("RGB")
    img.save(path, quality=95)
    return path

today_card_path = create_today_action_card()
live_card_path = create_live_check_card()
analysis_card_path = create_analysis_card()
strategy_card_path = create_strategy_card()
power_tqqq_blue_path = create_power_tqqq_blue_vibe_dashboard()

for p in [today_card_path, live_card_path, analysis_card_path, strategy_card_path, power_tqqq_blue_path, qqq_price_graph_path, drawdown_graph_path, buy_order_graph_path]:
    print("이미지 생성:", p)

# ======================================================
# 9. 텔레그램 메시지 / 발송
# ======================================================

info = action_info()
telegram_text = f"""
[QQQ / QLD / TQQQ 전략 알림]

날짜: {today_date.strftime('%Y-%m-%d')}
전략: {STRATEGY_NAME}

오늘 결론:
{info['title']} {info['verb']}

시장 상태:
{signal['market_type']}

판단 사유:
{signal['reason']}

가격:
QQQ 종가 {prices['QQQ']:.2f} / 현재가 {prices_live['QQQ']:.2f} / 괴리 {realtime_snapshot['QQQ']['Diff_%']:+.2f}%
QLD 종가 {prices['QLD']:.2f} / 현재가 {prices_live['QLD']:.2f} / 괴리 {realtime_snapshot['QLD']['Diff_%']:+.2f}%
TQQQ 종가 {prices['TQQQ']:.2f} / 현재가 {prices_live['TQQQ']:.2f} / 괴리 {realtime_snapshot['TQQQ']['Diff_%']:+.2f}%

실시간 실행 판단:
확인시각: {live_checked_at}
종가판단: {signal['action']} / {signal['market_type']}
실시간판단: {live_signal['action']} / {live_signal['market_type']}
최종상태: {live_final['status']}
최종문구: {live_final['final']}

주요 지표:
QQQ 20일선 {today['QQQ_MA20']:.2f}
QQQ 50일선 {today['QQQ_MA50']:.2f}
QQQ 100일선 {today['QQQ_MA100']:.2f}
QQQ 200일선 {today['QQQ_MA200']:.2f}
QQQ 200일선 위: {bool(today['Above_MA200'])}
MA50 상승중: {bool(today['MA50_Slope_Up'])}
1년 고점 대비 낙폭: {today['Drawdown_From_1Y_High'] * 100:.2f}%

보조지표:
RSI14 {today['RSI14']:.2f}
Stoch K/D {today['STOCH_K']:.2f} / {today['STOCH_D']:.2f}
CCI20 {today['CCI20']:.2f}
ADX14 {today['ADX14']:.2f}
OBV 추세 {'UP' if bool(today['OBV_Trend_Up']) else 'DOWN'}

계좌:
현금 {CASH:,.2f}
보유평가금액 {position_value:,.2f}
보유원금 {position_cost:,.2f}
포지션 수익률 {position_return * 100:.2f}%
총자산 {total_equity:,.2f}

보유조합: {holding_combo}
익절목표: {'없음' if target_return is None else f'{target_return * 100:.2f}%'}
현재 쿨다운 계산값: {COOLDOWN_DAYS_LEFT}일
쿨다운 계산근거: {COOLDOWN_AUTO_NOTE}
쿨다운 상태: {cooldown_status}
매수 기준: {buy_plan_text(signal)}
실시간 쿨다운 상태: {live_cooldown_status}
실시간 매수 기준: {buy_plan_text(live_signal)}
익절매도 발생 시 다음 쿨다운: {SELL_COOLDOWN_DAYS_AFTER_PROFIT}거래일

오늘 주문:
"""

if signal["action"] == "SELL_AND_NO_BUY" and not buy_df.empty:
    telegram_text += "\n전량매도하세요. 아래 수량을 각각 매도하세요.\n"
    for _, row in buy_df.iterrows():
        telegram_text += f"\n- [매도] {row['Ticker']}\n  기준가: {row['Price']:,.2f}\n  매도수량: {row['Sell_Shares']}주\n  예상매도금액: {row['Sell_Amount']:,.2f}\n  메모: {row.get('Note','')}\n"
    telegram_text += f"\n오늘은 재매수하지 않습니다.\n다음 거래일부터 쿨다운 {SELL_COOLDOWN_DAYS_AFTER_PROFIT}거래일 적용하세요.\n"
elif signal["action"] in ["BUY", "DOWNSHIFT"] and not buy_df.empty:
    if signal["action"] == "BUY":
        telegram_text += f"\n매수 기준: {buy_plan_text(signal)}\n"
    for _, row in buy_df.iterrows():
        if str(row.get("Order", "BUY")) == "SELL":
            telegram_text += f"\n- [매도] {row['Ticker']}\n  기준가: {row['Price']:,.2f}\n  매도수량: {row['Sell_Shares']}주\n  예상매도금액: {row['Sell_Amount']:,.2f}\n  메모: {row.get('Note','')}\n"
        else:
            telegram_text += f"\n- [매수] {row['Ticker']}\n  기준가: {row['Price']:,.2f}\n  매수가: {row['Buy_Amount']:,.2f}\n  예상수량: {row['Estimated_Shares']}주\n  비중: {float(row.get('Weight',0))*100:.2f}%\n  메모: {row.get('Note','')}\n"
else:
    telegram_text += f"\n오늘은 쉬세요.\n매수 주문 없음.\n쿨다운 상태: {cooldown_status}\n"

print("="*100)
print("텔레그램 발송")
print("="*100)

send_telegram_message(telegram_text)
send_telegram_photo(today_card_path, caption="오늘 매수·매도 판단")
send_telegram_photo(live_card_path, caption="실시간 실행 판단")
send_telegram_photo(analysis_card_path, caption="오늘 시장·계좌 분석")
send_telegram_photo(power_tqqq_blue_path, caption="POWER TQQQ BLUE · 공격형 계좌")
send_telegram_photo(qqq_price_graph_path, caption="QQQ 가격 + 이동평균선 + 볼린저밴드")
send_telegram_photo(drawdown_graph_path, caption="QQQ 낙폭 그래프")
send_telegram_photo(buy_order_graph_path, caption="오늘 매수 주문 그래프")
send_telegram_photo(strategy_card_path, caption="최종 전략 요약")

print("="*100)
print("생성된 파일")
print("="*100)
print("오늘 매수·매도 판단 카드:", today_card_path)
print("실시간 실행 판단 카드:", live_card_path)
print("오늘 분석 카드:", analysis_card_path)
print("POWER TQQQ BLUE 카드:", power_tqqq_blue_path)
print("QQQ 가격 그래프:", qqq_price_graph_path)
print("낙폭 그래프:", drawdown_graph_path)
print("매수 주문 그래프:", buy_order_graph_path)
print("전략 요약 카드:", strategy_card_path)
print("요약 CSV:", summary_csv_path)
print("종가 기준 주문 CSV:", order_csv_path)
print("실시간 기준 주문 CSV:", live_order_csv_path)
