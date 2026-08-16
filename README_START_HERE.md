# V14 99/1 GitHub 자동실행 세팅 가이드

이 패키지는 기존 ANCHOR 파일과 충돌하지 않도록 V14 전용 파일명으로 만들었습니다.

## 1. GitHub에 올릴 파일

압축을 풀면 아래 파일들이 있습니다. 이 파일들을 GitHub repo의 루트에 그대로 올리세요.

```text
v14_live_regular.py
v14_alert_monitor.py
v14_account_config.env
requirements.txt
.gitignore
.github/workflows/v14_regular.yml
.github/workflows/v14_alert_monitor.yml
```

기존 ANCHOR의 `account_config.env`, REGULAR workflow, ALERT workflow는 건드리지 않아도 됩니다.
V14는 `v14_account_config.env`만 사용합니다.

---

## 2. V14 보유량 입력 위치

`v14_account_config.env` 파일을 열어서 본인 V14 계좌 기준으로 수정하세요.

```env
V14_CASH=10000

V14_QQQ_SHARES=0
V14_QQQ_AVG_PRICE=0

V14_QLD_SHARES=5
V14_QLD_AVG_PRICE=90

V14_TQQQ_SHARES=0
V14_TQQQ_AVG_PRICE=0

# 마지막 익절 전량매도일
# 형식: YYYY-MM-DD
# 예시: V14_LAST_PROFIT_SELL_DATE=2026-07-08
# 평소에는 비워두세요.
V14_LAST_PROFIT_SELL_DATE=

V14_PERIOD=15y
```

중요: `V14_LAST_PROFIT_SELL_DATE`는 쿨다운 남은 일수를 직접 쓰는 곳이 아닙니다.
익절 전량매도를 실제로 한 날만 날짜를 넣으면 프로그램이 NYSE 거래일 기준으로 자동 계산합니다.

예:

```env
V14_LAST_PROFIT_SELL_DATE=2026-07-08
```

---

## 3. 텔레그램 Secrets 찾는 곳

GitHub repo 화면에서:

```text
Settings
→ Secrets and variables
→ Actions
→ Repository secrets
```

여기에 아래 2개가 있어야 합니다.

```text
TELEGRAM_BOT_TOKEN
TELEGRAM_CHAT_ID
```

이미 ANCHOR에서 쓰고 있다면 V14도 같은 이름을 그대로 사용합니다.
새로 만들 필요 없습니다.

없으면 `New repository secret` 버튼으로 추가하세요.

---

## 4. REGULAR와 ALERT 차이

### REGULAR: `v14_regular.yml`

장마감 후 하루 1번 전체 판단을 보냅니다.

포함 내용:

```text
BUY / SELL / DOWNSHIFT / WAIT
쿨다운 상태
매수 분할
QQQ/QLD 비율
매수수량/매도수량
CSV/PNG 결과 파일
텔레그램 전체 리포트
```

### ALERT: `v14_alert_monitor.yml`

장중 10분마다 긴급조건만 봅니다.

보내는 조건은 2개뿐입니다.

```text
1. 익절 전량매도 조건 발생
2. 다운시프트 조건 발생
```

V14는 매수신호가 자주 나올 수 있으므로 장중 매수후보는 알림 보내지 않습니다.
조건 없으면 아무 메시지도 안 보냅니다.

---

## 5. GitHub에 파일 올리는 순서

1. GitHub repo에 들어갑니다.
2. `Add file` → `Upload files`를 누릅니다.
3. 압축 푼 파일 전체를 드래그해서 올립니다.
4. 아래 파일들이 보이는지 확인합니다.

```text
v14_live_regular.py
v14_alert_monitor.py
v14_account_config.env
requirements.txt
.gitignore
.github/workflows/v14_regular.yml
.github/workflows/v14_alert_monitor.yml
```

5. 초록색 `Commit changes` 버튼을 누릅니다.

---

## 6. 처음 테스트하는 방법

GitHub repo에서:

```text
Actions
→ V14 Regular Dashboard
→ Run workflow
```

을 눌러 먼저 수동 실행하세요.

그다음:

```text
Actions
→ V14 Alert Monitor
→ Run workflow
```

도 한 번 눌러보세요.

ALERT는 긴급조건이 없으면 텔레그램이 안 오는 게 정상입니다.
GitHub Actions 로그에 `긴급조건 없음: 무음 종료`가 보이면 정상입니다.

---

## 7. 결과 파일 보는 법

REGULAR 실행이 끝나면 Actions 실행 화면 하단에 artifact가 생깁니다.

```text
v14-regular-results
```

이 안에 CSV/PNG 결과가 들어 있습니다.

ALERT는 긴급조건이 없으면 결과 파일이 거의 없고, 텔레그램도 안 옵니다.

---

## 8. 기존 ANCHOR와 충돌하지 않는 이유

ANCHOR가 쓰는 파일:

```text
account_config.env
기존 anchor workflow
```

V14가 쓰는 파일:

```text
v14_account_config.env
v14_regular.yml
v14_alert_monitor.yml
v14_live_regular.py
v14_alert_monitor.py
```

파일명이 전부 다르기 때문에 서로 덮어쓰지 않습니다.

---

## 9. 공개 repo 주의

`v14_account_config.env`에는 보유수량, 평단, 현금이 들어갑니다.
repo가 Public이면 이 값도 남에게 보입니다.
그게 싫으면 repo를 Private으로 바꾸거나 보유량도 GitHub Secrets 방식으로 바꿔야 합니다.

텔레그램 토큰은 절대 파일에 쓰지 말고 GitHub Secrets에만 넣으세요.
