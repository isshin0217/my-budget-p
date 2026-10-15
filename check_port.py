"""
check_port.py
Apps Script(Code.gs)로 이식한 로직이 파이썬 원본과 같은 결과를 내는지 대조한다.

파싱 규칙을 고칠 때 한쪽만 고치면 폰과 엑셀의 숫자가 갈린다.
규칙을 건드렸으면 이걸 돌려서 양쪽이 여전히 일치하는지 확인할 것.

    python check_port.py
"""
import re
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
from parser import _APPROVAL_NO_NUM_RE, _APPROVAL_RE, CATEGORY_RULES, PERIOD_BY_CATEGORY
import report

src = (ROOT / "apps_script" / "Code.gs").read_text(encoding="utf-8")
fail = 0


def check(label, ok, detail=""):
    global fail
    print(f"  {'OK  ' if ok else 'FAIL'}  {label}{'' if ok else '  ' + detail}")
    if not ok:
        fail += 1


# ── 1. 정규식 ──────────────────────────────────────────────────────────────
print("정규식")


def js_regex(name):
    """Code.gs 의 new RegExp(...) 를 파이썬 패턴 문자열로 옮긴다.

    한계를 알고 쓸 것: 이건 JS 패턴을 파이썬 re 로 **다시 컴파일**해서 비교한다.
    즉 파이썬끼리 견주는 것이라 정규식 '엔진' 차이는 원리상 못 잡는다.
    실제로 V8 은 (?:...)? 안의 빈 반복을 거부하는데 파이썬 re 는 허용해서,
    결제유형 없는 문자의 월이 한 자리 잘려 나가는 걸 이 검사가 통과시켰다.
    엔진 차이까지 보려면 node 로 Code.gs 를 직접 돌려 출력을 대조해야 한다.
    """
    block = src.split(f"const {name} = new RegExp(")[1].split(");")[0]
    # 주석을 먼저 걷어낸다 — 주석에 따옴표가 들어가면 패턴 조각으로 잘못 읽힌다
    block = re.sub(r"//[^\n]*", "", block)
    js_parts = re.findall(r"'((?:[^'\\]|\\.)*)'", block)
    # JS 문자열 리터럴의 이스케이프만 풀고(\\ → \), 한글은 건드리지 않는다
    js = "".join(js_parts).replace("\\\\", "\\")
    # JS는 (?<name>...), 파이썬은 (?P<name>...) — 문법 차이만 맞춰 비교.
    # 다만 lookbehind 인 (?<! 와 (?<= 는 양쪽 문법이 같으므로 건드리면 안 된다.
    # 그냥 "(?<" 를 바꾸면 (?<!\d) 가 (?P<!\d) 가 되어 컴파일부터 실패한다.
    js = js.replace(r"\/", "/")
    return re.sub(r"\(\?<(?![!=])", "(?P<", js)


def compare_regex(label, js_name, py_re, samples):
    """JS 쪽 정규식이 파이썬과 같은 패턴이고 같은 캡처를 내는지."""
    js_norm = js_regex(js_name)
    try:
        r = re.compile(js_norm)
        check(f"{label}: JS 정규식이 컴파일된다", True)
    except re.error as e:
        check(f"{label}: JS 정규식이 컴파일된다", False, str(e))
        return
    check(f"{label}: 파이썬 패턴과 동일 (명명그룹 문법 차이 제외)", js_norm == py_re.pattern,
          f"\n     JS: {js_norm}\n     PY: {py_re.pattern}")
    misses = [s for s in samples if not (r.search(s) and py_re.search(s))]
    if misses:
        check(f"{label}: 샘플 {len(samples)}건이 모두 매칭된다", False, f"실패: {misses}")
        return
    same = all(r.search(s).groupdict() == py_re.search(s).groupdict() for s in samples)
    check(f"{label}: 샘플 {len(samples)}건 캡처 결과 동일", same)


samples = [
    "신한카드(1234)승인 홍*동 94,400원(일시불)07/01 19:55 폴리스 유한",
    "신한카드(1234)취소 홍*동 12,000원(일시불)07/03 21:40 배달의민족",
    "신한카드(1234)승인 홍*동 55,000원 07/21 10:00 SKT통신요금",
    # 우리카드 — 카드사 이름에 '카드'가 없고 결제유형이 괄호 밖에 온다
    "우리(5678)승인 김*수님 36,400원 일시불 07/20 21:54 유한회사 미소",
    "우리(5678)승인 김*수님 240,000원 3개월 07/22 15:00 무신사 스토어",
    # 취소 문자 — '매출취소' 이고 시각이 없다
    "신한카드 (1234) 매출취소 홍*동 29,902원 07/12 네이버페이",
]
compare_regex("카드번호 있는 형식", "APPROVAL_RE", _APPROVAL_RE, samples)

# 현대카드 — 카드번호 괄호가 아예 없다
samples_no_num = [
    "현대카드Z ontact 승인 홍*동 10,000원 일시불 07/27 17:10 모바일티머니선 누적206,250원",
    "현대카드M 승인 홍*동 45,000원 일시불 07/28 09:05 이마트 성수",
]
compare_regex("카드번호 없는 형식", "APPROVAL_NO_NUM_RE", _APPROVAL_NO_NUM_RE, samples_no_num)

# 외화 해외결제 — 금액이 원이 아니고 날짜 뒤에 온다 (2026-10-17 실제 문자)
from parser import _FOREIGN_RE
samples_foreign = [
    "현대 대한항공120 해외승인 홍*동님 10/17 16:40 JPY 2,000.00 UBER*TRIPHELP.UBER.COM",
    "현대 대한항공120 해외승인 취소 홍*동님 10/17 16:41 JPY 1,258.00 UBER*PENDING",
    "현대 대한항공120 해외승인 홍*동님 10/18 09:12 USD 12.5 STARBUCKS SEATTLE",
]
compare_regex("외화 해외결제 형식", "FOREIGN_RE", _FOREIGN_RE, samples_foreign)
# 원화 패턴이 외화 문자를 먼저 집으면 FOREIGN_RE 까지 가지 못한다
stolen = [s for s in samples_foreign if _APPROVAL_RE.search(s) or _APPROVAL_NO_NUM_RE.search(s)]
check("외화 문자는 원화 패턴에 안 걸린다", not stolen, f"걸림: {stolen}")
# 반대로 원화 문자가 외화 패턴에 걸리면 안 된다
grabbed = [s for s in samples + samples_no_num if _FOREIGN_RE.search(s)]
check("원화 문자는 외화 패턴에 안 걸린다", not grabbed, f"걸림: {grabbed}")

# 번호 있는 형식이 먼저 걸려야 한다 — 헐거운 쪽이 앞서면 캡처가 엉뚱해진다
from parser import parse_message
hyundai = parse_message(
    "[Web발신]\n현대카드Z ontact 승인\n홍*동\n10,000원 일시불\n07/27 17:10\n"
    "모바일티머니선\n누적206,250원")
check("현대카드 문자가 파싱된다", hyundai is not None)
if hyundai:
    check("가맹점에서 누적액 꼬리표가 떨어진다", hyundai["merchant"] == "모바일티머니선",
          f"실제: {hyundai['merchant']!r}")
    check("금액이 10000", hyundai["amount"] == 10000, f"실제: {hyundai['amount']}")
    check("카드번호가 없으면 빈 값", hyundai["last4"] == "", f"실제: {hyundai['last4']!r}")

# ── 1-b. 가맹점 꼬리표 제거 ────────────────────────────────────────────────
js_trailing = re.search(r"const TRAILING_RE = /(.+?)/;", src).group(1)
from parser import _TRAILING_RE, clean_merchant
check("꼬리표 정규식 동일", js_trailing == _TRAILING_RE.pattern,
      f"\n     JS: {js_trailing}\n     PY: {_TRAILING_RE.pattern}")

trails = [
    ("카카오T일반 누적732,232원", "카카오T일반"),
    ("스타벅스 강남 누적 1,234,567 원", "스타벅스 강남"),
    ("이마트 잔액50,000원", "이마트"),
    ("폴리스 유한", "폴리스 유한"),          # 꼬리표 없으면 그대로
    ("누적식당", "누적식당"),                # 가맹점명에 '누적'이 들어간 경우
]
bad = [(raw, clean_merchant(raw), want) for raw, want in trails if clean_merchant(raw) != want]
check(f"꼬리표 제거 {len(trails)}건", not bad, f"어긋남: {bad}")

# ── 2. 카테고리 / 주기 테이블 ──────────────────────────────────────────────
print("\n카테고리 · 주기 테이블")
js_cats = re.findall(r"^\s*\['([^']+)',\s*\[", src, re.M)
py_cats = [c for c, _ in CATEGORY_RULES]
check("카테고리 목록과 순서 동일", js_cats == py_cats, f"\n     JS: {js_cats}\n     PY: {py_cats}")

# PERIOD_BY_CATEGORY 블록만 떼어서 본다.
# (PERIOD_KEY 도 '한글' 값을 가지고 있어 파일 전체를 훑으면 섞인다)
period_block = src.split("const PERIOD_BY_CATEGORY = {")[1].split("};")[0]
js_period = dict(re.findall(r"'([^']+)':\s*'(일상|주기|고정)'", period_block))
check("주기 매핑 동일", js_period == PERIOD_BY_CATEGORY,
      f"\n     JS만: {set(js_period) - set(PERIOD_BY_CATEGORY)}"
      f"\n     PY만: {set(PERIOD_BY_CATEGORY) - set(js_period)}")

# ── 영어 표시 라벨 ─────────────────────────────────────────────────────────
print("\n표시 라벨 (Code.gs 전용)")
label_block = src.split("const CATEGORY_LABEL = {")[1].split("};")[0]
js_labels = dict(re.findall(r"'([^']+)':\s*'([^']+)'", label_block))

py_cat_names = set(c for c, _ in CATEGORY_RULES) | {"미분류"}
check("모든 카테고리에 영어 라벨이 있다", py_cat_names <= set(js_labels),
      f"라벨 없음: {sorted(py_cat_names - set(js_labels))}")
check("라벨에 군더더기가 없다", set(js_labels) <= py_cat_names,
      f"규칙에 없는 라벨: {sorted(set(js_labels) - py_cat_names)}")
check("영어 라벨이 서로 겹치지 않는다", len(set(js_labels.values())) == len(js_labels))

# 주기(고정/주기/일상)를 시트에 적는 Type 열은 없앴다 — 앱 어디에서도 읽지 않는
# 칸이었다. 그래서 PERIOD_LABEL / PERIOD_KEY 대조도 함께 뺀다.
# PERIOD_BY_CATEGORY 표 자체는 위에서 계속 대조한다: parseMessage_ 가 parse_message 의
# 이식본이라 같은 필드를 내놓아야 하고, 파이썬 월간요약의 '주기별' 집계가 그 값을 쓴다.
check("Ledger 에 Type 열이 없다", "'Type'" not in src.split("const HEADERS = [")[1].split("]")[0])

# 카테고리별 키워드 집합을 정확히 대조한다.
# 같은 키워드가 다른 카테고리에 들어가 있으면 분류 결과가 갈리므로 이게 진짜 검사다.
js_rules = {}
for cat, body in re.findall(r"^\s*\['([^']+)',\s*\[(.*?)\]\],\s*$", src, re.M | re.S):
    js_rules[cat] = set(re.findall(r"'([^']*)'", body))
py_rules = {cat: set(kw) for cat, kw in CATEGORY_RULES}

diffs = []
for cat in py_rules:
    only_py = py_rules[cat] - js_rules.get(cat, set())
    only_js = js_rules.get(cat, set()) - py_rules[cat]
    if only_py:
        diffs.append(f"{cat}: 파이썬에만 {sorted(only_py)}")
    if only_js:
        diffs.append(f"{cat}: Code.gs에만 {sorted(only_js)}")
check("카테고리별 키워드 집합 동일", not diffs, "\n     " + "\n     ".join(diffs))

# 같은 가맹점명이 양쪽에서 같은 카테고리로 분류되는지 (순서까지 반영한 최종 확인)
def js_classify(merchant):
    m = merchant.lower()
    for cat, _ in CATEGORY_RULES:                 # 순서는 위에서 동일함을 확인함
        if any(k.lower() in m for k in js_rules.get(cat, set())):
            return cat
    return "미분류"

from parser import classify
probes = ["무신사 스토어", "스타벅스 강남", "이자카야 목로", "SKT통신요금", "넷플릭스",
          "올리브영 신촌", "다이소 홍대", "쿠팡", "GS25 역삼", "카카오T", "폴리스 유한"]
bad = [(p, classify(p), js_classify(p)) for p in probes if classify(p) != js_classify(p)]
check(f"가맹점 {len(probes)}건 분류 결과 동일", not bad, f"불일치: {bad}")


# ── 3. 주 시작일 계산 ──────────────────────────────────────────────────────
print("\n주 계산")


def js_week_start(d, start_index):
    """Code.gs weekStartOf_ 를 그대로 옮긴 것 (getDay: 일=0)."""
    js_get_day = (d.weekday() + 1) % 7          # 파이썬 월=0 → JS 일=0
    monday_based = (js_get_day + 6) % 7
    back = (monday_based - start_index + 7) % 7
    return d - timedelta(days=back)


mismatch = []
for offset in range(400):
    d = date(2026, 1, 1) + timedelta(days=offset)
    for start_index in range(7):
        py = report.week_start_of(datetime(d.year, d.month, d.day, 12, 0), start_index)
        if js_week_start(d, start_index) != py:
            mismatch.append((d, start_index))
check("400일 × 시작요일 7가지 = 2800건 주 시작일 일치", not mismatch,
      f"불일치 {len(mismatch)}건 예: {mismatch[:3]}")


def js_days_left(now, end):
    import math
    today_midnight = date(now.year, now.month, now.day)
    return max(math.ceil((end - today_midnight).days) + 1, 0)


mismatch = []
for offset in range(400):
    now = datetime(2026, 1, 1, 15, 30) + timedelta(days=offset)
    start = report.week_start_of(now, 0)
    end = start + timedelta(days=6)
    py = max((end - now.date()).days + 1, 0)
    if js_days_left(now, end) != py:
        mismatch.append(now.date())
check("남은 일수 400일치 일치", not mismatch, f"불일치: {mismatch[:3]}")

# ── 4. 상수 ────────────────────────────────────────────────────────────────
print("\n상수")
check("WEEKS_PER_MONTH = 4.35", "const WEEKS_PER_MONTH = 4.35" in src)
check("고정비 감지 26~35일", "gap >= 26 && gap <= 35" in src)
check("고정비 감지 15% 이내", "<= 0.15" in src)
# 토큰이 곧 웹앱의 비밀번호다. 예시값이 남아 있거나 비어 있으면 사실상 공개 상태다.
_token = re.search(r"const\s+TOKEN\s*=\s*'([^']*)'", src)
check("TOKEN 이 비어 있지 않다", bool(_token and _token.group(1)))
check("TOKEN 이 기본 예시값이 아니다", "CHANGE-ME" not in src)

print("\n" + ("모두 통과" if not fail else f"실패 {fail}건"))
sys.exit(1 if fail else 0)
