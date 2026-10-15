/**
 * Code.gs — 카드 문자 가계부 (Google Apps Script 웹앱)
 *
 * 아이폰 단축어가 카드 문자를 POST하면 파싱해서 구글시트에 기록하고,
 * '이번 주에 얼마 더 써도 되는지'를 계산해 돌려준다.
 *
 *   POST ?t=토큰   문자 1건 저장 → 요약 문자열 응답 (단축어 알림용)
 *   GET  ?t=토큰   폰에서 보는 대시보드 (홈 화면에 추가해서 사용)
 *
 * 계산 모델:
 *   계획 한도(planAllow) = (월 예산 − 월 고정비) ÷ 4.35
 *   실제 한도(weeklyAllow) = 계획 한도 − 통장이 못 대는 몫
 *   이번 주 남은 돈       = 실제 한도 − 이번 주 변동비 − 예약분
 *   이번 달 카드값        = 이번 달 카드 승인 총액 (고정비 포함)
 *
 * 예산은 주 단위로 보고, 카드값은 한 달에 한 번 빠져나간다. 주기가 서로 달라서
 * 둘을 따로 센다 — 주간 페이스는 '이번 주', 결제일에 나갈 돈은 '이번 달'.
 *
 * 예산은 계획이고 통장 잔액은 현실이다. 계획만 보면 통장이 마이너스인데도 "이번 주
 * ₩310,000 쓸 수 있다"고 말하게 되므로, 잔액 예측(forecast_)의 최저점이 바닥을
 * 뚫는 만큼 한도를 깎는다. 잔액을 안 적어두면(0) 예전처럼 계획 한도가 곧 실제 한도다.
 *
 * 파이썬 report.py 는 계획 한도까지만 같다 — 잔액·예측은 이쪽에만 있다.
 * check_port.py 로 대조할 때 그 차이를 감안할 것.
 *
 * 설치는 설치안내.md 참고.
 */

// ── 반드시 바꿀 것 ──────────────────────────────────────────────
// 웹앱 URL은 '링크가 있는 모든 사용자'에게 열려 있으므로(appsscript.json 의
// ANYONE_ANONYMOUS) 토큰으로 막는다. 이 값이 곧 가계부의 비밀번호다.
// 길고 아무 의미 없는 문자열로 바꾸세요.
//
// 주의: 이 값은 코드에 있으므로 깃 이력에 그대로 남는다. 나중에 여기서 지워도
// 옛 커밋에서 읽힌다. 이 저장소를 공개 저장소로 올릴 거면 그 전에 값을 바꾸고,
// 단축어와 대시보드 링크의 ?t= 도 같이 고칠 것.
const TOKEN = 'CHANGE-ME';  // 설치안내.md 대로 길고 아무 의미 없는 문자열로 바꿀 것

// 시트와 대시보드에 보이는 글자는 영어. 코드 내부 키는 한글 그대로 둔다.
// (분류 규칙이 한글 가맹점명을 매칭해야 하고, 파이썬 쪽과 대조 검증도 해야 하므로)
const SHEET = 'Ledger';
const SETTINGS = 'Settings';
const FIXED_SHEET = 'Fixed';     // 매달 나가는 돈 (월세·보험처럼 카드를 안 거치는 것 포함)
const EXTRA_SHEET = 'Extra';     // 환급·캐시백 등 일회성 수입 / 카드 밖 지출
const MONTHLY_SHEET = 'Monthly'; // 달마다 한 줄. 지나간 달을 돌아보는 기록

// 지난달 값(고정비·월급)은 그때의 설정이라 나중에 다시 계산할 수 없다.
// 그래서 그 달이 '현재'일 때 찍어두고, 지나간 뒤에는 덮어쓰지 않는다.
const MONTHLY_HEADERS = ['Month', 'Card', 'Off-card', 'Total out', 'Income', 'Saved',
                         'Budget', 'Top categories'];

// 'Day', 'Match' 는 나중에 덧붙인 열이다. 기존 시트를 깨지 않으려고 맨 뒤에 둔다.
//   Name  : 사람이 알아보는 이름 (대시보드에 이게 나온다)
//   Match : 카드 문자에 찍히는 가맹점 문자열. 비워두면 Name 으로 맞춘다.
const FIXED_HEADERS = ['Name', 'Amount', 'Every', 'Paid by', 'Note', 'Day', 'Match'];
const EXTRA_HEADERS = ['Date', 'Name', 'Amount', 'Note'];

// 사람이 적는 Tag 를 맨 끝에 둔다 — 노란 칸이 표 한가운데를 가르지 않도록.
// 열은 이름으로 찾으므로(headerIndex_) 시트에서 순서를 바꿔도 깨지지 않는다.
// MerchantEN 은 저장할 때 한 번 번역해 담아두는 숨은 칸이다.
// 화면을 열 때마다 번역하면 느려서, 미리 넣어두고 읽기만 한다.
// AutoCategory 는 '저장하던 그때 규칙이 뭐라고 했는지'를 적어두는 숨은 칸이다.
// 이게 없으면 규칙을 고쳤을 때 과거 행이 통째로 '사람이 고친 것'으로 둔갑한다(learnedCategory_).
// Foreign 은 외화 결제의 근거(원래 금액·환율·수수료)다. Amount 가 추정치라 청구서가
// 나오면 손으로 고치게 되는데, 그때 무엇을 고치는지 보이도록 숨기지 않는다.
const HEADERS = ['Date', 'Amount', 'Category', 'Merchant',
                 'Card', 'Last4', 'Payment', 'Kind', 'MerchantEN', 'AutoCategory', 'Tag',
                 'Foreign'];
// 사람에게 보일 일이 없는 칸. 새로 붙일 때 접어둔다.
const HIDDEN_COLUMNS = {'MerchantEN': true, 'AutoCategory': true};

// 이제 안 쓰는 열. 'Clean up ledger' 가 이미 있는 시트에서 지워준다.
//
// Type(Daily/Weekly/Fixed)은 카테고리에서 자동으로 따라오는 값이었는데, 정작 이 앱의
// 계산·표시 어디에서도 읽지 않았다. 저장될 때 채워지고 아무도 안 보는 칸이었다.
// 파이썬 쪽 월간요약의 '주기별' 집계는 엑셀의 제 열을 따로 쓰므로 영향이 없다.
const RETIRED_COLUMNS = ['Type'];
const WEEKS_PER_MONTH = 4.35;   // 365 ÷ 7 ÷ 12
// 주기 길이가 7일이 아닐 수도 있으므로 '한 달 = 며칠'로 환산해 쓴다.
// 7 / 30.45 = 1 / 4.35 이라 7일 주기에서는 예전 숫자가 그대로 나온다.
const DAYS_PER_MONTH = 7 * WEEKS_PER_MONTH;
const TZ = 'Asia/Seoul';

const BUDGET_KEY = 'Monthly budget';
const WEEKSTART_KEY = 'Week starts on';
const CYCLE_DAYS_KEY = 'Cycle length (days)';
const CYCLE_START_KEY = 'Cycle starts on';
const CYCLE_BUDGET_KEY = 'Cycle budget';
const FIXED_KEY = 'Fixed costs (manual)';
const BALANCE_KEY = 'Account balance';
const BALANCE_DATE_KEY = 'Balance as of';
const OVERDRAFT_KEY = 'Overdraft limit';
const INCOME_KEY = 'Monthly income';
const PAYDAY_KEY = 'Payday';
const CARD_DAY_KEY = 'Card payment day';
const OVERSEAS_FEE_KEY = 'Overseas fee %';
const FX_FALLBACK_KEY = 'FX fallback';

// Note 칸은 시트를 여는 사람이 읽는 설명서다. 값만 적어두면 몇 달 뒤에 본인도
// 뭘 넣었는지 헷갈리므로, 무엇을 넣는지 + 예시 + 비웠을 때 어떻게 되는지까지 적는다.
const DEFAULT_SETTINGS = [
  [BUDGET_KEY, 2000000,
   'REQUIRED. Everything you may spend in one month, fixed costs included. Example: 2000000'],
  [CYCLE_DAYS_KEY, 'Month',
   'How big a chunk you watch your spending in. Month = calendar month, 7 = a week at a time'],
  [CYCLE_START_KEY, '',
   'Only needed when the length above is NOT 7. Any past date one cycle began. Example: 2026-07-01'],
  [CYCLE_BUDGET_KEY, 0,
   'Fill ONLY if you already know what you may spend per cycle. Overrides the maths above. 0 = off'],
  [WEEKSTART_KEY, 'Mon',
   'Mon Tue Wed Thu Fri Sat Sun - the day your week resets. Ignored if "Cycle starts on" is set'],
  [FIXED_KEY, 0,
   'Rough total of monthly bills. Used ONLY while the Fixed sheet is empty. 0 = use Fixed sheet instead'],
  [BALANCE_KEY, 0,
   'What is in your bank account right now - typed by hand, nothing is connected. ' +
   'Caps what you may spend when the forecast says the money is not there. 0 = off'],
  [BALANCE_DATE_KEY, '',
   'The day you typed the balance above. Older dates are wound forward. Blank = today. Example: 2026-07-28'],
  [OVERDRAFT_KEY, 0,
   'Overdraft you may dip into below zero, as a positive number. Spending is only capped ' +
   'past this. 0 = none. Example: 5000000'],
  [INCOME_KEY, 0,
   'Salary that lands each month. Used for the balance forecast only, never for the budget. Example: 3000000'],
  [PAYDAY_KEY, 25,
   'Day of month the salary lands. 25 = the 25th. In short months it falls back to the last day'],
  [CARD_DAY_KEY, 0,
   'Day of month the card bill leaves your account. 0 = off. ' +
   'Card spending from that same month is billed. Example: 31'],
  [OVERSEAS_FEE_KEY, 1.3,
   'Added on top of the exchange rate for foreign-currency card payments (brand + overseas fee). ' +
   'Check your card bill. Example: 1.3'],
  [FX_FALLBACK_KEY, '',
   'Used only when the live exchange rate is unavailable. KRW per 1 unit. Example: JPY=9.1, USD=1380'],
];

// 카드 문자에 찍히는 카드사 이름 → 화면에 쓸 짧은 이름.
// '현대카드Z ontact' 처럼 뒤에 상품명이 붙어 오므로 포함 여부로 찾는다.
const CARD_BRANDS = [
  ['신한', 'Shinhan'], ['현대', 'Hyundai'], ['삼성', 'Samsung'], ['국민', 'KB'], ['KB', 'KB'],
  ['롯데', 'Lotte'], ['우리', 'Woori'], ['하나', 'Hana'], ['농협', 'NH'], ['BC', 'BC'],
  ['씨티', 'Citi'], ['카카오', 'Kakao'], ['토스', 'Toss'],
];

const FORECAST_DAYS = 35;   // 한 달 + 여유. 다음 급여일까지 반드시 포함되도록.

// 월=0 … 일=6. 설정 시트에는 영어로 적지만 한글 입력도 받아준다.
const WEEKDAY_EN = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
const WEEKDAYS_KO = '월화수목금토일';


// ══════════════════════════════════════════════════════ 파서 (parser.py 이식)

// 순서가 곧 우선순위다. 먼저 걸리는 카테고리가 이긴다.
// '배달'은 '식비'보다, '택시'는 '교통'보다 반드시 앞에 와야 한다.
// 뒤에 두면 배민이 식비로, 택시가 교통으로 먼저 걸려 새 카테고리가 영영 안 쓰인다.
const CATEGORY_RULES = [
  ['배달', ['배달', '배민', '우아한형제들', '요기요', '쿠팡이츠', 'uber eats', 'ubereats']],
  // 카카오T는 사실상 택시다. '카카오모빌리티'로 찍히는 건 주차·대리도 섞여 '교통'에 둔다.
  // 'uber' 는 Uber Eats 가 위 '배달'에 먼저 걸리므로 여기서는 승차만 남는다.
  ['택시', ['택시', '카카오t', 'uber']],
  ['식비', ['김밥', '식당', '분식', '국밥', '고기', '치킨', '피자', '버거', '맥도날드',
           '롯데리아', 'food', '레스토랑',
           '이자카야', '호프', '포차', '주점', '술집', '곱창', '막창', '초밥', '스시',
           '파스타', '칼국수', '냉면', '떡볶이', '베이커리', '제과', '빵']],
  // '매머드'는 문자에 '커피' 없이 상호만 찍혀 와서 따로 넣어야 잡힌다.
  // 표기가 '매머드'/'메머드' 둘 다 돌아다녀서 양쪽을 넣는다.
  ['카페', ['스타벅스', '카페', '커피', '이디야', '투썸', '메가', '매머드', '메머드',
           '커피빈', 'coffee', '빽다방']],
  ['마트/편의점', ['이마트', '홈플러스', '롯데마트', 'gs25', 'cu', '세븐일레븐', 'emart',
                '코스트코', '노브랜드', '편의점', '마트']],
  ['교통', ['카카오모빌리티', '코레일', 'ktx', '지하철', '버스', '주유',
          's-oil', 'gs칼텍스', '하이패스', '주차']],
  ['온라인쇼핑', ['쿠팡', '네이버페이', '11번가', '지마켓', '옥션', 'aliexpress', 'amazon']],
  ['통신', ['skt', 'kt', 'lg유플러스', 'lgu', '통신', '아이즈비전']],
  ['의료', ['병원', '약국', '의원', '치과', '한의원', 'clinic']],
  ['구독/디지털', ['netflix', '넷플릭스', 'youtube', '유튜브', 'google', 'apple.com', 'icloud',
                'spotify', '멜론', '지니']],
  ['의류/미용', ['무신사', '올리브영', '미용실', '헤어', '네일', '지그재그', '에이블리', '유니클로',
               'zara', 'h&m', '스파오', '아디다스', '나이키', '화장품', '이니스프리']],
  ['생필품', ['다이소', '생활용품', '문구', '세탁', '철물', '홈센터']],
  ['경조사/선물', ['화환', '꽃', '플라워', '선물', '기프트', '축의', '조의']],
  ['주거/고정', ['관리비', '월세', '임대', '전기', '가스', '수도', '보험', '헬스', '피트니스', '학원']],
];

const PERIOD_BY_CATEGORY = {
  '식비': '일상', '카페': '일상', '교통': '일상', '배달': '일상', '택시': '일상',
  '마트/편의점': '주기', '온라인쇼핑': '주기', '의류/미용': '주기',
  '생필품': '주기', '의료': '주기', '경조사/선물': '주기', '미분류': '주기',
  '통신': '고정', '구독/디지털': '고정', '주거/고정': '고정',
};

// 시트에 찍을 영어 라벨. 규칙의 키는 한글이라 파이썬 쪽과 그대로 대조된다.
const CATEGORY_LABEL = {
  '식비': 'Food', '카페': 'Cafe', '마트/편의점': 'Groceries', '교통': 'Transportation',
  '배달': 'Delivery', '택시': 'Taxi',
  '온라인쇼핑': 'Online', '통신': 'Phone', '의료': 'Medicine', '구독/디지털': 'Subscription',
  '의류/미용': 'Clothing/Beauty', '생필품': 'Household', '경조사/선물': 'Gifts',
  '주거/고정': 'Housing/Bills', '미분류': 'Uncategorized',
};

// 라벨을 바꾸면 이미 저장된 행은 옛 이름으로 남아, 분석에서 같은 지출이 두 항목으로 갈린다.
// 더 나쁜 건 학습이다 — 옛 이름은 지금 규칙이 내놓는 값과 다르므로 '사람이 고친 것'으로
// 읽혀서, 새 결제에 옛 이름을 도로 물려준다. 'Clean up ledger' 를 누를 때 옮겨준다.
const LEGACY_LABELS = {'Medical': 'Medicine', 'Transport': 'Transportation'};

function classify_(merchant) {
  const m = (merchant || '').toLowerCase();
  for (const [category, keywords] of CATEGORY_RULES) {
    if (keywords.some(k => m.indexOf(k.toLowerCase()) !== -1)) return category;
  }
  return '미분류';
}

function periodOf_(category) {
  return PERIOD_BY_CATEGORY[category] || '주기';
}

// 카드사마다 표기가 달라 느슨하게 잡는다.
//   신한카드(1234)승인 홍*동 94,400원(일시불)07/01 19:55 폴리스 유한
//   우리(5678)승인 김*수님 36,400원 일시불 07/20 21:54 유한회사 미소
//   신한카드 (1234) 매출취소 홍*동 29,902원 07/12 네이버페이
const APPROVAL_RE = new RegExp(
  '(?<card>[^\\s(]{1,12})\\s*\\((?<last4>\\d{3,4})\\)\\s*' +
  // 취소 문자는 '매출취소' 처럼 앞에 말이 더 붙어 온다. 접두사를 받아주지 않으면
  // 취소가 통째로 안 잡혀 환불이 영영 반영되지 않는다.
  '(?<kind>(?:매출|해외)?(?:승인|취소))\\s*' +
  '(?<name>\\S+?)\\s+' +
  '(?<amount>[\\d,]+)\\s*원\\s*' +
  // 결제유형은 괄호에 싸이기도 하고 그냥 붙어 오기도 한다.
  '(?:\\(?\\s*(?<paytype>[^()/]*?)\\s*\\)?\\s*)?' +
  // 날짜 앞의 (?<!\d) 가 핵심이다. V8 은 (?:...)? 안의 그룹이 빈 채로 도는 걸 거부해서,
  // 결제유형이 없는 문자(모든 취소 문자, 일시불을 안 붙이는 카드)에서 게으른 paytype 이
  // 억지로 한 글자를 먹는다. 그게 하필 월의 첫 자리라 12/25 가 2/25 로 읽혀 두 달 뒤에
  // 저장됐다. 01~09월은 앞의 0이 떨어져도 값이 같아 여태 안 드러났을 뿐이다.
  // 이 lookbehind 를 걸면 숫자 뒤에서는 날짜가 시작될 수 없어 paytype 이 도로 비워진다.
  // (파이썬 re 는 빈 반복을 허용해서 이 버그가 없다 — 그래서 check_port 가 못 잡았다.)
  //
  // 취소 문자에는 시각이 아예 없다(07/12 네이버페이). 시각을 필수로 두면 이런 문자가
  // 전부 파싱 실패로 빠지므로 선택으로 둔다. 없으면 00:00 으로 본다.
  '(?<!\\d)(?<mmdd>\\d{1,2}\\/\\d{1,2})(?:\\s+(?<hm>\\d{1,2}:\\d{2}))?' +
  '\\s*(?<merchant>.*)$'
);

// 카드번호가 아예 없는 형식. 현대카드가 이렇게 보낸다.
//   현대카드Z ontact 승인 홍*동 10,000원 일시불 07/27 17:10 모바일티머니선 누적206,250원
//
// 위 정규식은 '(네자리)' 를 기준점으로 삼아 문장을 끊는데, 이 문자엔 괄호가 없어
// 통째로 파싱 실패한다. 그래서 기준점을 '승인/취소' 로 바꾼 패턴을 따로 둔다.
// 이쪽이 훨씬 헐거우므로 위 정규식이 실패했을 때만 쓴다.
const APPROVAL_NO_NUM_RE = new RegExp(
  '^(?<card>.{1,20}?)\\s*' +
  '(?<kind>(?:매출|해외)?(?:승인|취소))\\s+' +
  '(?<name>\\S+?)\\s+' +
  '(?<amount>[\\d,]+)\\s*원\\s*' +
  '(?:\\(?\\s*(?<paytype>[^()/]*?)\\s*\\)?\\s*)?' +
  '(?<!\\d)(?<mmdd>\\d{1,2}\\/\\d{1,2})(?:\\s+(?<hm>\\d{1,2}:\\d{2}))?' +   // 위와 같은 이유
  '\\s*(?<merchant>.*)$'
);

// 외화로 찍히는 해외결제. 현대카드가 이렇게 보낸다.
//   현대 대한항공120 해외승인 홍*동님 10/17 16:40 JPY 2,000.00 UBER*TRIPHELP.UBER.COM
//   현대 대한항공120 해외승인 취소 홍*동님 10/17 16:41 JPY 1,258.00 UBER*PENDING
//
// 금액이 원이 아니고 날짜 뒤에 온다. 위 두 패턴은 '숫자+원' 을 기준점으로 삼으므로
// 이 문자는 통째로 파싱 실패해, 해외에서 쓴 돈이 한도·카드값에서 몽땅 빠졌다.
// 원화 금액은 여기서 정하지 않는다 — 환율이 필요해서 doPost 의 settleForeign_ 가 맡는다.
const FOREIGN_RE = new RegExp(
  '^(?<card>.{1,20}?)\\s*' +
  '(?<kind>해외(?:승인|취소)(?:\\s*취소)?)\\s+' +
  '(?<name>\\S+?)\\s+' +
  '(?<mmdd>\\d{1,2}\\/\\d{1,2})(?:\\s+(?<hm>\\d{1,2}:\\d{2}))?\\s+' +
  '(?<cur>[A-Z]{3})\\s*(?<famt>\\d[\\d,]*(?:\\.\\d+)?)' +
  '\\s*(?<merchant>.*)$'
);

// 가맹점명 뒤에 붙는 꼬리표. 카드사가 "카카오T일반 누적732,232원" 처럼 보낸다.
// 이걸 안 떼면 결제할 때마다 누적액이 달라져 매번 다른 가맹점으로 인식되고,
// 고정비 자동 감지·태그 학습·중복 방지가 전부 무력해진다.
const TRAILING_RE = /\s*(?:누적|누계|총누적|잔액|잔여한도)\s*[\d,]+\s*원?\s*$/;

/** 가맹점명에서 누적액 같은 꼬리표를 떼어낸다. */
function cleanMerchant_(merchant) {
  return String(merchant || '').replace(TRAILING_RE, '').trim();
}

/**
 * 화면에 보여줄 때만 한글을 영어로 번역한다. 시트에는 원본을 그대로 둔다.
 * 구글 번역이라 상호·법인명은 어설프게 나온다 — 읽히는 정도로만 쓴다.
 * 같은 문자열은 캐시(6시간)해서 번역 호출을 아낀다.
 */
function toEnglish_(text) {
  const s = String(text || '').trim();
  if (!s || !/[가-힣]/.test(s)) return s;       // 한글이 없으면 손대지 않는다
  const cache = CacheService.getScriptCache();
  const key = 'en_' + s;
  const hit = cache.get(key);
  if (hit !== null) return hit;
  let out = s;
  try { out = LanguageApp.translate(s, 'ko', 'en') || s; } catch (err) { /* 실패하면 원문 */ }
  cache.put(key, out, 21600);
  return out;
}

/**
 * 연도 없는 MM/DD HH:MM 을 현재연도로 보정. 미래로 튀면 작년으로 본다.
 * 시각이 없는 문자(취소)는 00:00 으로 본다 — 날짜만 맞으면 주간 계산은 그대로 맞다.
 */
function resolveYear_(mmdd, hm) {
  const now = new Date();
  const [month, day] = mmdd.split('/').map(Number);
  const [hh, mm] = String(hm || '00:00').split(':').map(Number);
  let dt = new Date(now.getFullYear(), month - 1, day, hh, mm);
  if ((dt - now) / 86400000 > 1) dt = new Date(now.getFullYear() - 1, month - 1, day, hh, mm);
  return dt;
}

/** 카드 문자 1건 → 레코드. 형식이 안 맞으면 null. */
function parseMessage_(text) {
  const norm = String(text || '').replace(/\[Web\s*발신\]/g, ' ').replace(/\s+/g, ' ').trim();
  // 카드번호가 있는 형식이 먼저. 실패하면 번호 없는 형식(현대카드)으로 한 번 더.
  const m = APPROVAL_RE.exec(norm) || APPROVAL_NO_NUM_RE.exec(norm);
  if (!m) return parseForeign_(norm);

  const g = m.groups;
  let amount = parseInt(g.amount.replace(/,/g, ''), 10);
  if (g.kind.indexOf('취소') !== -1) amount = -amount;   // '취소' / '매출취소' 둘 다

  const merchant = cleanMerchant_(g.merchant);
  const category = classify_(merchant);
  return {
    dt: resolveYear_(g.mmdd, g.hm),
    amount: amount,
    category: category,
    merchant: merchant,
    period: periodOf_(category),
    card: g.card,
    last4: g.last4 || '',        // 번호 없는 형식에서는 빈 값
    paytype: (g.paytype || '').trim(),
    kind: g.kind,
  };
}

/**
 * 외화 해외결제 문자 → 레코드. amount 는 아직 0 이다 — settleForeign_ 가 환율로 채운다.
 * 형식이 안 맞으면 null.
 */
function parseForeign_(norm) {
  const f = FOREIGN_RE.exec(norm);
  if (!f) return null;
  const g = f.groups;
  const merchant = cleanMerchant_(g.merchant);
  const category = classify_(merchant);
  return {
    dt: resolveYear_(g.mmdd, g.hm),
    amount: 0,
    category: category,
    merchant: merchant,
    period: periodOf_(category),
    card: g.card,
    last4: '',
    paytype: '',
    kind: g.kind.replace(/\s+/g, ''),          // '해외승인 취소' → '해외승인취소'
    currency: g.cur,
    foreignAmount: parseFloat(g.famt.replace(/,/g, '')),
    cancel: g.kind.indexOf('취소') !== -1,
  };
}


// ══════════════════════════════════════════════════════════════ 외화 → 원

const FX_SHEET = 'FX';            // GOOGLEFINANCE 를 돌리는 숨은 칸
const FOREIGN_MATCH_DAYS = 60;    // 해외 취소가 짝을 찾아 거슬러 올라가는 날 수

/** 2000 → '2,000.00'. 원장의 Foreign 칸과 취소 짝 맞추기가 같은 모양을 쓴다. */
function fmtForeign_(n) {
  const [i, d] = (Math.round(Number(n) * 100) / 100).toFixed(2).split('.');
  return i.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '.' + d;
}

/** Foreign 칸 'JPY 2,000.00 @9.12 +1.3%' → 'JPY 2,000.00'. 짝 맞추기의 열쇠. */
function foreignKey_(note) {
  const m = /^([A-Z]{3}) ([\d,]+\.\d{2})/.exec(String(note || '').trim());
  return m ? m[1] + ' ' + m[2] : '';
}

/** 'JPY=9.1, USD=1380' → {JPY: 9.1, USD: 1380}. */
function fallbackRates_(raw) {
  const out = {};
  String(raw || '').split(/[,;\n]/).forEach(part => {
    const m = /([A-Za-z]{3})\s*[=:]\s*([\d.]+)/.exec(part);
    if (m && Number(m[2]) > 0) out[m[1].toUpperCase()] = Number(m[2]);
  });
  return out;
}

/**
 * 1 외화 = 몇 원. 못 구하면 0.
 *
 * 숨은 시트 칸에 GOOGLEFINANCE 를 넣고 읽는다. UrlFetchApp 으로 환율 API 를 부르면
 * '외부 서비스 연결' 권한이 새로 생겨, 배포한 사람이 다시 승인하기 전까지 웹앱이
 * 통째로 멈춘다. 시트 수식은 이미 가진 권한 안이다.
 *
 * 그래도 GOOGLEFINANCE 가 가끔 #N/A 를 낸다. 그때는 설정의 FX fallback 을 쓴다.
 */
function fxRate_(cur, settings) {
  if (cur === 'KRW') return 1;
  try {
    const ss = book_();
    let sh = ss.getSheetByName(FX_SHEET);
    if (!sh) { sh = ss.insertSheet(FX_SHEET); sh.hideSheet(); }
    const cell = sh.getRange('A1');
    // cur 는 FOREIGN_RE 가 영문 대문자 세 글자로만 받아오므로 수식에 그대로 넣어도 된다
    cell.setFormula(`=GOOGLEFINANCE("CURRENCY:${cur}KRW")`);
    for (let i = 0; i < 3; i++) {
      SpreadsheetApp.flush();
      const v = Number(cell.getValue());
      if (v > 0) return v;
      Utilities.sleep(700);
    }
  } catch (err) { /* 아래 fallback 으로 */ }
  return fallbackRates_(settings[FX_FALLBACK_KEY])[cur] || 0;
}

/**
 * 외화 레코드의 원화 금액(rec.amount)과 근거(rec.foreignNote)를 채운다.
 * 저장하지 말아야 하면 그 이유(사람이 읽을 문장)를, 아니면 '' 를 돌려준다.
 *
 * 승인은 환율 × (1 + 수수료) 로 추정한다. 실제 청구액은 매입일 환율로 며칠 뒤에
 * 정해지므로 몇백 원 어긋난다 — 청구서가 나오면 Amount 를 고치면 된다.
 *
 * 취소는 짝이 되는 승인 행의 원화 금액을 **그대로** 뺀다. 그날 환율로 다시 계산하면
 * 승인 날과 환율이 달라 몇십 원이 남고, 그 찌꺼기가 '산 적 없는 지출'로 쌓인다.
 * 짝은 같은 카드사·같은 통화·같은 외화 금액. 이미 취소로 상쇄된 승인은 다시 안 쓴다.
 *
 * 짝이 없는 'PENDING' 취소는 저장하지 않는다. Uber 같은 곳은 임시금액을 승인했다가
 * 풀고 실제 요금을 따로 승인한다. 임시승인 문자를 안 보냈으면 그 취소는 되돌릴 게
 * 없는 것이라, 저장하면 쓰지도 않은 돈이 환불된 것처럼 한도가 늘어난다.
 */
function settleForeign_(rec) {
  const settings = readSettings_();
  const key = rec.currency + ' ' + fmtForeign_(rec.foreignAmount);

  if (rec.cancel) {
    const since = fmtD_(new Date(rec.dt.getTime() - FOREIGN_MATCH_DAYS * 86400000));
    const brand = cardBrand_(rec.card);
    const same = loadRecords_(since)
      .filter(r => foreignKey_(r.foreign) === key && cardBrand_(r.card) === brand);
    const paid = same.filter(r => r.amount > 0);
    const undone = same.filter(r => r.amount < 0).length;
    if (paid.length > undone) {
      const hit = paid[paid.length - 1];
      rec.amount = -hit.amount;
      rec.foreignNote = `${key} undoes ${hit.date}`;
      return '';
    }
    if (/PENDING/i.test(rec.merchant)) {
      return `Hold released, nothing to undo - not saved\n${rec.merchant} ${key}`;
    }
  }

  const rate = fxRate_(rec.currency, settings);
  if (!rate) {
    return `No exchange rate for ${rec.currency} - not saved.\n` +
           `Put e.g. "${rec.currency}=9.1" in ${SETTINGS} > ${FX_FALLBACK_KEY}, then send again.`;
  }
  const fee = Number(settings[OVERSEAS_FEE_KEY]) || 0;
  const krw = Math.round(rec.foreignAmount * rate * (1 + fee / 100));
  rec.amount = rec.cancel ? -krw : krw;
  rec.foreignNote = `${key} @${Math.round(rate * 100) / 100} +${fee}%`;
  return '';
}

/** 단축어 알림에 쓰는 '가맹점 금액'. 외화면 추정치임을 밝힌다. */
function recLabel_(rec) {
  const name = toEnglish_(rec.merchant);
  return rec.currency
    ? `${name} ≈${won_(rec.amount)} (${rec.currency} ${fmtForeign_(rec.foreignAmount)})`
    : `${name} ${won_(rec.amount)}`;
}


// ══════════════════════════════════════════════════════════════ 시트 다루기

function fmtDT_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd HH:mm'); }
function fmtD_(d)  { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
function fmtMD_(d) { return Utilities.formatDate(d, TZ, 'MM/dd'); }

function book_() { return SpreadsheetApp.getActiveSpreadsheet(); }

/** 시트가 없으면 만든다. 처음 실행 때 한 번만 일한다. */
function ensureSheets_() {
  const ss = book_();
  let ws = ss.getSheetByName(SHEET);
  if (!ws) {
    ws = ss.insertSheet(SHEET);
    ws.appendRow(HEADERS);
    ws.getRange(1, 1, 1, HEADERS.length)
      .setBackground('#0A2342').setFontColor('#FFFFFF').setFontWeight('bold');
    ws.setFrozenRows(1);
    const col = h => HEADERS.indexOf(h) + 1;       // 1부터 세는 열 번호
    ws.setColumnWidth(col('Merchant'), 200);
    ws.getRange(1, col('Date'), ws.getMaxRows()).setNumberFormat('@');   // 날짜값으로 변환 방지
    ws.getRange(1, col('Amount'), ws.getMaxRows()).setNumberFormat('#,##0');
    // 사람이 적는 칸은 Tag 하나뿐이다. 맨 끝에 두고 노랗게 칠한다.
    ws.getRange(1, col('Tag'), ws.getMaxRows()).setBackground('#FFF6E5');
    ws.setColumnWidth(col('Tag'), 140);
    // 기계용 칸은 접어둔다. 숨기면 Tag 가 Type 바로 옆에 붙는다.
    // 칸 수를 박아두면 열을 하나 끼워 넣을 때 조용히 어긋난다 — AutoCategory 를 넣었을 때
    // 실제로 'Card 부터 5칸'의 끝이 밀려, 새로 만든 시트에서만 그 칸이 드러나 있었다.
    // 이름으로 숨기면 HEADERS 를 고쳐도 따라온다.
    ws.hideColumns(col('Card'), 4);                                  // Card·Last4·Payment·Kind
    Object.keys(HIDDEN_COLUMNS).forEach(h => ws.hideColumns(col(h)));
  } else {
    ensureLedgerColumns_(ws);   // 이미 있는 시트에 빠진 열(MerchantEN 등)을 붙인다
  }
  let st = ss.getSheetByName(SETTINGS);
  if (!st) {
    st = ss.insertSheet(SETTINGS);
    st.appendRow(['Item', 'Value', 'Note']);
    st.getRange(1, 1, 1, 3)
      .setBackground('#0A2342').setFontColor('#FFFFFF').setFontWeight('bold');
    st.setColumnWidth(1, 160); st.setColumnWidth(3, 300);
  }
  ensureSettingsRows_(st);

  // 예시 줄은 금액을 0으로 둔다. 0이면 계산에서 걸러지므로 지우지 않아도 숫자가 틀어지지 않는다.
  // 시트에 보이는 글자는 전부 영어다.
  ensureListSheet_(ss, FIXED_SHEET, FIXED_HEADERS, [
    ['(example) Rent', 0, 'Month', 'Bank', "Never hits your card — listing it here is the only way it counts", 25, ''],
    ['(example) Netflix', 0, 'Month', 'Card', 'Match is optional when the name already fits', 22, ''],
    ['(example) Cleaner', 0, 'Week', 'Cash', 'Amount for ONE visit', '', ''],
    ['(example) Taxi to work', 0, 'Mon-Fri', 'Card', 'Amount for ONE day, not the week', '', '카카오T'],
  ]);
  ensureListSheet_(ss, EXTRA_SHEET, EXTRA_HEADERS, [
    ['', '(example) Insurance refund', 0, '+ money in, - money out that skipped your card'],
  ]);
  return ws;
}

/**
 * 날짜 칸을 전부 텍스트 'yyyy-MM-dd HH:mm' 로 통일한 뒤 날짜순으로 정렬한다.
 *
 * 손으로 적은 행은 구글시트가 날짜값('2026. 7. 20')으로 바꿔버려, 자동 입력한
 * 텍스트('2026-07-20')와 형식이 갈린다. 그 상태로 문자열 정렬하면 두 덩어리가
 * 따로 놀아 새 결제가 엉뚱한 자리에 낀다. 매번 형식을 맞춰 이 문제를 없앤다.
 */
function normalizeDatesAndSort_(ws) {
  const last = ws.getLastRow();
  if (last < 3) return;                 // 데이터가 한 줄뿐이면 정렬할 것도 없다
  const idx = headerIndex_(ws);
  const dcol = (idx['Date'] === undefined ? 0 : idx['Date']) + 1;

  const rng = ws.getRange(2, dcol, last - 1, 1);
  const norm = rng.getValues().map(r => [toDTString_(r[0])]);
  rng.setNumberFormat('@').setValues(norm);   // 날짜값 → 통일된 텍스트

  ws.getRange(2, 1, last - 1, ws.getLastColumn()).sort({column: dcol, ascending: true});
  markMonths_(ws, idx);
}

/**
 * 달이 바뀌는 자리를 눈에 띄게 만든다.
 *
 * 한 시트에 계속 쌓이는 구조라, 아무 표시가 없으면 7월 끝과 8월 시작이 구분되지 않는다.
 * 달마다 배경을 번갈아 칠해서 덩어리로 보이게 한다.
 *
 * 선은 긋지 않는다. 선은 데이터가 아니라 '그 자리'에 남는데, 새 결제가 들어올 때마다
 * 정렬로 행이 움직여도 따라가지 않아 엉뚱한 줄에 밑줄이 계속 쌓였다. 색만으로도
 * 달 구분은 충분하다. 예전에 그어둔 선은 여기서 지운다.
 *
 * 배경은 행 전체를 한 번에 칠하므로 Tag 칸의 노란색이 지워진다. 같은 배열 안에서
 * 다시 노랗게 칠해 되살린다 — 안 그러면 사람이 적는 칸 표시가 사라진다.
 */
function markMonths_(ws, idx) {
  const last = ws.getLastRow();
  const width = ws.getLastColumn();
  if (last < 2 || width < 1) return;

  const dcol = (idx['Date'] === undefined ? 0 : idx['Date']);
  const tagCol = idx['Tag'];
  const rows = ws.getRange(2, 1, last - 1, width).getValues();

  const shades = ['#FFFFFF', '#F3F6FB'];
  const colors = [];
  let seen = '', band = -1;

  rows.forEach(r => {
    const month = String(toDTString_(r[dcol]) || '').slice(0, 7);
    if (month !== seen) {
      seen = month;
      band++;
    }
    const fill = shades[band % 2];
    const row = new Array(width).fill(fill);
    if (tagCol !== undefined && tagCol < width) row[tagCol] = '#FFF6E5';
    colors.push(row);
  });

  const data = ws.getRange(2, 1, last - 1, width);
  data.setBackgrounds(colors);
  data.setBorder(false, false, false, false, false, false);   // 쌓인 밑줄 제거
}

/**
 * 손으로 넣은 행까지 싹 정리한다: 카테고리·가맹점을 영어로 채우고,
 * 날짜 형식을 통일한 뒤 정렬한다. 과거 내역을 붙여넣은 다음 한 번 누르면 된다.
 *
 * 자동 저장 행은 이미 영어라 그대로 지나간다. 번역이 걸리는 부분이라
 * 화면 로딩이 아니라 이 수동 실행에서만 돈다.
 */
/**
 * 없앨 열이 실제로 있을 때만 한 번 묻는다.
 *
 * 열을 지우는 건 되돌릴 수 없고 메뉴 이름('Clean up')만 봐서는 그런 일이 일어나는 줄
 * 모른다. 값 자체는 카테고리에서 언제든 다시 나오므로 잃는 정보는 없지만, 남의 시트에서
 * 열이 소리 없이 사라지는 건 그것대로 놀랄 일이다.
 *
 * 카테고리는 여기서 묻지 않는다 — 정리 작업이 사람이 적은 값을 덮어쓰지 않기 때문이다.
 */
function confirmDropColumns_(ws) {
  const idx = headerIndex_(ws);
  const found = RETIRED_COLUMNS.filter(h => idx[h] !== undefined);
  if (!found.length) return true;

  const ui = SpreadsheetApp.getUi();
  const res = ui.alert('Remove unused column',
    `This will delete the ${found.join(', ')} column from the Ledger.\n\n` +
    'Nothing reads it — the value always follows from Category, so nothing is lost.\n' +
    'Your own edits in other columns are not touched.\n\nContinue?',
    ui.ButtonSet.YES_NO);
  return res === ui.Button.YES;
}

/**
 * 이제 안 쓰는 열을 시트에서 없앤다.
 *
 * 그냥 두면 새 결제부터는 빈칸으로 남아 반쯤 채워진 열이 되고, 보는 사람은
 * 그게 고장인지 안 쓰는 건지 알 수 없다. 값은 카테고리에서 언제든 다시 나오므로
 * 잃는 정보도 없다. 이름으로 찾아 지우므로 열 순서를 바꿔놨어도 맞게 지운다.
 */
function dropRetiredColumns_(ws) {
  RETIRED_COLUMNS.forEach(name => {
    const idx = headerIndex_(ws);          // 하나 지울 때마다 번호가 밀리므로 매번 다시 읽는다
    if (idx[name] !== undefined) ws.deleteColumn(idx[name] + 1);
  });
}

function fixLedgerOrder() {
  const ws = book_().getSheetByName(SHEET);
  if (ws) {
    ensureLedgerColumns_(ws);
    if (!confirmDropColumns_(ws)) return;
    dropRetiredColumns_(ws);
    backfillEnglish_(ws);
    reclassifyAuto_(ws);          // 이름을 옮긴 뒤에 돌아야 옛 라벨과 안 헷갈린다
    normalizeDatesAndSort_(ws);
  }
  // 카테고리를 통째로 다시 칠했으니 캐시된 화면은 거짓말이 된다. onEdit 은 스크립트가
  // 고친 것은 못 잡으므로(단순 트리거) 여기서 직접 버려야 한다.
  dropViewCache_();
  SpreadsheetApp.getActiveSpreadsheet().toast('Ledger cleaned up.');
}

/**
 * 시트에 적힌 카테고리 값을 '지금 쓰는 이름'으로 맞춘다.
 * 한글이면 영어로, 이름이 바뀐 것이면 새 이름으로. 이미 맞으면 그대로 돌려준다.
 */
function currentLabel_(v) {
  const s = String(v || '').trim();
  if (!s) return s;
  if (/[가-힣]/.test(s)) return CATEGORY_LABEL[s] || toEnglish_(s);
  return LEGACY_LABELS[s] || s;
}

/** 비어 있는 영어 칸(카테고리·MerchantEN)을 채운다. 이미 영어면 건드리지 않는다. */
function backfillEnglish_(ws) {
  const last = ws.getLastRow();
  if (last < 2) return;
  const idx = headerIndex_(ws);
  const cCol = idx['Category'], mCol = idx['Merchant'], eCol = idx['MerchantEN'];
  const aCol = idx['AutoCategory'];
  const rng = ws.getRange(2, 1, last - 1, ws.getLastColumn());
  const vals = rng.getValues();
  let changed = false;

  vals.forEach(r => {
    // Category 와 AutoCategory 를 **같이** 옮긴다. 한쪽만 옮기면 둘이 갈라지고,
    // 갈라진 순간 그 행은 '사람이 고친 것'으로 읽힌다(learnedCategory_). 카테고리
    // 이름을 한 번 바꿀 때마다 과거 행이 통째로 그렇게 둔갑하는 게 원래 버그였다.
    [cCol, aCol].forEach(col => {
      if (col === undefined) return;
      const cur = String(r[col] || '').trim();
      const next = currentLabel_(cur);
      if (next !== cur) { r[col] = next; changed = true; }
    });
    if (mCol !== undefined && eCol !== undefined) {
      const merch = String(r[mCol] || '').trim();
      if (merch && !String(r[eCol] || '').trim()) {   // 번역이 비어 있으면 채운다
        r[eCol] = toEnglish_(merch);
        changed = true;
      }
    }
  });
  if (changed) rng.setValues(vals);
}

/**
 * 규칙이 바뀐 뒤 옛 라벨로 남아 있는 행을 지금 규칙으로 다시 분류한다.
 *
 * 분류는 저장할 때 한 번만 한다. 그래서 키워드를 고치거나 카테고리를 쪼개도(카카오T를
 * 교통에서 택시로) 지나간 행은 옛 이름 그대로 남고, 지난달과 이번 달이 다른 이름으로
 * 갈려 비교가 무의미해진다.
 *
 * 손대지 않은 행(Category === AutoCategory)만 옮긴다 — 사람이 고친 행은 건드리지 않는다.
 *
 * AutoCategory 가 빈 옛 행에는 **지금 규칙이 내놓는 값**을 적어 넣는다. Category 는
 * 손대지 않는다. 그러면 시트 값이 규칙과 다른 행은 저절로 '사람이 고친 것'으로 남는다.
 *
 * 반대로(빈 칸에 Category 를 그대로 복사) 하면 손으로 고쳐둔 행이 전부 '규칙대로
 * 저장된 것'으로 읽혀 규칙 값으로 되돌아간다. 고친 게 몇십 줄이면 그게 다 날아간다.
 * 정리 작업이 사람이 타이핑한 값을 지우는 건 어떤 이유로도 안 된다 — 되돌릴 수 없고,
 * 그 사실이 화면 어디에도 안 남는다.
 *
 * 대신 규칙을 바꾼 뒤에도 옛 행은 옛 이름을 지킨다(사람이 고친 것으로 보이므로).
 * 그 가맹점의 새 결제가 옛 이름으로 들어오면 한 번만 고쳐주면 된다 — 학습이 가장
 * 최근에 고친 값을 따르므로 그 한 번으로 굳는다.
 */
function reclassifyAuto_(ws) {
  const last = ws.getLastRow();
  if (last < 2) return;
  const idx = headerIndex_(ws);
  const cCol = idx['Category'], mCol = idx['Merchant'], aCol = idx['AutoCategory'];
  if (cCol === undefined || mCol === undefined || aCol === undefined) return;

  const rng = ws.getRange(2, 1, last - 1, ws.getLastColumn());
  const vals = rng.getValues();
  let changed = false;

  vals.forEach(r => {
    const merch = String(r[mCol] || '').trim();
    if (!merch) return;
    const cat = String(r[cCol] || '').trim();

    const ruleKey = classify_(merch);
    const ruleLabel = CATEGORY_LABEL[ruleKey] || ruleKey;

    let auto = String(r[aCol] || '').trim();
    // 옛 행을 메꾼다. Category 가 아니라 **규칙 값**을 적는다 — 위 주석 참고.
    if (!auto && cat) { auto = ruleLabel; r[aCol] = ruleLabel; changed = true; }
    if (cat !== auto) return;                                   // 사람이 고친 행은 그대로

    const key = ruleKey;
    const label = CATEGORY_LABEL[key] || key;
    if (label === cat) return;

    r[cCol] = label;
    r[aCol] = label;
    changed = true;
  });

  if (changed) rng.setValues(vals);
}

/**
 * 이미 있는 Ledger 에 빠진 열을 뒤에 붙인다. 적어둔 값은 건드리지 않는다.
 * MerchantEN(번역 저장 칸)·AutoCategory(규칙이 내놓은 값)를 나중에 추가했기 때문에 필요하다.
 */
function ensureLedgerColumns_(ws) {
  const idx = headerIndex_(ws);
  HEADERS.forEach(h => {
    if (idx[h] === undefined) {
      const c = ws.getLastColumn() + 1;
      ws.getRange(1, c).setValue(h)
        .setBackground('#0A2342').setFontColor('#FFFFFF').setFontWeight('bold');
      if (HIDDEN_COLUMNS[h]) ws.hideColumns(c);
    }
  });
}

/**
 * 설정 시트에 빠진 항목을 채워 넣는다. 이미 적어둔 값은 건드리지 않는다.
 * 나중에 설정이 늘어나도 기존 시트에 저절로 붙도록.
 */
function ensureSettingsRows_(st) {
  const last = st.getLastRow();
  const have = last > 1
    ? st.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0] || '').trim())
    : [];
  DEFAULT_SETTINGS.forEach(row => {
    if (have.indexOf(row[0]) === -1) st.appendRow(row);
  });
  refreshSettingsNotes_(st);
}

/**
 * 설명(Note) 칸을 최신 문구로 맞춘다.
 *
 * 설명은 우리가 쓰는 안내문이지 사용자가 적은 값이 아니다. 그래서 덮어써도 된다.
 * 이게 없으면 설명을 아무리 고쳐도 이미 만들어진 시트에는 영영 반영되지 않는다.
 * Value 열은 사용자가 적은 값이므로 절대 건드리지 않는다.
 */
function refreshSettingsNotes_(st) {
  const last = st.getLastRow();
  if (last < 2) return;

  const want = {};
  DEFAULT_SETTINGS.forEach(([key, , note]) => { want[key] = note; });

  const items = st.getRange(2, 1, last - 1, 1).getValues();
  const notes = st.getRange(2, 3, last - 1, 1).getValues();   // Note 열만 읽고 쓴다
  let changed = false;

  items.forEach((row, i) => {
    const key = String(row[0] || '').trim();
    if (want[key] !== undefined && notes[i][0] !== want[key]) {
      notes[i][0] = want[key];
      changed = true;
    }
  });

  if (changed) {
    st.getRange(2, 3, last - 1, 1).setValues(notes);
    st.setColumnWidth(3, 560);      // 설명이 길어져서 넓혀준다
  }
}

/**
 * Fixed / Extra 처럼 사용자가 직접 채우는 목록 시트를 만든다.
 * 이미 있으면 부족한 열만 뒤에 덧붙인다 — 적어둔 내용은 건드리지 않는다.
 */
function ensureListSheet_(ss, name, headers, samples) {
  const existing = ss.getSheetByName(name);
  if (existing) {
    const width = Math.max(existing.getLastColumn(), 1);
    const cur = existing.getRange(1, 1, 1, width).getValues()[0].map(v => String(v || '').trim());
    headers.forEach((h, i) => {
      if (cur.indexOf(h) === -1) {
        existing.getRange(1, i + 1).setValue(h)
          .setBackground('#0A2342').setFontColor('#FFFFFF').setFontWeight('bold');
      }
    });
    return;
  }
  const sh = ss.insertSheet(name);
  sh.appendRow(headers);
  sh.getRange(1, 1, 1, headers.length)
    .setBackground('#0A2342').setFontColor('#FFFFFF').setFontWeight('bold');
  sh.setFrozenRows(1);
  samples.forEach(r => sh.appendRow(r));
  // 예시 줄은 지우고 쓰라는 뜻으로 흐리게
  if (samples.length) {
    sh.getRange(2, 1, samples.length, headers.length).setFontColor('#9AA0A6').setFontStyle('italic');
  }
  sh.setColumnWidth(1, 150); sh.setColumnWidth(2, 110);
  sh.setColumnWidth(3, 90); sh.setColumnWidth(4, 260);
}

/**
 * Monthly 시트에 이번 달 한 줄을 써 넣는다(있으면 갱신).
 *
 * 카드 합계와 카테고리는 Ledger 만 있으면 언제든 다시 셀 수 있다. 하지만 고정비·월급은
 * '그때의 Settings' 라 지나가면 복원할 방법이 없다. 그래서 지나간 달 줄은 건드리지 않고,
 * 이번 달 줄만 계속 덮어쓴다 — 달이 바뀌면 그 값이 그대로 굳는다.
 */
function updateMonthly_(s, force) {
  // 말일에만 기록한다. 결제가 들어올 때마다 덮어쓰면 한 달 내내 숫자가 출렁이고,
  // 그 달이 끝나기 전 값은 어차피 중간 결과라 남겨둘 이유가 없다.
  // 메뉴에서 부를 때는(force) 언제든 쓴다 — 말일에 결제가 없으면 줄이 안 생기므로
  // 손으로 남길 길은 열어둬야 한다.
  const today = new Date();
  const lastDay = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  if (!force && today.getDate() !== lastDay) return false;

  const ss = book_();
  let sh = ss.getSheetByName(MONTHLY_SHEET);
  if (!sh) {
    sh = ss.insertSheet(MONTHLY_SHEET);
    sh.appendRow(MONTHLY_HEADERS);
    sh.getRange(1, 1, 1, MONTHLY_HEADERS.length)
      .setBackground('#0A2342').setFontColor('#FFFFFF').setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 90);
    sh.setColumnWidth(MONTHLY_HEADERS.length, 320);
    sh.getRange(1, 2, sh.getMaxRows(), 6).setNumberFormat('#,##0');
  }

  const month = Utilities.formatDate(new Date(), TZ, 'yyyy-MM');
  // 달 전체로 센 목록을 쓴다 — 화면용 s.categories 는 지난달과 견주려고 날짜를 잘라둔 값이다
  const tops = (s.categoriesFull || []).slice(0, 4)
    .map(c => `${c.name} ${won_(c.amount)}`).join(', ');

  const row = [month, s.monthCharged, s.monthOut - s.monthCharged, s.monthOut,
               s.monthIn, s.savings, s.budget, tops];

  const last = sh.getLastRow();
  const months = last > 1
    ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0] || '').trim())
    : [];
  const at = months.indexOf(month);

  if (at === -1) {
    sh.appendRow(row);
  } else {
    sh.getRange(at + 2, 1, 1, row.length).setValues([row]);
  }
  return true;
}

/** 시트 메뉴에서 이번 달 요약을 지금 찍을 때. 말일이 아니어도 쓴다. */
function updateMonthly() {
  updateMonthly_(summarizeWeek_(), true);
  dropViewCache_();      // 스크립트가 고친 것은 onEdit 이 못 잡는다
  SpreadsheetApp.getActiveSpreadsheet().toast('Monthly updated.');
}

/** Fixed 시트 → [{name, amount, every, byCard}]. 금액이 없는 줄은 버린다. */
function readFixed_() {
  const sh = book_().getSheetByName(FIXED_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, FIXED_HEADERS.length).getValues()
    .filter(r => r[0] && Number(r[1]))
    .map(r => ({
      name: String(r[0]).trim(),
      amount: Number(r[1]),
      every: normalizeEvery_(r[2]),
      // 'Card' 로 적으면 카드 결제 → Ledger 에도 찍히므로 이중 계산을 막아야 한다
      byCard: String(r[3] || '').trim().toLowerCase().indexOf('card') === 0,
      paidBy: String(r[3] || '').trim() || 'Bank',   // 화면에 '현금/카드' 를 적으려고 원문도 둔다
      day: Math.min(31, Math.max(0, Math.floor(Number(r[5]) || 0))),   // 결제일 (월 단위 항목만)
      // 카드 내역과 맞출 문자열. 비워두면 Name 을 쓴다.
      match: String(r[6] || '').trim() || String(r[0]).trim(),
    }));
}

/** 그 달에 없는 날짜(2월 31일 등)는 말일로 당긴다. */
function dayInMonth_(year, month, day) {
  const last = new Date(year, month + 1, 0).getDate();
  return Math.min(day, last);
}

/** 'yyyy-MM'. month 가 -1 이나 12 여도 알아서 해를 넘긴다. */
function monthKey_(year, month) {
  return Utilities.formatDate(new Date(year, month, 1), TZ, 'yyyy-MM');
}

/**
 * '현대카드Z ontact' → 'Hyundai'. 못 알아보면 첫 낱말을 그대로 쓴다.
 *
 * 카드사가 안 적힌 행(손으로 넣은 과거 내역)은 'Card' 로 뭉뚱그리면 안 된다.
 * 예전에는 그렇게 해서, 대시보드에 없는 카드사의 청구서가 한 장 더 있는 것처럼
 * 보였다 — 가장 큰 줄이 하필 '어느 카드인지 모르는 돈' 이었다.
 */
function cardBrand_(raw) {
  const s = String(raw || '').trim();
  if (!s) return 'unknown card';
  const up = s.toUpperCase();
  const hit = CARD_BRANDS.filter(b => up.indexOf(b[0].toUpperCase()) !== -1)[0];
  return hit ? hit[1] : s.split(/\s+/)[0];
}

/**
 * 그 날 쓴 돈이 통장에서 빠져나가는 결제일. 결제일 당일에 쓴 것은 그날 청구서에 실린다.
 * month 는 0부터. 12월에 넘겨도 Date 가 알아서 해를 넘긴다.
 */
function billDayFor_(year, month, day, cardDay) {
  const here = dayInMonth_(year, month, cardDay);
  return day <= here ? new Date(year, month, here)
                     : new Date(year, month + 1, dayInMonth_(year, month + 1, cardDay));
}

/**
 * 청구서별·카드사별 승인 합계. 키는 그 돈이 빠져나가는 **결제일**이다.
 *   {'2026-08-31': {Shinhan: 1200000, Hyundai: 80000}}
 *
 * 달로 묶지 않는다. 결제일이 말일이면 그 달 전체가 한 장이 되지만, 15일이면 16일에
 * 쓴 돈은 다음 장으로 넘어간다 — 달로 묶으면 결제일에 아직 쓰지도 않은 돈까지
 * 청구서에 실린다. 아래 pending 도 같은 규칙으로 쌓이므로 둘이 어긋나지 않는다.
 */
function cardTotalsByBill_(records, cardDay) {
  const out = {};
  records.forEach(r => {
    const s = String(r.date || '').slice(0, 10);
    if (s.length < 10) return;
    const [y, m, d] = s.split('-').map(Number);
    const k = fmtD_(billDayFor_(y, m - 1, d, cardDay));
    const brand = cardBrand_(r.card);
    out[k] = out[k] || {};
    out[k][brand] = (out[k][brand] || 0) + r.amount;
  });
  return out;
}

/**
 * 통장 잔액을 하루씩 굴려 앞으로 FORECAST_DAYS 일을 내다본다.
 * 잔액을 안 적어두면(0) 건너뛴다 — 은행과 연결된 게 아니라 직접 적은 값이기 때문.
 *
 * 매일 쓰는 돈 = (월 예산 − 월 고정비) ÷ 30.
 * 택시·청소 같은 주 단위 지출도 이 안에서 나가므로 따로 빼지 않는다.
 *
 * 다만 '쓰는 날'과 '통장에서 나가는 날'은 다르다. 카드로 긁으면 그날이 아니라
 * 결제일에 한 달치가 통째로 빠진다. 그래서 결제일(Card payment day)을 적어뒀으면
 * 하루치를 그 달 몫으로 쌓아두기만 하고, 결제일에 몰아서 뺀다. 예전에는 매일 빼면서
 * 결제일에 또 빼서 같은 돈이 두 번 나갔고, 최저 잔액이 최대 한 달치만큼 낮게 나왔다.
 *
 * 청구 기간은 '그 달 1일~말일'이다. 결제일이 말일 근처인 카드를 쓰고 있어서,
 * 8/31 에 빠지는 건 8월에 쓴 돈이다. 예전에는 무조건 직전 달을 청구했는데, 그러면
 * 8월 대시보드가 7월 원장을 청구서라고 내놔 카드사 앱과 100만원 넘게 어긋났다.
 */
function forecast_(settings, monthlyRows, fixedTotal, cycle, records) {
  const balance = Number(settings[BALANCE_KEY]) || 0;
  // 마이너스 통장 한도. 0 아래로 얼마까지 내려가도 되는지를 양수로 적는다.
  // 이게 있으면 잔액이 음수로 내려가는 건 사고가 아니다 — 한도를 넘을 때가 사고다.
  const overdraft = Math.abs(Number(settings[OVERDRAFT_KEY]) || 0);
  if (!balance && !overdraft) return null;

  const budget = Number(settings[BUDGET_KEY]) || 0;
  const income = Number(settings[INCOME_KEY]) || 0;
  const payday = Math.min(31, Math.max(0, Math.floor(Number(settings[PAYDAY_KEY]) || 0)));

  // 주기 예산을 직접 적었으면 그쪽이 실제 씀씀이에 가깝다
  const cycleBudget = Number(settings[CYCLE_BUDGET_KEY]) || 0;
  const dailyBurn = cycleBudget > 0
    ? cycleBudget / cycle.length
    : Math.max(0, budget - fixedTotal) / 30;

  // 잔액 기준일이 지났으면 그날부터 굴려온다 → 오래된 잔액은 저절로 낮게 잡힌다
  const raw = settings[BALANCE_DATE_KEY];
  const today = new Date();
  let cursor = raw ? new Date(toDTString_(raw).slice(0, 10)) : new Date(today);
  if (isNaN(cursor.getTime()) || cursor > today) cursor = new Date(today);
  cursor = new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate());

  const endDate = new Date(today.getFullYear(), today.getMonth(), today.getDate() + FORECAST_DAYS);
  // 목록에는 '이번 달에 앞으로 나갈 돈'만 적는다. 잔액 최저점은 그 뒤까지 내다본다.
  const listEnd = new Date(today.getFullYear(), today.getMonth() + 1, 0);
  const todayMidnight = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  // 카드값은 한 달치가 결제일에 통째로 빠진다. 이걸 안 넣으면 예측이 크게 빗나간다.
  const cardDay = Math.min(31, Math.max(0, Math.floor(Number(settings[CARD_DAY_KEY]) || 0)));
  const cardTotals = cardDay ? cardTotalsByBill_(records || [], cardDay) : {};

  let bal = balance;
  let min = balance;
  let minDate = new Date(cursor);
  const schedule = [];
  // 카드로 쓴 돈이 결제일까지 머무는 곳. 키는 cardTotals 와 같은 결제일이다.
  // {'2026-08-31': 쌓인 금액}
  const pending = {};

  while (cursor <= endDate) {
    const y = cursor.getFullYear(), m = cursor.getMonth(), d = cursor.getDate();
    const upcoming = cursor >= todayMidnight;
    const inList = upcoming && cursor <= listEnd;

    if (payday && d === dayInMonth_(y, m, payday) && income) {
      bal += income;
      if (inList) schedule.push({date: new Date(cursor), name: 'Salary', amount: income, via: 'in'});
    }

    // 하루치 씀씀이를 '어느 청구서에 실릴 돈'으로 먼저 쌓는다.
    //
    // 결제일을 적어뒀으면 이 돈은 오늘 통장에서 나가지 않는다 — 결제일에 한꺼번에
    // 빠진다. 매일 빼면서 결제일에 또 빼면 같은 돈이 두 번 나가, 최저 잔액이 한 달치만큼
    // 낮게 찍히고 빨간 'Account runs out' 경고가 잘못 뜬다. 결제일을 안 적었으면 언제
    // 빠지는지 알 길이 없으니 예전처럼 매일 뺀다.
    //
    // 결제일이 지난 뒤에 쓴 돈은 이번 청구서에 못 실린다 — 다음 달 몫이다.
    // 쌓는 일이 결제일 처리보다 먼저여야 결제일 당일에 쓴 돈이 그날 청구서에 들어간다.
    if (cardDay) {
      const k = fmtD_(billDayFor_(y, m, d, cardDay));
      pending[k] = (pending[k] || 0) + dailyBurn;
    } else {
      bal -= dailyBurn;
    }

    // 카드 결제일 — 그 달에 카드로 쓴 돈이 통째로 빠져나간다.
    //
    // 청구액은 원장 합계 **그대로**다. 카드사별로 한 줄씩 세우므로 카드사 앱의 결제
    // 예정액과 바로 견줄 수 있다. 예전에는 아직 안 끝난 달이면 여기에 '앞으로 쓸 몫'
    // (예산에서 나온 하루치)을 더해 'Card bill · projected' 한 줄로 뭉갰는데, 그러면
    // 숫자가 어긋났을 때 원장이 틀린 건지 예산이 틀린 건지 가릴 수가 없었다.
    //
    // 어느 원장 행이 이 장에 실리는지는 cardTotalsByBill_ 이 이미 갈라놨다.
    if (cardDay && d === dayInMonth_(y, m, cardDay)) {
      const billKey = fmtD_(cursor);
      const bills = cardTotals[billKey] || {};
      Object.keys(bills).sort().forEach(brand => {
        const amount = bills[brand];
        if (!amount) return;   // 0 만 건너뛴다. 취소가 승인보다 많은 달은 음수 그대로 둔다
        bal -= amount;
        if (inList) {
          schedule.push({date: new Date(cursor), name: 'Card bill · ' + brand,
                         amount: -amount, via: 'Card'});
        }
      });

      // 아직 안 쓴 몫. 원장에 없는 돈이므로 청구서와 같은 날 **따로** 세운다.
      // 이게 없으면 앞으로 쓸 돈이 통장에서 영영 안 나가 최저 잔액이 부풀고,
      // 그 최저점으로 주간 한도를 깎는 shortfall 이 늘 0이 된다.
      const toCome = Math.round(pending[billKey] || 0);
      if (toCome > 0) {
        bal -= toCome;
        if (inList) {
          schedule.push({date: new Date(cursor), name: 'Spending still to come',
                         amount: -toCome, via: 'Card'});
        }
      }
      delete pending[billKey];        // 청구됐으니 쌓아둔 몫은 비운다
    }

    monthlyRows.forEach(f => {
      // 카드로 내는 고정비는 이미 위의 카드값 안에 들어 있다. 또 빼면 두 번 나간다.
      // 카드 결제일을 안 적어둔 경우에만 예전처럼 개별로 뺀다.
      if (cardDay && f.byCard) return;
      if (f.day && d === dayInMonth_(y, m, f.day)) {
        bal -= f.amount;
        if (inList) {
          schedule.push({date: new Date(cursor), name: f.name, amount: -f.amount, via: f.paidBy});
        }
      }
    });

    if (bal < min) { min = bal; minDate = new Date(cursor); }
    cursor.setDate(cursor.getDate() + 1);
  }

  // 남은 지출을 카드와 통장(현금)으로 갈라 보여준다
  const out = schedule.filter(x => x.amount < 0);
  const outCard = out.filter(x => x.via === 'Card').reduce((s, x) => s - x.amount, 0);
  const outBank = out.filter(x => x.via !== 'Card').reduce((s, x) => s - x.amount, 0);

  return {
    balance: balance,
    min: Math.round(min),
    minDate: minDate,
    overdraft: overdraft,
    inOverdraft: min < 0,                // 마이너스 구간에 들어간다 (한도 안이면 경고는 아님)
    negative: min < -overdraft,          // 한도까지 다 쓴다 = 진짜 경고
    dailyBurn: Math.round(dailyBurn),
    cardDay: cardDay,
    outCard: outCard, outBank: outBank,
    schedule: schedule.slice(0, 12),
  };
}

/**
 * Every 칸의 값을 정해진 넷 중 하나로 맞춘다. 비어 있으면 월 단위로 본다.
 *
 * 'Weekday' 는 "평일마다"라는 뜻인데 한글로 읽으면 "주당"으로 오해하기 쉬워
 * 'Mon-Fri' 를 기본 표기로 쓴다. 예전 표기도 그대로 받는다.
 */
function normalizeEvery_(v) {
  const s = String(v || '').trim().toLowerCase();
  // 'month' 를 먼저 걸러야 한다 — 안 그러면 'mon' 접두사에 걸려 평일로 읽힌다
  if (s.indexOf('month') === 0 || s.charAt(0) === '월') return 'month';
  if (s.indexOf('mon') === 0 || s.indexOf('weekday') === 0 || s.indexOf('평일') === 0) {
    return 'weekday';
  }
  if (s.indexOf('week') === 0 || s.charAt(0) === '주') return 'week';
  if (s.indexOf('day') === 0 || s.indexOf('daily') === 0 ||
      s.indexOf('매일') === 0 || s.charAt(0) === '일') return 'day';
  return 'month';
}

/**
 * from~to 사이에 그 항목이 몇 번 나가는가.
 *   day     : 날 수 전부
 *   weekday : 월~금만
 *   week    : 7일마다 한 번 (올림). 무슨 요일인지는 모르므로 넉넉히 잡는다.
 *             7일 주기에서는 항상 1이 되어 예전 동작과 같다.
 */
function unitsBetween_(every, from, to) {
  // 자정끼리 견준다. from 에는 시각이 붙은 '지금'이 오고 to 는 마지막 날 자정이라,
  // 그냥 비교하면 주기 마지막 날 00:01 부터 to < from 이 되어 예약이 통째로 0이 된다.
  // 안 낸 청소비 5만원이 사라지고 '남은 돈'이 그만큼 늘어난다 — 가장 아껴야 할 날에.
  if (midnight_(to) < midnight_(from)) return 0;
  const days = Math.floor((midnight_(to) - midnight_(from)) / 86400000) + 1;
  if (every === 'week') return Math.ceil(days / 7);
  if (every === 'day') return days;

  let n = 0;
  const d = midnight_(from);
  while (d <= to) {
    const dow = d.getDay();                       // 0=일, 6=토
    if (dow !== 0 && dow !== 6) n++;
    d.setDate(d.getDate() + 1);
  }
  return n;
}

/** Extra 시트 → [{date, name, amount}]. + 는 들어온 돈, - 는 카드 밖 지출. */
function readExtra_() {
  const sh = book_().getSheetByName(EXTRA_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, EXTRA_HEADERS.length).getValues()
    .filter(r => r[0] && Number(r[2]))
    .map(r => ({
      date: toDTString_(r[0]).slice(0, 10),
      name: String(r[1] || '').trim(),
      amount: Number(r[2]),
    }));
}

/** 가맹점명과 Fixed 항목 이름이 같은 것인지. 양방향 부분일치로 느슨하게 본다. */
function nameMatches_(merchant, name) {
  const m = String(merchant || '').toLowerCase();
  const n = String(name || '').toLowerCase();
  return !!m && !!n && (m.indexOf(n) !== -1 || n.indexOf(m) !== -1);
}

/** 가맹점이 Fixed 목록의 '카드' 항목에 해당하는지. */
function matchesFixed_(merchant, fixedRows) {
  return fixedRows.some(f => f.byCard && nameMatches_(merchant, f.match));
}

/**
 * 내역 한 줄이 Fixed 항목에 해당하는지.
 * Match(없으면 Name)를 가맹점명과 맞춰보고, 그래도 안 되면 태그로 이어준다.
 */
function linkedToFixed_(r, f) {
  return nameMatches_(r.merchant, f.match) || nameMatches_(r.tag, f.name);
}

/**
 * 변동비에서 빼둘 줄을 고른다. {줄 번호: 가져간 고정비 이름} 으로 돌려준다.
 *
 * 이름까지 돌려주는 건 화면에 그대로 보여주기 위해서다. 이름 맞추기가 헐거워
 * 엉뚱한 결제를 집을 수 있는데, 안 보여주면 왜 남은 돈이 이상한지 알아낼 길이 없다.
 *
 * 빼는 건 '월 예산에서 이미 뺀 것'뿐이다 — Fixed 시트에 적었거나 자동으로 잡힌 항목.
 * 핵심은 **적어둔 금액만큼 한 건만** 뺀다는 것이다.
 *
 * 예전에는 이름이 걸리는 줄을 전부 뺐다. 그런데 이름 맞추기(nameMatches_)가 양방향
 * 부분일치라 Fixed 의 'GS 인터넷' 한 줄이 GS25 편의점 결제를 몽땅 삼켰다. 예산에서
 * 뺀 건 4만원인데 7만 5천원이 안 세지면 그 차액은 계산 어디에도 안 남고 '남은 돈'만
 * 부풀어 오른다 — 초록색 숫자를 믿고 더 쓰게 되는, 가장 나쁜 방향의 오차다.
 *
 * 여러 줄이 걸리면 적어둔 금액에 가장 가까운 한 건을 고른다. 그게 그 고정비다.
 * 한 번 고른 줄은 다시 안 고른다 — 두 고정비가 같은 결제를 나눠 가지면 또 어긋난다.
 */
function fixedChargeRows_(rows, fixedRows, detected) {
  const taken = {};
  const claim = (name, amount, matches) => {
    let best = -1, gap = Infinity;
    rows.forEach((r, i) => {
      if (taken[i] || !matches(r)) return;
      const d = Math.abs(r.amount - amount);
      if (d < gap) { gap = d; best = i; }
    });
    if (best !== -1) taken[best] = name;
  };

  // 시트에 적은 카드 고정비. 통장에서 나가는 것은 애초에 카드 내역에 없다.
  fixedRows.filter(f => f.byCard)
    .forEach(f => claim(f.name, f.amount, r => linkedToFixed_(r, f)));
  // 시트에 없어서 자동으로 잡힌 것. fixedTotal 에 더해진 것과 같은 목록이라야
  // '예산에서 뺀 것'과 '변동비에서 뺀 것'이 정확히 맞아떨어진다.
  Object.keys(detected).forEach(m => claim(m, detected[m], r => r.merchant === m));

  return taken;
}

function readSettings_() {
  const out = {};
  DEFAULT_SETTINGS.forEach(([k, v]) => { out[k] = v; });
  const st = book_().getSheetByName(SETTINGS);
  if (!st) return out;
  st.getRange(2, 1, Math.max(st.getLastRow() - 1, 1), 2).getValues()
    .forEach(([k, v]) => { if (k !== '' && v !== '') out[String(k).trim()] = v; });
  return out;
}

/**
 * 시트의 날짜시간 칸을 'yyyy-MM-dd HH:mm' 문자열로 통일한다.
 * 시트가 문자열을 날짜 값으로 자동 변환해버리는 경우가 있어 양쪽을 모두 받는다.
 */
function toDTString_(v) {
  return (v instanceof Date) ? fmtDT_(v) : String(v || '');
}

/**
 * 헤더 이름 → 0부터 세는 열 번호.
 * 위치로 읽으면 시트에서 열 하나만 옮겨도 값이 통째로 어긋난다.
 */
function headerIndex_(ws) {
  const width = Math.max(ws.getLastColumn(), HEADERS.length);
  const idx = {};
  ws.getRange(1, 1, 1, width).getValues()[0].forEach((h, i) => {
    const name = String(h || '').trim();
    if (name && idx[name] === undefined) idx[name] = i;
  });
  return idx;
}

/**
 * 화면 계산이 거슬러 올라가는 가장 먼 지점.
 *
 *   지난달 1일        카테고리 비교·카드 청구서 — 말일이면 61일 전
 *   고정비 자동 감지  마지막 결제 40일(FIXED_FRESH_DAYS) + 그 전 결제 35일 = 75일
 *   예산 주기         Cycle length 상한이 62일
 *
 * 가장 먼 것이 75일이라 120일로 넉넉히 둔다. 이 창을 좁히면 고정비가 조용히
 * 안 잡히고 '남은 돈'이 그만큼 부풀어 오른다 — 틀린 줄도 모르는 쪽으로 틀린다.
 */
const VIEW_WINDOW_DAYS = 120;

// 이보다 적으면 통째로 읽는다. 창을 쓰려면 Date 열을 따로 한 번 읽어야 하는데,
// 행이 적을 때는 그 왕복이 아껴지는 것보다 비싸다.
const WINDOW_MIN_ROWS = 300;

/** 화면이 읽어올 가장 이른 날짜. loadRecords_(since) 에 그대로 넘긴다. */
function viewWindowStart_(now) {
  const d = now || new Date();
  return fmtD_(new Date(d.getFullYear(), d.getMonth(), d.getDate() - VIEW_WINDOW_DAYS));
}

/**
 * Ledger 시트를 레코드 배열로 읽는다.
 *
 * since('yyyy-MM-dd')를 주면 그날 이후 행만 읽는다. 화면은 최근 몇 달만 쓰는데
 * 시트는 계속 쌓이므로, 그냥 읽으면 쓰지도 않을 작년 행까지 매번 실어 온다 —
 * 행이 늘수록 대시보드와 결제 저장이 같이 느려진다.
 *
 * 좁히는 방법은 Date 열만 먼저 읽어(한 열이라 11분의 1 무게) 시작 위치를 찾고,
 * 거기서부터 전체 열을 읽는 것이다. '마지막 N 행'으로 어림잡지 않는다 — 잘못
 * 잡으면 고정비가 안 잡히고 숫자가 조용히 틀린다. 시트가 정렬돼 있지 않더라도
 * '창 안에 드는 첫 행'부터 끝까지 통째로 읽으므로 빠지는 행이 없다.
 *
 * 학습(learnedCategory_·learnedTag_)은 전체 이력이 필요하므로 since 없이 부른다.
 *
 * 여기서 ensureSheets_() 를 부르면 안 된다. 화면을 열 때마다 시트 4개를 훑고
 * 열·설명·헤더를 전부 대조하게 되는데, 구글시트 호출은 한 번에 수백 ms 라
 * 데이터를 읽기도 전에 몇 초가 날아간다. 시트 손질은 쓰기 경로(doPost)와
 * 시트를 열 때(onOpen)에만 한다.
 */
function loadRecords_(since) {
  const ws = book_().getSheetByName(SHEET);
  if (!ws) return [];
  const last = ws.getLastRow();
  if (last < 2) return [];

  const idx = headerIndex_(ws);
  const at = (r, h) => (idx[h] === undefined ? '' : r[idx[h]]);
  const width = Math.max(ws.getLastColumn(), HEADERS.length);
  const rows = last - 1;

  // 읽을 범위를 창 안으로 좁힌다. Date 열만 먼저 보고 시작 행을 찾는다.
  let from = 2, count = rows;
  if (since && rows >= WINDOW_MIN_ROWS) {
    const dcol = (idx['Date'] === undefined ? 0 : idx['Date']) + 1;
    const dates = ws.getRange(2, dcol, rows, 1).getValues();
    let first = rows;
    for (let i = 0; i < rows; i++) {
      const d = toDTString_(dates[i][0]).slice(0, 10);
      if (d && d >= since) { first = i; break; }
    }
    if (first >= rows) return [];        // 창 안에 아무것도 없다
    from = 2 + first;
    count = rows - first;
  }

  return ws.getRange(from, 1, count, width).getValues()
    .filter(r => at(r, 'Date'))
    .map(r => ({
      dtStr: toDTString_(at(r, 'Date')),
      date: toDTString_(at(r, 'Date')).slice(0, 10),
      amount: Number(at(r, 'Amount')) || 0,
      category: at(r, 'Category') || 'Uncategorized',
      merchant: String(at(r, 'Merchant') || ''),
      tag: String(at(r, 'Tag') || '').trim(),
      merchantEn: String(at(r, 'MerchantEN') || '').trim(),
      // 저장하던 그때 규칙이 내놓은 값. Category 와 다르면 사람이 고친 것이다.
      autoCategory: String(at(r, 'AutoCategory') || '').trim(),
      card: String(at(r, 'Card') || '').trim(),     // 카드사별로 청구액을 나누는 데 쓴다
      foreign: String(at(r, 'Foreign') || '').trim(),   // 해외 취소의 짝을 찾는 데 쓴다
    }));
}

/**
 * 같은 가맹점에 붙은 태그를 물려받는다(학습).
 *
 * 단, **같은 태그가 2번 이상 일치할 때만**. 한 번 본 걸로 단정하면
 * 스타벅스를 데이트로 한 번 태그했다는 이유로 혼자 마신 커피까지
 * 데이트가 되어버린다. 갈 때마다 성격이 달라지는 가맹점은 태그가
 * 엇갈리므로 영영 자동으로 붙지 않는다 — 그게 맞는 동작이다.
 */
/**
 * 같은 가맹점의 카테고리를 물려받는다(학습).
 *
 * 물려받을 것은 '사람이 고친 값'뿐이다. 그 판별은 같은 행에 적어둔 AutoCategory —
 * 저장하던 **그때** 규칙이 내놓은 값 — 과 비교해서 한다.
 *
 * 오늘의 규칙과 비교하면 안 된다. 규칙을 한 번 고치는 순간 그 가맹점의 과거 행이
 * 전부 '사람이 고친 것'으로 둔갑해, 새 결제에 옛 카테고리를 도로 물려준다.
 * 카카오T를 교통에서 택시로 옮겨도 계속 교통으로 쌓이던 것이 이것 때문이다.
 *
 * AutoCategory 가 빈 행(이 칸이 생기기 전에 저장된 것)은 판별할 근거가 없으니 세지 않는다.
 * 'Clean up ledger' 가 한 번 채워준다.
 *
 * 한 번만 고쳐도 바로 물려받는다. 태그가 2번을 요구하는 것과 다른데, 카테고리는
 * 가맹점의 성질이라 방문마다 달라지지 않기 때문이다. '모바일티머니선'은 언제 긁어도 교통이다.
 *
 * 여러 값이 섞여 있으면 **가장 나중에 고친 값**을 따른다. 많이 쓰인 쪽이 아니다 —
 * 옛 행이 스무 줄 쌓여 있으면 오늘 고친 한 줄이 스무 표에 밀려, 고쳐도 고쳐도
 * 안 바뀌는 것처럼 보인다. 마음이 바뀌면 최근 것이 지금의 뜻이다.
 * 레코드는 날짜순으로 들어오므로 뒤에 있을수록 최근이다.
 */
function learnedCategory_(records, merchant) {
  let best = '', at = -1;
  records.forEach((r, i) => {
    if (r.merchant !== merchant) return;
    const cat = String(r.category || '').trim();
    const auto = String(r.autoCategory || '').trim();
    if (!cat || !auto || cat === auto) return;   // 규칙 그대로면 배울 게 없다
    if (i >= at) { best = cat; at = i; }
  });
  return best;
}

function learnedTag_(records, merchant) {
  const tags = records.filter(r => r.merchant === merchant && r.tag).map(r => r.tag);
  if (tags.length < 2) return '';
  return tags.every(t => t === tags[0]) ? tags[0] : '';
}

/**
 * 레코드를 시트에 추가한다. 중복이면 건너뛴다.
 */
function appendRecord_(rec) {
  const ws = ensureSheets_();
  const existing = loadRecords_();
  const dtStr = fmtDT_(rec.dt);

  // 외화는 원화 금액이 그때 환율이라 같은 문자를 다시 보내면 몇 원 달라진다.
  // 금액 대신 외화 금액으로 견줘야 다시 보낸 문자가 두 줄로 쌓이지 않는다.
  const fkey = foreignKey_(rec.foreignNote);
  const dup = existing.some(r =>
    r.dtStr === dtStr && r.merchant === rec.merchant &&
    (fkey ? foreignKey_(r.foreign) === fkey && (r.amount < 0) === (rec.amount < 0)
          : r.amount === rec.amount));
  if (dup) return false;

  // 값은 헤더 이름을 보고 자리를 찾는다 — 시트에서 열 순서를 바꿔도 그대로 동작한다
  const idx = headerIndex_(ws);
  const width = Math.max(ws.getLastColumn(), HEADERS.length);
  const row = new Array(width).fill('');
  const put = (h, v) => { if (idx[h] !== undefined) row[idx[h]] = v; };

  // 규칙이 내놓은 값이 기본. 그 가맹점을 손으로 고친 적이 있으면 그쪽을 따른다.
  const autoLabel = CATEGORY_LABEL[rec.category] || rec.category;
  const label = learnedCategory_(existing, rec.merchant) || autoLabel;

  put('Date', dtStr);
  put('Amount', rec.amount);
  put('Category', label);
  put('Merchant', rec.merchant);
  put('MerchantEN', toEnglish_(rec.merchant));   // 저장할 때 한 번만 번역해 둔다
  put('AutoCategory', autoLabel);                // 나중에 '사람이 고친 것'을 가려내는 기준
  put('Tag', learnedTag_(existing, rec.merchant));
  put('Card', rec.card);
  put('Last4', rec.last4);
  put('Payment', rec.paytype);
  put('Kind', rec.kind);
  put('Foreign', rec.foreignNote || '');
  ws.appendRow(row);

  normalizeDatesAndSort_(ws);   // 날짜 형식을 통일하고 날짜순으로 정렬
  return true;
}


// ══════════════════════════════════════════════════════════════════ 계산

// 마지막 결제가 이만큼 지났으면 이미 끊은 것으로 본다.
// 한 달 주기가 한 번 걸러지면 바로 사라지도록 40일로 둔다.
const FIXED_FRESH_DAYS = 40;

/**
 * 같은 가맹점이 26~35일 간격으로 15% 이내 금액 반복 → 고정비. {가맹점: 최근금액}
 *
 * 마지막 결제가 40일을 넘었으면 빼놓는다. 예전에는 전체 기록을 보고 한 번이라도
 * 그런 쌍이 있으면 영원히 고정비로 뒀다. 작년 1~2월에 두 번 간 미용실이 그 뒤로
 * 발길을 끊어도 매달 16만원씩 예산을 깎았고, 그 가게 지출은 변동비에서도 빠져
 * '남은 돈'이 실제보다 커 보였다. 화면에는 'Limit = ...' 줄에 이름 석 자만 스칠 뿐이라
 * 왜 예산이 줄었는지 알아낼 방법이 사실상 없었다.
 */
function detectFixed_(records) {
  const byMerchant = {};
  records.filter(r => r.amount > 0).forEach(r => {
    (byMerchant[r.merchant] = byMerchant[r.merchant] || []).push(r);
  });

  const now = new Date();
  const fixed = {};
  Object.keys(byMerchant).forEach(merchant => {
    const rows = byMerchant[merchant].sort((a, b) => a.dtStr.localeCompare(b.dtStr));
    const last = rows[rows.length - 1];
    // 아직 살아 있는 지출인가부터 본다. 아니면 간격을 볼 것도 없다.
    if ((now - new Date(last.date)) / 86400000 > FIXED_FRESH_DAYS) return;

    for (let i = 1; i < rows.length; i++) {
      const gap = (new Date(rows[i].date) - new Date(rows[i - 1].date)) / 86400000;
      const base = Math.max(rows[i].amount, rows[i - 1].amount) || 1;
      if (gap >= 26 && gap <= 35 &&
          Math.abs(rows[i].amount - rows[i - 1].amount) / base <= 0.15) {
        fixed[merchant] = last.amount;
        break;
      }
    }
  });
  return fixed;
}

function weekStartIndex_(settings) {
  const raw = String(settings[WEEKSTART_KEY] || 'Mon').trim();
  const en = WEEKDAY_EN.findIndex(d => d.toLowerCase() === raw.slice(0, 3).toLowerCase());
  if (en !== -1) return en;
  const ko = WEEKDAYS_KO.indexOf(raw.charAt(0));   // '월' 처럼 한글로 적어도 받아준다
  return ko === -1 ? 0 : ko;
}

/** 월=0 … 일=6 기준으로 그 날이 속한 주의 시작일. */
function weekStartOf_(d, startIndex) {
  const mondayBased = (d.getDay() + 6) % 7;
  const back = (mondayBased - startIndex + 7) % 7;
  const s = new Date(d.getFullYear(), d.getMonth(), d.getDate() - back);
  return s;
}

function midnight_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

/**
 * 이번 예산 주기의 시작일과 길이.
 *
 * 'Month' 로 적으면 달력 한 달을 통째로 쓴다 — 1일부터 말일까지. 달마다 날 수가
 * 달라도 예산은 한 달치 전부이므로 길이로 나누지 않는다(monthly 플래그).
 * 기준일(Cycle starts on)이 있으면 거기서부터 length 일씩 끊는다 — 요일과 무관하게
 * 정확히 반복되므로 10일 주기 같은 것도 어긋나지 않는다.
 * 기준일이 없고 7일 주기면 요일(Week starts on)로 잡는다.
 */
function cycleOf_(settings, today) {
  const rawLength = String(settings[CYCLE_DAYS_KEY] || '').trim().toLowerCase();
  if (rawLength.indexOf('month') === 0 || rawLength.charAt(0) === '월') {
    const start = new Date(today.getFullYear(), today.getMonth(), 1);
    const end = new Date(today.getFullYear(), today.getMonth() + 1, 0);   // 다음 달 0일 = 이번 달 말일
    return {start: start, end: end, length: end.getDate(), monthly: true};
  }

  let length = Math.floor(Number(settings[CYCLE_DAYS_KEY]) || 7);
  if (!(length >= 1 && length <= 62)) length = 7;

  const rawAnchor = String(settings[CYCLE_START_KEY] || '').trim()
    ? new Date(toDTString_(settings[CYCLE_START_KEY]).slice(0, 10))
    : null;
  const anchor = rawAnchor && !isNaN(rawAnchor.getTime()) ? midnight_(rawAnchor) : null;

  let start;
  if (anchor) {
    const elapsed = Math.floor((midnight_(today) - anchor) / 86400000);
    const done = Math.floor(elapsed / length);          // 음수(기준일이 미래)여도 맞게 떨어진다
    start = new Date(anchor.getFullYear(), anchor.getMonth(), anchor.getDate() + done * length);
  } else if (length === 7) {
    start = weekStartOf_(today, weekStartIndex_(settings));
  } else {
    // 기준일을 안 적었는데 7일도 아니면 고정된 기준점에서 끊는다.
    //
    // 예전에는 이번 달 1일에서 끊었다. 그러면 달마다 리셋되어, 10일 주기일 때 8/31 에
    // 하루짜리 주기가 새로 시작하면서 10일치 한도를 통째로 받고, 9/1 에 또 새 주기가
    // 시작해 같은 며칠에 두 번 한도를 줬다. 8/31 에 쓴 돈은 9/1 이 되면 변동비에서도
    // 빠졌다. 고정된 기준점에서 세면 주기가 달 경계와 무관하게 이어진다.
    const epoch = new Date(2000, 0, 1);
    const done = Math.floor((midnight_(today) - epoch) / 86400000 / length);
    start = new Date(epoch.getFullYear(), epoch.getMonth(), epoch.getDate() + done * length);
  }

  const end = new Date(start.getFullYear(), start.getMonth(), start.getDate() + length - 1);
  return {start: start, end: end, length: length, monthly: false};
}

/** 이번 주 현황. 파이썬 summarize_week()와 같은 값을 낸다. */
function summarizeWeek_() {
  const now = new Date();
  // 화면이 실제로 쓰는 범위만 읽는다 — 시트가 쌓여도 느려지지 않도록.
  // 창의 근거는 VIEW_WINDOW_DAYS 주석에.
  const records = loadRecords_(viewWindowStart_(now));
  const settings = readSettings_();

  const cycle = cycleOf_(settings, now);
  const start = cycle.start, end = cycle.end;
  const startStr = fmtD_(start), endStr = fmtD_(end);

  const budget = Number(settings[BUDGET_KEY]) || 0;

  // 고정비: Fixed 시트가 우선, 거기 없는 것만 자동 감지로 보충한다.
  const allRows = readFixed_();
  const fixedRows = allRows.filter(f => f.every === 'month');      // 월 단위 → 월 예산에서 차감
  const recurRows = allRows.filter(f => f.every !== 'month');      // 주/일 단위 → 이번 주에서 예약
  const detectedAll = detectFixed_(records);
  const detected = {};                       // 목록에 없어서 자동으로 잡은 것만
  Object.keys(detectedAll).forEach(m => {
    if (!matchesFixed_(m, fixedRows)) detected[m] = detectedAll[m];
  });

  let fixedTotal = fixedRows.reduce((s, f) => s + f.amount, 0)
                 + Object.keys(detected).reduce((s, k) => s + detected[k], 0);

  // Fixed 시트가 비어 있을 때만 Settings 의 수동 총액을 쓴다
  const manual = Number(settings[FIXED_KEY]) || 0;
  if (!fixedRows.length && manual > 0) fixedTotal = manual;

  // 이번 달 환급·캐시백은 그 달 예산에 더해져 남은 주에 고르게 퍼진다
  const monthPrefix = Utilities.formatDate(now, TZ, 'yyyy-MM');
  const extraRows = readExtra_().filter(x => x.date.slice(0, 7) === monthPrefix);
  const extraTotal = extraRows.reduce((s, x) => s + x.amount, 0);

  // 주기 한도. Cycle budget 을 직접 적었으면 그게 곧 자유롭게 쓸 돈이다
  // (고정비는 이미 빠진 금액이라는 뜻이므로 또 빼지 않는다).
  // 안 적었으면 월 예산에서 고정비를 빼고 주기 길이만큼 잘라 쓴다.
  const cycleBudget = Number(settings[CYCLE_BUDGET_KEY]) || 0;
  // 월 단위면 한 달 예산을 통째로 쓴다. 달마다 날 수가 달라도 예산은 그대로이므로
  // 길이로 환산하면 안 된다 — 31일 달에 예산이 1.8% 늘어나는 엉뚱한 일이 생긴다.
  const share = cycle.monthly ? 1 : cycle.length / DAYS_PER_MONTH;
  // 예산만 보고 낸 한도. '계획상 써도 되는 돈'이다.
  const planAllow = cycleBudget > 0
    ? Math.round(cycleBudget + extraTotal * share)
    : (budget ? Math.round((budget + extraTotal - fixedTotal) * share) : 0);

  // ── 통장이 감당 못 하는 만큼 한도를 깎는다 ───────────────────────────────
  // 예산은 계획이고 잔액은 현실이다. 둘을 따로 두면 통장이 마이너스인데도 화면은
  // "이번 주 ₩310,000 쓸 수 있다"고 말한다 — 계획상 맞는 말이지만 실제로 쓸 수가 없다.
  //
  // 현재 잔액을 그대로 빼지는 않는다. 25일에 월급이 들어온다면 지금 마이너스인 게
  // 사고가 아니기 때문이다. forecast_ 가 이미 급여·카드 결제일·고정비를 하루씩 굴려
  // 앞으로 35일의 최저점을 내놓으므로, 그 최저점이 바닥(마이너스 통장 한도)을 뚫고
  // 내려가는 만큼만 깎는다. 한도 안에서 노는 마이너스 통장은 건드리지 않는다 —
  // 그건 이미 forecast_ 가 사고가 아니라고 보는 구간이다(inOverdraft vs negative).
  //
  // 순환이 아니다. forecast_ 의 하루치는 이 한도가 아니라 '월 예산 ÷ 30'에서 나온다.
  // 계획대로 썼을 때 통장이 어디까지 파이는지를 먼저 재고, 그 결과로 계획을 깎는
  // 한 방향이다. 잔액을 안 적었으면(forecast_ 가 null) 아무것도 깎지 않는다.
  const fc = forecast_(settings, fixedRows, fixedTotal, cycle, records);
  const shortfall = fc ? Math.max(0, -fc.overdraft - fc.min) : 0;
  const weeklyAllow = Math.max(0, planAllow - shortfall);

  const thisWeek = records.filter(r => r.date >= startStr && r.date <= endStr);
  const charged = thisWeek.reduce((s, r) => s + r.amount, 0);

  // 카드값은 한 달에 한 번 빠져나간다 → 다음 결제일에 나갈 돈은 '이번 달 승인 총액'이다.
  // 예산 주기(주)와 결제 주기(월)가 다르므로 둘을 따로 센다.
  const monthCharged = records
    .filter(r => r.date.slice(0, 7) === monthPrefix)
    .reduce((s, r) => s + r.amount, 0);

  // ── 이번 달 저금 가능액 ─────────────────────────────────────────────────
  // 월급에서 이번 달에 실제로 나간 돈을 뺀 값. 예산을 넘겨 쓰면 그만큼 저절로 줄어든다.
  // 카드를 안 거친 지출(월세 같은 은행이체, Extra 의 마이너스)까지 빼야 실제와 맞는다.
  // 카드로 나간 고정비는 이미 monthCharged 에 들어 있으므로 여기서 또 더하지 않는다.
  const income = Number(settings[INCOME_KEY]) || 0;
  // 월세·보험처럼 월 단위로 나가는 것 + 청소비·택시처럼 주/일 단위로 나가는 것.
  // 뒤엣것을 빼먹고 있었다. fixedRows 는 월 단위만 걸러낸 목록이라, 현금으로 주는
  // 주 5만원 청소비가 한 달 20만원인데도 monthOut 에 한 푼도 안 들어갔다.
  // 그러면 '이번 달 저금 가능액'이 그만큼 부풀고 예산 초과 경고도 늦게 뜬다.
  // 더 나쁜 건 Monthly 시트다 — Off-card 와 Total out 이 낮게 굳고 되돌릴 수 없다.
  //
  // fixedTotal 에는 넣지 않는다. 주/일 단위는 주기 안에서 reserved 로 잡아두므로
  // 여기서 또 월 예산에서 빼면 같은 돈을 두 번 빼는 셈이다.
  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
  const monthEnd = new Date(now.getFullYear(), now.getMonth() + 1, 0);
  const offCardFixed = fixedRows.filter(f => !f.byCard).reduce((s, f) => s + f.amount, 0)
    + recurRows.filter(f => !f.byCard)
        .reduce((s, f) => s + f.amount * unitsBetween_(f.every, monthStart, monthEnd), 0);
  const extraIn = extraRows.filter(x => x.amount > 0).reduce((s, x) => s + x.amount, 0);
  const extraOut = extraRows.filter(x => x.amount < 0).reduce((s, x) => s - x.amount, 0);

  const monthOut = monthCharged + offCardFixed + extraOut;
  const monthIn = income + extraIn;
  const savings = monthIn - monthOut;

  // 통장이 마이너스면 그걸 메우기 전까지는 저금이 아니다. 화면의 'Could save' 에서
  // 그 구멍만큼 뺀다.
  //
  // savings 자체는 손대지 않는다. Monthly 시트의 Saved 열이 이 값을 쓰는데, 거기는
  // 그 달의 흐름(들어온 돈 − 나간 돈)을 남기는 영구 기록이다. 재고인 마이너스 잔액을
  // 섞으면 구멍이 메워지지 않는 동안 달마다 같은 −800,000 이 거듭 차감되고, 한 번
  // 굳은 지난 달 줄은 되돌릴 수도 없다.
  //
  // 한도는 보지 않는다. 마이너스 통장 한도가 넉넉해도 −800,000 은 갚아야 할 빚이다.
  // (한도는 '얼마까지 파여도 사고가 아닌가'를 말할 뿐 빚이 아니라는 뜻이 아니다.)
  const balanceHole = Math.max(0, -(Number(settings[BALANCE_KEY]) || 0));
  const savingsNet = savings - balanceHole;

  // 예산 초과는 '자유롭게 쓸 돈'이 아니라 이번 달 총지출 기준으로 본다.
  const overBudget = budget > 0 && monthOut > budget;

  // ── 이번 달 카테고리별 ───────────────────────────────────────────────────
  // 예산 주기와 무관하게 늘 '달력 한 달'을 본다 — 물어보는 건 언제나 '이번 달'이라서.
  //
  // 카드 지출만 센다. 통장에서 바로 나가는 고정비(이자·보험·곗돈)는 매달 같은 금액이라
  // '어디에 썼나'를 볼 때 알 게 없고, 덩치가 커서 다른 항목을 다 눌러버린다.
  // 총액은 아래 'Limit = ...' 줄에 따로 나오므로 정보가 사라지지도 않는다.
  //
  // 지난달과 나란히 놓는 자리라 '같은 날짜까지'만 센다. 달 전체와 견주면 월초마다
  // "지난달보다 훨씬 덜 썼다"가 나와 아무 쓸모가 없다. 컷은 두 가지를 지킨다.
  //  - 지난달 말일을 넘지 않는다. 3/30 에 '30일까지'로 자르면 2월은 28일이라 통째로
  //    들어와, 30일치와 28일치를 '같은 날짜까지'라며 견주게 된다.
  //  - 양쪽에 똑같이 건다. 한쪽만 자르면 그 차이가 그대로 거짓 증감이 된다.
  const prevPrefix = monthKey_(now.getFullYear(), now.getMonth() - 1);
  const prevMonthDays = new Date(now.getFullYear(), now.getMonth(), 0).getDate();
  const dayCut = Math.min(now.getDate(), prevMonthDays);
  const dayOf = r => Number(r.date.slice(8, 10));

  const catRows = records.filter(r =>
    r.date.slice(0, 7) === monthPrefix && dayOf(r) <= dayCut);
  const prevRows = records.filter(r =>
    r.date.slice(0, 7) === prevPrefix && dayOf(r) <= dayCut);

  // 비교용 총액. monthCharged(달 전체)는 카드값·저금 계산에 쓰는 값이라 여기 쓰면 안 된다.
  const monthCut = catRows.reduce((s, r) => s + r.amount, 0);
  const prevCharged = prevRows.reduce((s, r) => s + r.amount, 0);

  const catKey = r => String(r.category || '').trim() || 'Uncategorized';
  const catTotals = {};
  catRows.forEach(r => { catTotals[catKey(r)] = (catTotals[catKey(r)] || 0) + r.amount; });

  const prevCat = {};
  prevRows.forEach(r => { prevCat[catKey(r)] = (prevCat[catKey(r)] || 0) + r.amount; });

  // Monthly 시트에 남길 목록은 달 전체로 따로 센다. 위의 컷은 '지난달과 견주기' 위한
  // 것이라, 말일에 찍는 영구 기록에 쓰면 마지막 며칠이 빠진 채로 굳는다 — 3·5·7·10·12월
  // 31일에는 컷이 28~30이다. 금액 열은 달 전체(monthCharged)인데 카테고리만 잘려 있으면
  // 그 한 줄이 스스로 앞뒤가 안 맞는다.
  const catFull = {};
  records.filter(r => r.date.slice(0, 7) === monthPrefix)
    .forEach(r => { catFull[catKey(r)] = (catFull[catKey(r)] || 0) + r.amount; });
  const categoriesFull = Object.keys(catFull)
    .map(k => ({name: k, amount: catFull[k]}))
    .filter(c => c.amount > 0)
    .sort((a, b) => b.amount - a.amount);

  // 지난달에는 있었는데 이번 달엔 아직 안 쓴 카테고리도 보여준다 (0으로 줄었다는 정보)
  Object.keys(prevCat).forEach(k => { if (!(k in catTotals)) catTotals[k] = 0; });

  const categories = Object.keys(catTotals)
    .map(k => ({name: k, amount: catTotals[k], prev: prevCat[k] || 0}))
    .filter(c => c.amount !== 0 || c.prev !== 0)
    .sort((a, b) => b.amount - a.amount);

  // ── 가맹점 Top ───────────────────────────────────────────────────────────
  // 카테고리보다 구체적이라 행동을 바꾸기 쉽다. 'Delivery ₩320,000' 은 막연하지만
  // '배달의민족 12건 ₩320,000' 은 바로 줄일 결심이 선다. 건수도 같이 센다.
  //
  // 취소도 같이 센다. 승인만 세면 20만원 긁고 바로 취소한 가게가 20만원 쓴 것처럼
  // Top 에 올라앉는다. 합이 0 이하면 목록에서 뺀다.
  //
  // 이쪽은 달 전체다 — 위 카테고리 목록은 지난달과 견주려고 날짜를 잘라둔 값이라,
  // 말일 근처(3·5·7·10·12월 29~31일)에는 둘의 합계가 며칠치만큼 어긋난다.
  const merchAgg = {};
  records.filter(r => r.date.slice(0, 7) === monthPrefix).forEach(r => {
    // 묶는 건 원문 이름으로, 보여주는 건 번역으로. 번역이 겹쳐도 서로 섞이지 않는다.
    const key = r.merchant || '(unknown)';
    const m = merchAgg[key] || (merchAgg[key] = {name: r.merchantEn || key, amount: 0, count: 0});
    m.amount += r.amount;
    if (r.amount > 0) m.count++;         // 건수는 실제로 산 횟수
  });
  const merchants = Object.keys(merchAgg).map(k => merchAgg[k])
    .filter(m => m.amount > 0)
    .sort((a, b) => b.amount - a.amount).slice(0, 5);

  // ── 이번 달 착지 예상 ────────────────────────────────────────────────────
  // '남은 돈'만 보면 이대로 가면 어떻게 끝나는지가 안 보인다. 월초에 알아야 조절이 된다.
  // 카드값만 페이스로 늘린다 — 통장 고정비는 이미 한 달치라 늘리면 두 번 세는 셈이다.
  //
  // 페이스는 오늘 날짜로 나눈다 — dayCut(지난달 말일에 걸린 비교용 컷)이 아니다.
  // 3/30 에 dayCut 은 28 이라, 그걸로 나누면 이틀치를 없는 셈 치고 페이스가 부풀어 오른다.
  const dayOfMonth = now.getDate();
  const daysInMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
  const monthDaysLeft = Math.max(daysInMonth - dayOfMonth + 1, 0);   // 오늘 포함
  const projectedCard = Math.round(monthCharged / dayOfMonth * daysInMonth);
  const projectedOut = projectedCard + offCardFixed + extraOut;
  const projection = {
    // 1~2일차엔 한 건이 31배로 부풀어 헛소리가 된다. 사흘은 지나야 쓸 만하다.
    show: dayOfMonth >= 3 && monthCharged > 0,
    card: projectedCard,
    out: projectedOut,
    over: budget > 0 ? projectedOut - budget : 0,
    daysLeft: monthDaysLeft,
    // 예산에 맞추려면 남은 날 하루에 얼마까지 쓸 수 있나
    perDay: budget > 0 && monthDaysLeft > 0
      ? Math.round((budget - monthOut) / monthDaysLeft) : 0,
  };

  // 막대 길이는 이번 달과 지난달을 통틀어 가장 큰 항목을 100으로 잡는다.
  // 그래야 이번 달 막대와 지난달 옅은 막대의 길이를 그대로 견줄 수 있다 —
  // 두 달을 각자의 최대값으로 재면 길이가 서로 다른 걸 재게 되어 비교가 거짓이 된다.
  const catScale = categories.reduce((m, c) => Math.max(m, c.amount, c.prev), 0) || 1;
  categories.forEach(c => {
    c.bar = Math.max(0, Math.round(c.amount / catScale * 100));
    c.prevBar = Math.max(0, Math.round(c.prev / catScale * 100));
    // 비율은 같은 컷의 카드 총액 기준이다. 목록에 없는 통장 이체까지 분모에 넣으면
    // 퍼센트를 다 더해도 100이 안 돼서 읽는 사람이 헷갈린다.
    //
    // 막대(가장 큰 항목 대비)와 기준이 다르므로 같은 줄에 나란히 두지 않는다.
    // 예전에는 막대가 꽉 찼는데 옆 숫자가 38% 라, 한 줄에서 길이와 숫자가 서로
    // 다른 걸 재고 있었다. 이 값은 아래 보조 줄로 내린다.
    //
    // 지난달 산 걸 이번 달에 취소하면 그 항목은 음수가 된다. 막대와 똑같이 0에서 끊는다.
    c.share = monthCut > 0 ? Math.max(0, Math.round(c.amount / monthCut * 100)) : 0;
    c.delta = c.amount - c.prev;
  });
  // 변동비에서 빼는 건 '월 예산에서 이미 뺀 것'뿐이다 — Fixed 시트에 적었거나
  // 자동 감지된 항목. 카테고리가 구독/통신이라는 이유만으로 빼면 안 된다.
  // (넷플릭스 첫 결제처럼 아직 고정비로 안 잡힌 건 월 예산에서도 안 뺐으므로,
  //  여기서까지 빼버리면 그 돈이 계산 어디에도 안 남는다)
  // 금액만큼 한 건씩만 뺀다 — 자세한 이유는 fixedChargeRows_ 주석에.
  const takenRows = fixedChargeRows_(thisWeek, fixedRows, detected);
  const variable = thisWeek
    .filter((r, i) => !takenRows[i])
    .reduce((s, r) => s + r.amount, 0);

  // 어느 고정비가 어떤 결제를 가져갔는지 그대로 보여준다.
  // 이름 맞추기가 헐거워서 'GS' 한 줄이 GS25 편의점을 집을 수 있는데, 화면에
  // 안 나오면 남은 돈이 왜 이런지 알아낼 방법이 없다. 여기 이상한 줄이 보이면
  // Fixed 시트의 Match 칸에 정확한 상호를 적으라는 신호다.
  const claimed = thisWeek
    .map((r, i) => takenRows[i]
      ? {name: takenRows[i], merchant: r.merchantEn || r.merchant, amount: r.amount}
      : null)
    .filter(x => x);

  const daysLeft = Math.max(
    Math.ceil((end - new Date(now.getFullYear(), now.getMonth(), now.getDate())) / 86400000) + 1, 0);

  // 이번 주에 아직 나갈 돈을 미리 잡아둔다(예약).
  //
  // '남은 횟수 × 하루치' 로 계산한다. 지나간 날은 세지 않으므로,
  // 택시를 안 탄 날이 있으면 그만큼 예약이 줄어 쓸 수 있는 돈이 늘어난다.
  // 실제 결제액은 변동비에 그대로 잡히니, 예상보다 많이 썼으면 남은 돈이 더 줄어든다.
  const todayStr = fmtD_(now);
  const reserveRows = [];
  let reserved = 0;
  recurRows.forEach(f => {
    let units = unitsBetween_(f.every, now > start ? now : start, end);

    // 카드로 결제됐으면 그 회차는 이미 끝난 것 → 예약에서 뺀다.
    // Paid by 와 무관하게 항상 확인한다. 평소 현금으로 주다가 어쩌다 카드로 긁으면
    // 예약과 카드 지출이 둘 다 잡혀 같은 돈이 두 번 빠지기 때문.
    const paid = thisWeek.filter(r => linkedToFixed_(r, f));
    const doneThisUnit = f.every === 'week'
      ? paid.length > 0                              // 주 1회짜리는 이번 주에 찍혔는지
      : paid.some(r => r.date === todayStr);         // 매일/평일은 오늘 찍혔는지
    if (doneThisUnit) units -= 1;

    const rest = f.amount * Math.max(0, units);
    reserved += rest;
    reserveRows.push({
      name: f.name,
      expected: f.amount * unitsBetween_(f.every, start, end),
      remaining: rest,
    });
  });

  const remaining = weeklyAllow - variable - reserved;

  // ── 페이스 ───────────────────────────────────────────────────────────────
  // '얼마 썼나'만으로는 그게 빠른 건지 알 수 없다. 화요일에 한도의 40%를 썼다면
  // 7일 주기에서는 과속이고 3일 주기에서는 오히려 아끼는 중이다. 주기에서 지나간
  // 몫만큼을 '오늘까지 써도 됐던 돈'으로 잡아, 막대에 눈금 하나로 찍는다.
  //
  // 오늘을 포함해 센다 — 오늘 하루치는 오늘 다 써도 되는 돈이기 때문이다.
  const elapsedShare = cycle.length > 0
    ? Math.min(1, Math.max(0, (cycle.length - daysLeft + 1) / cycle.length))
    : 1;
  // 견주는 대상은 변동비다. 예약(reserved)은 아직 안 나간 돈이라 막대에 따로
  // 옅게 그려져 있고, 여기 섞으면 '이미 썼다'는 뜻이 되어 두 번 세는 셈이 된다.
  const paceSpend = Math.round(weeklyAllow * elapsedShare);
  const pace = {
    share: Math.round(elapsedShare * 100),
    spend: paceSpend,
    vs: paceSpend - variable,        // + 면 페이스보다 덜 쓴 것
    // 예산이 잡혀 있으면 보여준다. weeklyAllow 로 보면 안 된다 — 잔액이 모자라 한도가
    // 0으로 깎인 순간(가장 알아야 할 때) 페이스가 화면에서 통째로 사라진다.
    // 한도가 0이면 paceSpend 도 0이라 쓴 돈이 그대로 '초과'로 나오는 게 맞다.
    show: planAllow > 0 && cycle.length > 1,
  };

  return {
    start: start, end: end, cycleDays: cycle.length,
    // 주기를 부르는 말. 10일 주기에 '이번 주'는 거짓말이라 'cycle' 로 둔다.
    periodWord: cycle.monthly ? 'month' : (cycle.length === 7 ? 'week' : 'cycle'),
    budget: budget, fixedTotal: fixedTotal, detected: detected,
    fixedRows: fixedRows, extraRows: extraRows, extraTotal: extraTotal,
    reserveRows: reserveRows, reserved: reserved, claimed: claimed,
    forecast: fc,
    // planAllow 는 예산이 말하는 한도, weeklyAllow 는 통장까지 본 실제 한도다.
    // shortfall 이 그 차액 — 0 이 아니면 화면에 이유를 밝혀야 한다.
    weeklyAllow: weeklyAllow, planAllow: planAllow, shortfall: shortfall,
    charged: charged, monthCharged: monthCharged,
    variable: variable,
    income: income, monthIn: monthIn, monthOut: monthOut,
    // savings 는 이번 달 흐름(Monthly 기록용), savingsNet 은 마이너스 잔액까지 뺀 화면용.
    savings: savings, savingsNet: savingsNet, balanceHole: balanceHole,
    overBudget: overBudget, overBy: overBudget ? monthOut - budget : 0,
    // monthCut 은 prevCharged 와 같은 날짜까지 자른 이번 달 총액이다. 카드값·저금에 쓰는
    // monthCharged(달 전체)와 섞으면 안 된다 — 견줄 때는 반드시 이쪽을 쓴다.
    categories: categories, monthCut: monthCut, prevCharged: prevCharged,
    categoriesFull: categoriesFull,   // Monthly 기록용 — 달 전체
    merchants: merchants, projection: projection,
    remaining: remaining, daysLeft: daysLeft, pace: pace,
    perDay: daysLeft ? Math.round(remaining / daysLeft) : 0,
    rows: thisWeek,
  };
}

/** 12345 → "₩12,345". Intl 의존 없이 직접 찍는다. */
function won_(n) {
  const v = Math.round(Number(n) || 0);
  const digits = Math.abs(v).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (v < 0 ? '-' : '') + '₩' + digits;
}

/** 단축어 알림에 띄울 한 줄 요약. */
function summaryLine_(s) {
  // 예산 기준으로 본다. weeklyAllow 로 보면 잔액이 모자라 한도가 0으로 깎였을 때
  // "예산을 설정하세요"라는 엉뚱한 안내가 나간다 — 예산은 멀쩡히 적혀 있는데도.
  if (!s.planAllow) return `Set "${BUDGET_KEY}" or "${CYCLE_BUDGET_KEY}" in ${SETTINGS}`;
  const save = s.income ? ` · save ${won_(s.savingsNet)}` : '';
  const short = s.shortfall ? ` · account short ${won_(s.shortfall)}` : '';
  return s.remaining < 0
    ? `${won_(-s.remaining)} over this ${s.periodWord}${short} · card bill ${won_(s.monthCharged)}${save}`
    : `${won_(s.remaining)} left this ${s.periodWord}${short} · card bill ${won_(s.monthCharged)}${save}`;
}


// ═══════════════════════════════════════════════════════════════ 엔드포인트

function ok_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function authed_(e) {
  // TOKEN 을 빈 값으로 두면 아무도 들여보내지 않는다. 빈 값끼리 맞아떨어져
  // 통과되면 익명 공개 웹앱이 통째로 열리기 때문이다.
  if (!TOKEN) return false;
  return !!(e && e.parameter && e.parameter.t === TOKEN);
}

/**
 * 이 웹앱의 /exec 주소 + 토큰. 새로고침 링크에 쓴다.
 *
 * 화면은 googleusercontent.com 의 일회용 임시 주소에 뜨는데, 그 주소를 다시 부르면
 * 스크립트가 돌지 않아 빈 화면이 나온다. 반드시 /exec 로 되돌아가야 다시 계산된다.
 */
function selfUrl_() {
  try {
    return ScriptApp.getService().getUrl() + '?t=' + encodeURIComponent(TOKEN);
  } catch (err) {
    return '';        // 주소를 못 얻으면 링크를 그냥 안 보여준다
  }
}

// 시트 왕복 한 번에 300~600ms 가 그냥 깔린다. 화면 한 장에 네 번(Ledger·Settings·
// Fixed·Extra) 이라 캐시가 없으면 매번 2~3초다. 그래서 넉넉히 잡는다.
// 숫자가 오래돼 보일 걱정은 없다 — 결제가 들어오거나(doPost) 시트를 손으로
// 고치면(onEdit) 캐시가 즉시 깨지고, Refresh 링크는 아예 캐시를 건너뛴다.
const VIEW_CACHE_SECONDS = 900;
const VIEW_KEYS = ['view_html', 'view_text'];

/** 캐시된 화면을 버린다. 데이터가 바뀌는 모든 길목에서 부른다. */
function dropViewCache_() {
  try { CacheService.getScriptCache().removeAll(VIEW_KEYS); } catch (err) { /* 없으면 말고 */ }
}

/**
 * 자정까지 남은 초.
 *
 * 화면에는 '남은 며칠'과 주기 범위가 박혀 있다. 캐시가 자정을 넘겨 살아 있으면
 * 새 주기가 시작됐는데 지난 주기 숫자를 보여준다 — 23:58 에 만든 화면이 월요일
 * 아침에 '1 day left' 로 나오는 식이다.
 */
function secondsToMidnight_() {
  const now = new Date();
  const next = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
  return Math.max(60, Math.floor((next - now) / 1000));
}

/**
 * 만들어진 화면을 잠깐 캐시하고, 이번 요청이 얼마나 걸렸는지 화면에 박아 준다.
 *
 * 대시보드 한 번 그리려면 시트를 네 번 읽어야 하고 그게 로딩 시간의 대부분이다.
 * 새로고침을 연달아 누르거나 폰·PC에서 같이 볼 때 같은 계산을 반복할 이유가 없다.
 *
 * 소요 시간은 캐시된 본문에 {{SERVED}} 자리를 비워뒀다가 내보낼 때 채운다.
 * 본문에 같이 캐시해버리면 처음 만들 때 시간이 계속 박혀 나와, 캐시가 먹히는지
 * 아닌지를 화면만 보고는 영영 알 수 없다.
 */
function cachedView_(key, build, fresh) {
  const at = Date.now();
  const cache = CacheService.getScriptCache();
  const hit = fresh ? null : cache.get(key);
  const cached = hit !== null;

  const out = cached ? hit : build();
  if (!cached) {
    // 자정을 넘겨 살아 있지 않도록 잘라 둔다 — 날짜가 바뀌면 화면의 '남은 며칠'과
    // 주기 범위가 거짓이 된다.
    const ttl = Math.min(VIEW_CACHE_SECONDS, secondsToMidnight_());
    // 100KB 를 넘으면 캐시가 거부한다. 그래도 화면은 나가야 하므로 조용히 넘어간다.
    try { cache.put(key, out, ttl); } catch (err) { /* 캐시 없이 진행 */ }
  }

  const stamp = `${Utilities.formatDate(new Date(), TZ, 'HH:mm')} · ` +
                `${Date.now() - at}ms · ${cached ? 'cached' : 'fresh'}`;
  return out.split('{{SERVED}}').join(stamp);
}

/** 아이폰 단축어가 문자 텍스트를 보내는 곳. */
function doPost(e) {
  if (!authed_(e)) return ok_({ok: false, message: 'Not authorized'});

  let text = '';
  if (e.postData && e.postData.contents) {
    try {
      const body = JSON.parse(e.postData.contents);
      text = body.text || body.message || e.postData.contents;
    } catch (err) {
      text = e.postData.contents;      // 순수 텍스트로 보낸 경우
    }
  }
  text = text || (e.parameter && e.parameter.text) || '';

  const rec = parseMessage_(text);
  if (!rec) {
    return ok_({ok: false, message: "Couldn't read this message", raw: String(text).slice(0, 120)});
  }

  // 문자가 연달아 오면 같은 행에 동시에 쓰다 깨질 수 있어 잠근다.
  const lock = LockService.getScriptLock();
  let added, s;
  try {
    lock.waitLock(20000);
    // 외화는 원화로 바꾼 뒤에야 저장할 수 있다. 짝 찾기가 원장을 읽으므로 잠금 안에서 한다.
    if (rec.currency) {
      const skip = settleForeign_(rec);
      if (skip) return ok_({ok: true, added: false, message: skip});
    }
    added = appendRecord_(rec);
    s = summarizeWeek_();
    if (added) {
      updateMonthly_(s);                       // 말일이면 이번 달 줄을 남긴다(아니면 그냥 지나간다)
      dropViewCache_();                        // 새 결제가 반영된 화면을 보여줘야 한다
    }
  } catch (err) {
    // 무엇이 틀렸는지 그대로 알려준다. 예전에는 전부 'Busy' 라고 해서, 저장은 됐는데
    // 요약 계산에서 터진 경우와 진짜 잠금 충돌을 구별할 수 없었다. 앞엣것이면 다시
    // 보내도 중복 검사에 걸려 영영 성공 응답을 못 받는다 — 이미 저장돼 있는데도.
    // 토큰으로 막힌 개인용 엔드포인트라 메시지를 그대로 내보내도 된다.
    return ok_({ok: false, added: !!added,
                message: (added ? 'Saved, but the summary failed: ' : 'Failed: ') +
                         (err && err.message ? err.message : String(err))});
  } finally {
    lock.releaseLock();
  }

  return ok_({
    ok: true,
    added: added,
    // 중복이라 안 저장했으면 그렇다고 말해야 한다. 예전에는 저장된 것과 똑같은 알림이
    // 떠서, 같은 날 같은 금액의 두 번째 결제가 조용히 사라진 걸 알 방법이 없었다.
    // (취소 문자는 시각이 00:00 로 고정이라 이 충돌이 특히 잘 난다.)
    message: added
      ? `${recLabel_(rec)}\n${summaryLine_(s)}`
      : `Already recorded, not saved again\n${recLabel_(rec)}`,
    remaining: s.remaining,
    charged: s.charged,
  });
}

/**
 * 폰에서 여는 대시보드.
 *   ?view=text 를 붙이면 순수 텍스트로 준다 — 단축어에서 바로 띄우는 용도.
 *   iOS가 script.google.com 을 구글 앱으로 가로채 Safari가 안 열릴 때 이 경로를 쓴다.
 */
function doGet(e) {
  const view = e && e.parameter ? e.parameter.view : '';

  if (!authed_(e)) {
    return view === 'text'
      ? ContentService.createTextOutput('Invalid link')
      : HtmlService.createHtmlOutput('<p style="font:16px sans-serif">Invalid link.</p>');
  }

  // ?view=debug — 어디서 시간이 가는지 재서 보여준다.
  // '느리다'는 체감만으로는 우리 코드 탓인지 구글 탓인지 가릴 수 없다.
  if (view === 'debug') return ContentService.createTextOutput(timingReport_());

  // 시트가 아직 없을 때만 만든다. 매번 확인하면 그게 곧 로딩 지연이다.
  if (!book_().getSheetByName(SHEET)) ensureSheets_();

  // Refresh 링크에 붙는다. 캐시를 건너뛰고 지금 값으로 다시 계산한다.
  const fresh = !!(e.parameter && e.parameter.fresh);

  if (view === 'text') {
    return ContentService.createTextOutput(
      cachedView_('view_text', () => summaryText_(summarizeWeek_()), fresh));
  }
  return HtmlService.createHtmlOutput(
    cachedView_('view_html', () => dashboardHtml_(summarizeWeek_()), fresh))
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle('Left to spend');
}

/**
 * 각 단계가 몇 ms 걸리는지 재서 돌려준다.
 *
 * 여기 합계가 작은데 화면이 느리면 남은 시간은 전부 구글 몫이다 — 스크립트를 깨우는
 * 시간과, HtmlService 가 화면을 iframe 안에 다시 띄우느라 오가는 왕복. 그건 코드로
 * 줄일 수 없으니, 그 사실을 아는 것 자체가 결론이다.
 */
function timingReport_() {
  const t = [];
  const step = (name, fn) => {
    const at = Date.now();
    const out = fn();
    t.push([name, Date.now() - at]);
    return out;
  };

  const total = Date.now();
  // 둘을 같이 재서 창이 실제로 무엇을 아껴주는지 눈으로 보게 한다.
  const all = step('loadRecords_ 전체 (옛 방식)', () => loadRecords_());
  const records = step('loadRecords_ 최근 ' + VIEW_WINDOW_DAYS + '일 (지금 방식)',
                       () => loadRecords_(viewWindowStart_()));
  step('readSettings_', () => readSettings_());
  step('readFixed_', () => readFixed_());
  step('readExtra_', () => readExtra_());
  step('detectFixed_ (계산만)', () => detectFixed_(records));
  const s = step('summarizeWeek_ (전체 다시)', () => summarizeWeek_());
  step('dashboardHtml_ (그리기)', () => dashboardHtml_(s));
  step('selfUrl_', () => selfUrl_());

  const lines = ['Timing (ms) — 캐시를 일부러 건너뛰고 전부 다시 계산한 값', ''];
  t.forEach(([name, ms]) => lines.push(`  ${String(ms).padStart(6)}  ${name}`));
  lines.push('', `  ${String(Date.now() - total).padStart(6)}  합계 (같은 일을 두 번 재므로 부풀려져 있다)`);
  lines.push('', `실제 화면 한 장 = summarizeWeek_ 한 번 ≈ 위의 '전체 다시' 값`);
  lines.push(`Ledger 행 수: 전체 ${all.length} → 화면이 읽는 것 ${records.length}`);

  // 캐시가 실제로 먹히는지는 이 화면으로 알 수 없다. 대시보드 맨 아래를 봐야 한다.
  const cache = CacheService.getScriptCache();
  lines.push('', '캐시 상태',
             `  view_html: ${cache.get('view_html') === null ? '비어 있음' : '들어 있음'}`,
             `  view_text: ${cache.get('view_text') === null ? '비어 있음' : '들어 있음'}`);
  lines.push('', '대시보드 맨 아래 줄에 "시각 · 소요ms · cached/fresh" 가 찍힌다.',
             'cached 인데도 느리면 남은 시간은 구글 몫이다 — 스크립트를 깨우는 시간과',
             'HtmlService 가 iframe 으로 다시 띄우는 왕복. 그건 코드로 줄일 수 없다.');
  return lines.join('\n');
}

/** 단축어의 Show Result 로 띄울 여러 줄 요약. */
function summaryText_(s) {
  if (!s.planAllow) return `Set "${BUDGET_KEY}" or "${CYCLE_BUDGET_KEY}" in ${SETTINGS}`;

  const lines = [
    s.remaining < 0 ? `Over budget this ${s.periodWord}` : `Left to spend this ${s.periodWord}`,
    won_(Math.abs(s.remaining)),
    '',
    `${fmtMD_(s.start)} – ${fmtMD_(s.end)} · ${s.daysLeft} ${s.daysLeft === 1 ? 'day' : 'days'} left`,
    // 넘겼을 때도 하루치를 남긴다. 넘긴 순간이 '앞으로 하루에 얼마까지'가 가장
    // 알고 싶은 때인데, 예전에는 문구만 뜨고 숫자가 사라졌다.
    s.remaining < 0
      ? `${won_(0)} a day · ${won_(-s.remaining)} over already`
      : `${won_(s.perDay)} a day from here`,
    // 화면의 페이스 눈금과 같은 값. 텍스트에는 눈금을 그릴 수 없으니 말로 적는다.
    ...(s.pace && s.pace.show
      ? [s.pace.vs >= 0
          ? `${won_(s.pace.vs)} under pace (${won_(s.pace.spend)} was fine by now)`
          : `${won_(-s.pace.vs)} over pace (${won_(s.pace.spend)} was fine by now)`]
      : []),
    '',
    `Spent    ${won_(s.variable)} / ${won_(s.weeklyAllow)}`,
    `Card bill ${won_(s.monthCharged)} this month`,
  ];

  const p = s.projection;
  if (p && p.show) {
    lines.push('', `On track for ${won_(p.out)} this month`,
               p.over > 0 ? `  ${won_(p.over)} over budget` : `  ${won_(-p.over)} under budget`);
    if (p.perDay > 0) {
      lines.push(`  ${won_(p.perDay)}/day for the last ${p.daysLeft} ${p.daysLeft === 1 ? 'day' : 'days'}`);
    }
  }

  if (s.overBudget) {
    lines.push('', `⚠ Over the monthly budget by ${won_(s.overBy)}`,
               `  ${won_(s.monthOut)} spent of ${won_(s.budget)}`);
  }
  if (s.income) {
    lines.push('', `Could save ${won_(s.savingsNet)} this month`,
               `  ${won_(s.monthIn)} in - ${won_(s.monthOut)} out` +
               (s.balanceHole ? ` - ${won_(s.balanceHole)} to clear the account` : ''));
  }

  if (s.categories.length) {
    lines.push('', 'Where the card went this month',
               `  ${won_(s.monthCut)} vs ${won_(s.prevCharged)} same days last month`);
    s.categories.slice(0, 12).forEach(c => {
      const d = c.delta ? `  (${c.delta > 0 ? '+' : ''}${won_(c.delta)})` : '';
      lines.push(`  ${c.name}  ${won_(c.amount)}  ${c.share}%${d}`);
    });
  }

  if (s.merchants.length) {
    lines.push('', 'Top merchants this month');
    s.merchants.forEach(m => lines.push(`  ${m.name}  ${m.count}x  ${won_(m.amount)}`));
  }

  if (s.reserved) {
    lines.push('', `Reserved ${won_(s.reserved)} still to come`);
    s.reserveRows.filter(r => r.remaining).forEach(r => {
      lines.push(`  ${r.name}  ${won_(r.remaining)}`);
    });
  }
  if (s.claimed.length) {
    lines.push('', `Counted as fixed this ${s.periodWord}`);
    s.claimed.forEach(c => lines.push(`  ${c.name}  <- ${c.merchant}  ${won_(c.amount)}`));
    lines.push('  (already inside the limit. Wrong merchant? Fill the Fixed sheet Match column)');
  }
  // 화면의 cutNote 와 같은 자리·같은 뜻. 경고가 아니라 한도의 출처를 밝히는 줄이다.
  // 예산과 한도가 다르면 그 차이가 어디서 왔는지는 어딘가에 남아 있어야 한다.
  if (s.shortfall) {
    lines.push('', `Limit    ${won_(s.planAllow)} by budget`,
               `         - ${won_(s.shortfall)} your account can't cover`,
               `         = ${won_(s.weeklyAllow)}`);
  }
  if (s.fixedTotal) lines.push('', `Fixed    ${won_(s.fixedTotal)} /month`);
  if (s.extraTotal) {
    lines.push(`Extra    ${won_(s.extraTotal)} this month`);
    s.extraRows.forEach(x => lines.push(`  ${x.name}  ${won_(x.amount)}`));
  }

  const f = s.forecast;
  if (f) {
    // 화면과 같은 이유로 날짜를 뺐다(forecastBlock 주석 참고). 대신 금액을 넣는다 —
    // 예전 경고 줄은 날짜만 있고 얼마나 파이는지가 없어서 정작 쓸모 있는 쪽이 빠져 있었다.
    lines.push('', f.negative
      ? `⚠ ${f.overdraft ? 'Past the overdraft limit' : 'Account runs out'} — ${won_(f.min)} at the lowest`
      : `Lowest balance ${won_(f.min)}`);
    if (f.schedule.length) {
      lines.push('', 'Still to leave your account this month',
                 `  Bank/cash ${won_(f.outBank)} · Card ${won_(f.outCard)}`);
      f.schedule.forEach(x => {
        lines.push(`  ${fmtMD_(x.date)}  ${x.name} (${x.via === 'in' ? 'income' : x.via})  ${won_(x.amount)}`);
      });
    }
  }

  if (s.rows.length) {
    lines.push('', `This ${s.periodWord}`);
    s.rows.slice().reverse().slice(0, 8).forEach(r => {
      // 날짜 + 영어 카테고리를 앞세우고, 저장해둔 영어 가맹점명을 뒤에 둔다
      const label = r.tag || CATEGORY_LABEL[r.category] || r.category;
      const md = r.date.slice(5).replace('-', '/');
      lines.push(`  ${md}  ${label}  ${won_(r.amount)}   ${r.merchantEn || r.merchant}`);
    });
  }
  lines.push('', '{{SERVED}}');    // 내보낼 때 '시각 · 소요ms · cached/fresh' 로 바뀐다
  return lines.join('\n');
}

/** 시트 메뉴에서 수동으로 확인할 때. */
function showThisWeek() {
  SpreadsheetApp.getUi().alert(summaryLine_(summarizeWeek_()));
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Budget')
    .addItem('This week', 'showThisWeek')
    .addItem('Update monthly', 'updateMonthly')
    .addItem('Clean up ledger', 'fixLedgerOrder')
    .addToUi();

  // 새 설정 항목·열은 시트를 열 때 붙인다. 예전에는 대시보드를 열 때마다 확인했는데
  // 그게 로딩을 몇 초씩 잡아먹었다. 스키마는 자주 바뀌지 않으니 여기가 맞는 자리다.
  // 권한이 모자라 실패해도 메뉴는 떠야 하므로 삼킨다.
  try { ensureSheets_(); } catch (err) { /* 메뉴만이라도 */ }
}

/**
 * 시트를 손으로 고치면 캐시를 버린다.
 *
 * 화면 캐시를 15분으로 길게 잡은 대신 이게 있어야 한다. 카테고리를 고치거나
 * 예산을 바꿨는데 15분간 옛 숫자가 보이면 고장난 것으로 보이기 때문이다.
 * 단순 트리거라 따로 설치할 필요 없이 저장만 하면 동작한다.
 */
function onEdit(e) {
  dropViewCache_();
}


// ════════════════════════════════════════════════════════════════ 대시보드

function dashboardHtml_(s) {
  const allow = s.weeklyAllow || 1;
  const pct = Math.min(100, Math.round(s.variable / allow * 100));
  // 예약분은 막대에 옅은 색으로 이어 붙여 '이미 임자 있는 돈'을 보여준다
  const pctReserved = Math.min(100 - pct, Math.round(s.reserved / allow * 100));
  const over = s.remaining < 0;
  const accent = over ? '#C0362C' : '#1B7F3B';

  // 막대 위 페이스 눈금. 표면색으로 낸 홈이라 채워진 쪽에서도 빈 쪽에서도 보인다.
  // 선을 얹으면 진한 막대 위에서는 어두운 선이, 옅은 트랙 위에서는 밝은 선이
  // 묻힌다 — 어느 색을 골라도 절반은 안 보인다.
  const pace = s.pace || {show: false};
  const paceMark = pace.show
    ? `<em class="pace" style="left:${Math.min(99, Math.max(0, pace.share))}%"></em>` : '';
  const paceWord = !pace.show ? ''
    : pace.vs >= 0 ? ` &nbsp;·&nbsp; ${won_(pace.vs)} under pace`
    : ` &nbsp;·&nbsp; <b class="up">${won_(-pace.vs)} over pace</b>`;

  const reserveList = s.reserveRows.filter(r => r.remaining).length
    ? `<h2>Still to come this ${s.periodWord}</h2><ul>${s.reserveRows.filter(r => r.remaining).map(r => `
        <li><span class="m">${escapeHtml_(r.name)}</span>
            <span class="a">${won_(r.remaining)}</span></li>`).join('')}</ul>`
    : '';

  // 이대로 가면 월말에 얼마가 되는지. 남은 돈보다 이게 먼저 눈에 들어와야 조절이 된다.
  const p = s.projection;
  const projBlock = p && p.show
    ? `<div class="proj${p.over > 0 ? ' bad' : ''}">
         <b>On track for ${won_(p.out)} this month</b>
         <span>${p.over > 0
           ? `${won_(p.over)} over budget`
           : `${won_(-p.over)} under budget`}${p.perDay > 0
           ? ` &nbsp;·&nbsp; ${won_(p.perDay)}/day for the last ${p.daysLeft} ${p.daysLeft === 1 ? 'day' : 'days'}`
           : ' &nbsp;·&nbsp; budget already spent'}</span>
       </div>`
    : '';

  const merchList = s.merchants.length
    ? `<h2>Top merchants this month</h2><ul>${s.merchants.map(m => `
        <li><span class="m">${escapeHtml_(m.name)}<em class="cnt">${m.count}x</em></span>
            <span class="a">${won_(m.amount)}</span></li>`).join('')}</ul>`
    : '';

  // 예산을 넘겼을 때만 뜨는 경고. '자유롭게 쓸 돈'이 아니라 이번 달 총지출 기준이다.
  // 위의 큰 숫자와 헷갈리지 않게 '총지출 기준'임을 문구에 박아둔다.
  // 큰 숫자는 고정비를 뺀 자유 한도 기준이라 둘은 서로 다른 값이다.
  const warnBlock = s.overBudget
    ? `<div class="warn">
         <b>Total spending passed the monthly budget</b>
         <span>${won_(s.monthOut)} / ${won_(s.budget)} &nbsp; +${won_(s.overBy)}</span>
         <span class="sub">card ${won_(s.monthCharged)} + off-card ${won_(s.monthOut - s.monthCharged)}</span>
       </div>`
    : '';

  // 월급을 적어둔 경우에만. 쓴 만큼 저금할 돈이 깎이는 게 눈에 보이도록.
  // 마이너스 잔액은 메우기 전까지 저금이 아니므로 여기서 빼고 보여준다.
  // 뺀 사실을 밑줄에 그대로 적는다 — 숫자만 낮아지면 어디서 깎였는지 알 길이 없다.
  const savingsBlock = s.income
    ? `<div class="card wide${s.savingsNet < 0 ? ' bad' : ''}">
         <b>Could save this month</b>
         <span>${won_(s.savingsNet)}</span>
         <em>${won_(s.monthIn)} in &minus; ${won_(s.monthOut)} out${
           s.balanceHole ? ` &minus; ${won_(s.balanceHole)} to clear the account` : ''}</em>
       </div>`
    : '';

  const f = s.forecast;
  // 마이너스 통장이 있으면 0 아래로 내려가는 것 자체는 정상이다. 한도를 넘을 때만 경고.
  // 한도를 안 적어둔 사람에게 '한도를 넘었다'고 하면 없는 한도를 말하는 셈이라 문구를 나눈다.
  const fcLabel = !f ? ''
    : f.negative ? (f.overdraft ? 'Past the overdraft limit' : 'Account runs out')
    : f.inOverdraft ? 'Lowest balance (using overdraft)'
    : 'Lowest balance';
  // 날짜는 안 적는다. 카드 결제일이 말일이면 급여일까지 통장에서 나가는 게 없어
  // 최저점이 거의 항상 '오늘'로 잡히고, 그러면 "08/24 에 바닥난다"가 적어둔 잔액을
  // 되풀이하는 말이 된다. 앞날의 예정일은 바로 아래 목록에 날짜와 함께 그대로 나온다.
  const forecastBlock = f
    ? `<div class="fc${f.negative ? ' bad' : ''}">
         <b>${fcLabel}</b>
         <span>${won_(f.min)}</span>
       </div>
       ${f.schedule.length ? `
       <h2>Still to leave your account this month</h2>
       <p class="split">Bank / cash <b>${won_(f.outBank)}</b> &nbsp;·&nbsp; Card <b>${won_(f.outCard)}</b></p>
       <ul class="sched">${f.schedule.map(x => `
         <li><span class="m">${fmtMD_(x.date)} &nbsp; ${escapeHtml_(x.name)}
             <em class="${x.via === 'Card' ? 'via card' : 'via'}">${escapeHtml_(
               x.via === 'in' ? 'income' : x.via)}</em></span>
             <span class="a">${won_(x.amount)}</span></li>`).join('')}</ul>` : ''}`
    : '';

  // 이번 달 카테고리별.
  //
  // 막대는 '돈'만 나타낸다 — 이번 달은 진한 막대, 지난달은 바로 아래 옅은 막대.
  // 둘이 같은 자로 그려지므로 길이를 그대로 견줄 수 있다. 예전에는 지난달을
  // 화살표 글자로만 적어서, 비교하려고 만든 구획인데 비교가 눈에 안 보였다.
  //
  // 퍼센트는 분모가 달라(전체 대비) 막대와 한 줄에 두면 서로 다른 걸 재게 된다.
  // 보조 줄로 내려 '지난달 금액 · 증감 · 이번 달의 몇 %' 를 함께 읽게 한다.
  const arrow = d => d > 0 ? `<em class="up">+${won_(d).replace('₩', '₩')}</em>`
    : d < 0 ? `<em class="down">${won_(d)}</em>` : '';
  const catBlock = s.categories.length
    ? `<h2>Where the card went this month</h2>
       <p class="split">${won_(s.monthCut)} vs ${won_(s.prevCharged)} over the same days last month
          ${s.prevCharged > 0 ? `<b class="${s.monthCut > s.prevCharged ? 'up' : 'down'}">${
            s.monthCut > s.prevCharged ? '+' : ''}${
            Math.round((s.monthCut - s.prevCharged) / s.prevCharged * 100)}%</b>` : ''}</p>
       <ul class="cats">${s.categories.slice(0, 12).map(c => `
         <li>
           <div class="row">
             <span class="m">${escapeHtml_(c.name)}</span>
             <span class="a">${won_(c.amount)}</span>
           </div>
           <i style="width:${c.bar}%"></i>
           ${c.prev > 0 ? `<u style="width:${c.prevBar}%"></u>` : ''}
           <span class="sub">${c.prev > 0
             ? `last month ${won_(c.prev)}${c.delta ? ' · ' + arrow(c.delta) : ''}`
             : 'new this month'} &nbsp;·&nbsp; ${c.share}% of the month</span>
         </li>`).join('')}</ul>`
    : '';

  const recent = s.rows.slice().reverse().slice(0, 8).map(r => `
    <li>
      <span class="m"><span class="dt">${escapeHtml_(r.date.slice(5).replace('-', '/'))}</span>
        ${escapeHtml_(r.tag || CATEGORY_LABEL[r.category] || r.category)}
        <span class="sub">${escapeHtml_(r.merchantEn || r.merchant)}</span></span>
      <span class="a">${won_(r.amount)}</span>
    </li>`).join('');

  // 한도가 어떻게 나온 숫자인지 밝혀둔다. 이게 없으면 큰 숫자의 출처를 알 길이 없다.
  const fixedNames = s.fixedRows.map(f => f.name).concat(Object.keys(s.detected));
  // planAllow 를 쓴다 — 이 줄이 설명하는 건 '예산이 말하는 한도'다. 깎인 뒤의 값을
  // 여기 넣으면 "예산 − 고정비 = 0" 이라는 산수가 안 맞는 문장이 된다.
  const fixedNote = s.fixedTotal
    ? `<p class="note">Limit = ${won_(s.budget)} budget &minus; ${won_(s.fixedTotal)} fixed costs${fixedNames.length ? ` (${escapeHtml_(fixedNames.slice(0, 3).join(', '))}${fixedNames.length > 3 ? ` +${fixedNames.length - 3}` : ''})` : ''} = <b>${won_(s.planAllow)}</b> free to spend this ${s.periodWord}.</p>`
    : '';

  // 잔액 때문에 깎인 몫. 화면 위쪽 빨간 배너는 없앴다 — 최저점 날짜가 대개 '오늘'이라
  // (말일 결제면 급여일까지 통장에서 나가는 게 없다) 날짜를 알려줘 봐야 적어둔 잔액을
  // 되풀이하는 것뿐이었기 때문이다.
  //
  // 그래도 이 줄은 남긴다. 예산 2,000,000 을 적어뒀는데 한도가 1,200,000 으로 나오면
  // 그 차이가 어디서 왔는지 화면 어딘가에는 반드시 있어야 한다. fixedNote 와 따로 두는
  // 것은 고정비가 없어도(fixedTotal 이 0이어도) 삭감은 일어나기 때문이다 — 붙여 뒀더니
  // Fixed 시트가 빈 사람에게는 설명이 통째로 사라졌다.
  // 문장은 홀로 서야 한다. 'of that' 처럼 앞줄을 가리키면 Fixed 시트가 빈 사람에게는
  // fixedNote 가 통째로 없어서 가리킬 대상이 사라진다.
  const cutNote = s.shortfall
    ? `<p class="note">The limit is <b>${won_(s.weeklyAllow)}</b>, not ${won_(s.planAllow)}:
        over the next ${FORECAST_DAYS} days the balance forecast comes up ${won_(s.shortfall)}
        short of ${s.forecast && s.forecast.overdraft ? 'your overdraft limit' : 'zero'},
        so that much is taken off.</p>`
    : '';

  const extraNote = s.extraTotal
    ? `<p class="note">${won_(s.extraTotal)} extra this month (${escapeHtml_(s.extraRows.map(x => x.name).join(', '))}) added to the limit above.</p>`
    : '';

  // 어떤 결제가 '이미 예산에 든 돈'으로 처리됐는지. 이름 맞추기가 헐거워 엉뚱한 걸
  // 집을 수 있으므로, 가맹점명을 그대로 보여줘서 눈으로 확인할 수 있게 한다.
  const claimedBlock = s.claimed.length
    ? `<h2>Counted as fixed this ${s.periodWord}</h2>
       <ul>${s.claimed.map(c => `
         <li><span class="m">${escapeHtml_(c.name)}<em class="via">${escapeHtml_(c.merchant)}</em></span>
             <span class="a">${won_(c.amount)}</span></li>`).join('')}</ul>
       <p class="note">Already inside the limit above, so not counted twice.
          If a merchant here looks wrong, put the exact text from the card message
          into the <b>Match</b> column of the Fixed sheet.</p>`
    : '';

  // 자동 새로고침(meta refresh)은 쓰지 않는다. 그건 임시 주소를 다시 부르는 셈이라
  // 90초 뒤 반드시 빈 화면이 됐다. 대신 언제 기준인지 적고, 직접 누르는 링크를 둔다.
  // Refresh 는 &fresh=1 로 캐시를 건너뛴다. 안 그러면 눌러도 캐시된 같은 화면이 온다.
  // {{SERVED}} 는 내보낼 때 '시각 · 소요ms · cached/fresh' 로 바뀐다 (cachedView_).
  const url = selfUrl_();
  const footer = `<p class="foot">{{SERVED}}${
    url ? ` · <a href="${escapeHtml_(url + '&fresh=1')}" target="_top">Refresh</a>` : ''}</p>`;

  return `
<style>
  :root { color-scheme: light dark; }
  * { box-sizing: border-box; }
  body { margin:0; padding:22px 18px 40px; font:16px/1.5 -apple-system,BlinkMacSystemFont,"Apple SD Gothic Neo",sans-serif;
         background:#fff; color:#1a1a1a; -webkit-text-size-adjust:100%; }
  h1 { font-size:13px; font-weight:600; letter-spacing:.02em; color:#6b7280; margin:0 0 4px; }
  .period { font-size:13px; color:#9ca3af; margin-bottom:22px; }
  .big { font-size:44px; font-weight:700; letter-spacing:-.02em; color:${accent}; margin:0; }
  .cap { font-size:14px; color:#6b7280; margin:2px 0 20px; }
  .bar { height:10px; border-radius:5px; background:#e9ecef; overflow:hidden; margin-bottom:8px;
         display:flex; position:relative; }
  .bar > i { display:block; height:100%; width:${pct}%; background:${accent}; }
  .bar > u { display:block; height:100%; width:${pctReserved}%; background:${accent}; opacity:.35; }
  /* 페이스 눈금 — 색을 얹지 않고 표면색으로 홈을 낸다. 채워진 쪽에서도 빈 쪽에서도 보인다. */
  .bar > .pace { position:absolute; top:0; bottom:0; width:3px; margin-left:-1px;
                 background:#fff; }
  .scale { display:flex; justify-content:space-between; font-size:12px; color:#9ca3af; margin-bottom:26px; }
  .grid { display:grid; grid-template-columns:1fr 1fr; gap:12px; margin-bottom:24px; }
  .card { background:#f6f7f9; border-radius:12px; padding:14px; }
  .card b { display:block; font-size:12px; font-weight:500; color:#6b7280; margin-bottom:4px; }
  .card span { font-size:19px; font-weight:600; }
  .card.wide { grid-column:1 / -1; }
  .card em { display:block; font-style:normal; font-size:12px; color:#9ca3af; margin-top:4px;
             font-variant-numeric:tabular-nums; }
  .card.bad span { color:#C0362C; }
  /* 세로로 쌓는다. 예전에는 좌우 flex 였는데 오른쪽 숫자에 white-space:nowrap 이
     걸려 있어서, '₩2,081,180 / ₩2,000,000 +₩81,180' 처럼 긴 문자열이 폭을 다
     차지하고 왼쪽 라벨을 한 글자짜리 세로 기둥으로 짓눌렀다. 폰 폭에서는 라벨과
     금액이 한 줄에 같이 들어갈 수가 없다. 바로 아래 .proj 도 세로로 쌓아서 멀쩡하다. */
  .warn { background:#fdecea; border-radius:12px; padding:13px 14px; margin-bottom:20px; }
  .warn b { display:block; font-size:13px; color:#C0362C; }
  .warn span { display:block; font-size:13px; color:#C0362C;
               font-variant-numeric:tabular-nums; margin-top:3px; }
  .warn .sub { font-weight:400; font-size:12px; opacity:.75; }
  .note { font-size:12px; color:#9ca3af; margin:0 0 22px; }
  h2 { font-size:13px; color:#6b7280; margin:0 0 8px; }
  ul { list-style:none; margin:0; padding:0; }
  li { display:flex; justify-content:space-between; gap:12px; padding:9px 0; border-bottom:1px solid #eceef1; font-size:14px; }
  li:last-child { border-bottom:0; }
  .m { overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .m .dt { color:#9ca3af; font-size:12px; font-variant-numeric:tabular-nums; margin-right:6px; }
  .m .sub { color:#9ca3af; font-size:12px; margin-left:6px; }
  .m em { font-style:normal; font-size:11px; color:#8b5cf6; background:#f3e8ff; border-radius:4px; padding:1px 5px; margin-left:4px; }
  .a { font-variant-numeric:tabular-nums; white-space:nowrap; }
  .empty { color:#9ca3af; font-size:14px; }
  .foot { margin:26px 0 0; font-size:12px; color:#9ca3af; }
  .foot a { color:#1B7F3B; text-decoration:none; font-weight:600; }
  .fc { display:flex; justify-content:space-between; align-items:baseline; gap:10px;
        background:#f6f7f9; border-radius:12px; padding:13px 14px; margin-bottom:8px; }
  .fc b { font-size:12px; font-weight:500; color:#6b7280; }
  .fc span { font-size:16px; font-weight:600; font-variant-numeric:tabular-nums; }
  .fc.bad { background:#fdecea; }
  .fc.bad b, .fc.bad span { color:#C0362C; }
  @media (prefers-color-scheme: dark) { .fc.bad, .warn { background:#3a1d1a; } }
  .sched { margin-bottom:24px; }
  .sched li { font-size:13px; color:#6b7280; padding:7px 0; }
  .split { font-size:13px; color:#6b7280; margin:0 0 10px; }
  .split b { font-variant-numeric:tabular-nums; color:#1a1a1a; }
  .via { font-style:normal; font-size:11px; color:#6b7280; background:#eceef1;
         border-radius:4px; padding:1px 5px; margin-left:6px; }
  .via.card { color:#8a5a00; background:#fdf1d8; }
  .cats { margin-bottom:24px; }
  .cats li { display:block; padding:10px 0; border-bottom:1px solid #eceef1; }
  .cats .row { display:flex; justify-content:space-between; gap:12px;
               font-size:14px; margin-bottom:5px; }
  /* 막대는 데이터색(파랑)이다. 초록·빨강은 '예산 안/밖'을 뜻하는 상태색이라
     여기 쓰면 카테고리 막대가 '좋다/나쁘다'로 읽힌다. */
  .cats i { display:block; height:6px; border-radius:3px; background:#3A6EA5; }
  .cats u { display:block; height:4px; border-radius:2px; background:#3A6EA5;
            opacity:.28; margin-top:2px; }
  .cats .sub { display:block; font-size:11px; color:#9ca3af; margin-top:5px;
               font-variant-numeric:tabular-nums; }
  .cats .a em, .cats .m em { font-style:normal; font-size:12px; color:#9ca3af; margin-left:5px; }
  .up { color:#C0362C !important; }
  .down { color:#1B7F3B !important; }
  .split b.up, .split b.down { font-weight:700; margin-left:4px; }
  .proj { background:#eef6f0; border-radius:12px; padding:13px 14px; margin-bottom:20px; }
  .proj b { display:block; font-size:15px; font-weight:700; color:#1B7F3B;
            font-variant-numeric:tabular-nums; }
  .proj span { font-size:12px; color:#6b7280; font-variant-numeric:tabular-nums; }
  .proj.bad { background:#fdecea; }
  .proj.bad b { color:#C0362C; }
  .cnt { font-style:normal; font-size:11px; color:#6b7280; background:#eceef1;
         border-radius:4px; padding:1px 5px; margin-left:6px; }
  /* 덩어리 구분. 화면 하나에 블록이 열둘이라 어디가 어디인지 안내가 필요하다. */
  .grp { font-size:11px; font-weight:700; letter-spacing:.08em; text-transform:uppercase;
         color:#9ca3af; margin:30px 0 12px; padding-bottom:6px;
         border-bottom:1px solid #eceef1; }
  /* 한도가 어떻게 나왔는지는 늘 펼쳐둘 필요가 없다. 궁금할 때만 연다. */
  .more { margin-top:26px; }
  .more summary { font-size:12px; color:#6b7280; cursor:pointer; padding:8px 0; }
  .more[open] summary { margin-bottom:10px; }
  .more h2 { margin-top:14px; }
  @media (prefers-color-scheme: dark) {
    body { background:#131417; color:#e8e8ea; }
    .card { background:#1e2025; }
    .bar { background:#2a2d34; }
    li { border-bottom-color:#232529; }
    .m .sub { color:#6b7480; }
    .m em { background:#2e1f45; color:#c4a5fb; }
    .split b { color:#e8e8ea; }
    .via, .cnt { background:#2a2d34; color:#9aa0a6; }
    .via.card { background:#3d3116; color:#e8c980; }
    .proj { background:#16241a; }
    .proj.bad { background:#3a1d1a; }
    /* 페이스 홈은 표면색이라 모드마다 따라가야 한다. 밝은 색으로 두면 어두운
       화면에서 홈이 아니라 밝은 선으로 보인다. */
    .bar > .pace { background:#131417; }
    .grp { color:#6b7480; border-bottom-color:#232529; }
    .more summary { color:#9aa0a6; }
  }
</style>

<h1>${over ? `Over budget this ${s.periodWord}` : `Left to spend this ${s.periodWord}`}</h1>
<div class="period">${fmtMD_(s.start)} – ${fmtMD_(s.end)} · ${s.daysLeft} ${s.daysLeft === 1 ? 'day' : 'days'} left</div>

<p class="big">${won_(Math.abs(s.remaining))}</p>
<p class="cap">${over
  ? `${won_(0)} a day &nbsp;·&nbsp; ${won_(-s.remaining)} over already`
  : `${won_(s.perDay)} a day from here`}${paceWord}</p>

<div class="bar"><i></i><u></u>${paceMark}</div>
<div class="scale"><span>${won_(s.variable)} spent${s.reserved ? ` · ${won_(s.reserved)} reserved` : ''}</span><span>${won_(s.weeklyAllow)} limit</span></div>

${projBlock}
${warnBlock}

<h3 class="grp">This month</h3>
<div class="grid">
  <div class="card"><b>Card bill this month</b><span>${won_(s.monthCharged)}</span></div>
  <div class="card"><b>Monthly budget</b><span>${won_(s.budget)}</span></div>
  ${savingsBlock}
</div>
${catBlock}
${merchList}

${forecastBlock || reserveList ? `<h3 class="grp">Ahead</h3>${forecastBlock}${reserveList}` : ''}

<h3 class="grp">This ${s.periodWord}</h3>
${recent ? `<ul>${recent}</ul>` : '<p class="empty">Nothing yet.</p>'}

${claimedBlock || fixedNote || cutNote || extraNote ? `
<details class="more">
  <summary>How the limit was worked out</summary>
  ${claimedBlock}${fixedNote}${cutNote}${extraNote}
</details>` : ''}
${footer}
`;
}

function escapeHtml_(s) {
  return String(s).replace(/[&<>"']/g, c =>
    ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[c]));
}


// ══════════════════════════════════════════════════════════════════ 테스트

/** 스크립트 편집기 드롭다운에서 이걸 골라 실행해 동작을 확인한다. */
function runTest() {
  const md = Utilities.formatDate(new Date(), TZ, 'MM/dd');

  // 취소 문자는 '매출취소' 에 시각도 없다 — 승인과 형식이 달라 따로 확인한다
  const cancel = '[Web발신]\n신한카드 (1234)  매출취소  홍*동       29,902원  ' + md + ' 네이버페이';
  const c = parseMessage_(cancel);
  Logger.log(c && c.amount === -29902
    ? 'Cancel OK: ' + c.merchant + ' ' + c.amount
    : 'Cancel FAILED: ' + JSON.stringify(c));

  // 현대카드는 카드번호 괄호가 없어 다른 정규식으로 넘어간다 — 별도로 확인한다
  const hyundai = '[Web발신]\n현대카드Z ontact 승인\n홍*동\n10,000원 일시불\n' +
                  md + ' 17:10\n모바일티머니선\n누적206,250원';
  const h = parseMessage_(hyundai);
  Logger.log(h && h.amount === 10000 && h.merchant === '모바일티머니선'
    ? 'Hyundai OK: ' + h.merchant + ' ' + h.amount
    : 'Hyundai FAILED: ' + JSON.stringify(h));

  // 외화 해외결제 — 금액이 원이 아니고 날짜 뒤에 온다. 파싱만 확인한다(저장하지 않는다)
  const abroad = parseMessage_('[Web발신]\n현대 대한항공120 해외승인\n홍*동님\n' +
                               md + ' 16:40\nJPY 2,000.00\nUBER*TRIPHELP.UBER.COM');
  Logger.log(abroad && abroad.currency === 'JPY' && abroad.foreignAmount === 2000 && !abroad.cancel
    ? 'Foreign OK: ' + abroad.merchant + ' ' + abroad.currency + ' ' + abroad.foreignAmount +
      ' -> ' + recLabel_((settleForeign_(abroad), abroad))
    : 'Foreign FAILED: ' + JSON.stringify(abroad));
  const abroadCancel = parseMessage_('[Web발신]\n현대 대한항공120 해외승인 취소\n홍*동님\n' +
                                     md + ' 16:41\nJPY 1,258.00\nUBER*PENDING');
  Logger.log(abroadCancel && abroadCancel.cancel && abroadCancel.foreignAmount === 1258
    ? 'Foreign cancel OK: ' + abroadCancel.merchant + ' ' + abroadCancel.foreignAmount
    : 'Foreign cancel FAILED: ' + JSON.stringify(abroadCancel));

  const sample = '[Web 발신]\n신한카드(1234)승인 홍*동 6,500원(일시불)' + md + ' 08:30 스타벅스 강남';
  const rec = parseMessage_(sample);
  Logger.log(rec ? 'Parsed OK: ' + rec.merchant + ' ' + rec.amount : 'Parse failed');
  if (rec) {
    Logger.log('Saved: ' + appendRecord_(rec));
    Logger.log(summaryLine_(summarizeWeek_()));
  }
}
