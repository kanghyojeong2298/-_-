"""
소포수령증 자동화 웹앱 (Streamlit)
실행: streamlit run app.py
"""

import streamlit as st
import tempfile
import os
import sys
import re

# ── 허용된 직원 이메일 목록 (배포 후 여기에 직원 이메일 추가) ──────────────
ALLOWED_EMAILS = [
    "m0311@taxexpert.kr",
    "m1007@taxexpert.kr",
    "m0607@taxexpert.kr",
    "m1225@taxexpert.kr",
    "m1024@taxexpert.kr",
    "m1211@taxexpert.kr",
    "m0127@taxexpert.kr",
    "m0224@taxexpert.kr",
    "m1018@taxexpert.kr",  # 관리자
]
import calendar
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent))

from modules.pdf_parser   import parse_pdf, detect_pdf_type
from modules.excel_writer import generate_excel

CURRENCIES = ["MYR", "PHP", "SGD", "THB", "TWD", "VND", "JPY"]

CURRENCY_NAMES = {
    "MYR": "말레이시아 링깃 (MYR)", "PHP": "필리핀 페소 (PHP)",
    "SGD": "싱가포르 달러 (SGD)", "THB": "태국 바트 (THB)",
    "TWD": "대만 달러 (TWD)",     "VND": "베트남 동 (VND)",
    "JPY": "일본 엔 (JPY) (100)",
}

# SMBS 통화명 → 코드 매핑
SMBS_NAME_TO_CODE = {
    "MYR": "MYR", "말레이시아": "MYR",
    "PHP": "PHP", "필리핀":    "PHP",
    "SGD": "SGD", "싱가포르":   "SGD",
    "THB": "THB", "태국":      "THB",
    "TWD": "TWD", "대만":      "TWD",
    "VND": "VND", "베트남":     "VND",
    "JPY": "JPY", "일본":      "JPY",
}

DEFAULT_RATES = {
    "MYR": 0.0, "PHP": 0.0, "SGD": 0.0, "THB": 0.0,
    "TWD": 0.0, "VND": 0.0, "JPY": 0.0,
}

# ── 고정환율(내장) 자동 적용 ─────────────────────────────────────────
# data/fixed_rates_2025.json 에 통화별 일별 매매기준율이 들어 있습니다.
# 직원이 환율을 입력할 필요 없이, 소포수령증 발행일에 맞는 환율이 자동 적용됩니다.
# 환율을 갱신하려면 이 JSON 파일만 교체(수정)하면 전체에 반영됩니다.
AUTO_RATE_LABEL = "🔒 자동 (내장 고정환율)"

# 환율 JSON을 여러 위치에서 자동 탐색 (data/ 폴더 안이든, 루트든, 어디 두든 찾음)
_HERE = os.path.dirname(os.path.abspath(__file__))
_FIXED_RATES_CANDIDATES = [
    os.path.join(_HERE, 'data', 'fixed_rates_2025.json'),
    os.path.join(_HERE, 'fixed_rates_2025.json'),
    os.path.join(os.getcwd(), 'data', 'fixed_rates_2025.json'),
    os.path.join(os.getcwd(), 'fixed_rates_2025.json'),
]


def load_fixed_rates() -> dict:
    """내장 JSON에서 통화별 일별 환율을 읽어, 앱 내부 환율 형식(daily 포함)으로 반환."""
    import json
    raw = None
    for _p in _FIXED_RATES_CANDIDATES:
        try:
            with open(_p, encoding='utf-8') as f:
                raw = json.load(f)
            break
        except Exception:
            continue
    if raw is None:
        return {}
    rates_raw = raw.get('rates', {})
    result = {}
    for cur, daymap in rates_raw.items():
        daily = [{'date': d, 'rate': float(r), 'change': 0.0, 'cross': 0.0}
                 for d, r in sorted(daymap.items())]
        if not daily:
            continue
        rs = [d['rate'] for d in daily if d['rate'] > 0]
        result[cur] = {
            'period':        f"{daily[0]['date']} ~ {daily[-1]['date']}",
            'currency':      cur,
            'currency_name': CURRENCY_NAMES.get(cur, cur),
            'average':       round(sum(rs) / len(rs), 2) if rs else 0.0,
            'min':           min(rs) if rs else 0.0,
            'max':           max(rs) if rs else 0.0,
            'min_date': '', 'max_date': '', 'range': 0.0, 'cross_rate': 0.0,
            'daily':         daily,
        }
    return result

# ── 페이지 설정 ─────────────────────────────────────────────────
st.set_page_config(page_title="소포수령증 자동화", page_icon="📦", layout="centered")

# ── Google OAuth 인증 (google_credentials.json 파일이 있을 때만 활성화) ──────
_CREDENTIALS_FILE = os.path.join(os.path.dirname(__file__), 'google_credentials.json')
_AUTH_ENABLED = (
    os.path.exists(_CREDENTIALS_FILE)
    and '"client_id"' in open(_CREDENTIALS_FILE, encoding='utf-8').read()
    and '여기에' not in open(_CREDENTIALS_FILE, encoding='utf-8').read()
)

if _AUTH_ENABLED:
    try:
        from streamlit_google_auth import Authenticate as _GoogleAuth
        _redirect_uri = os.environ.get(
            'REDIRECT_URI',
            'http://localhost:8501'
        )
        _cookie_secret = os.environ.get('COOKIE_SECRET', 'sopo-cookie-secret-2024')
        _authenticator = _GoogleAuth(
            secret_credentials_path=_CREDENTIALS_FILE,
            cookie_name='soposuryjungjeung_auth',
            cookie_key=_cookie_secret,
            redirect_uri=_redirect_uri,
        )
        _authenticator.check_authentification()

        if not st.session_state.get('connected'):
            st.markdown("""
            <div style="text-align:center; padding:3rem 1rem;">
                <p style="font-size:2rem; font-weight:700; color:#1f4e79;">📦 소포수령증 자동화</p>
                <p style="color:#555; margin-bottom:2rem;">사용하려면 Google 계정으로 로그인하세요.</p>
            </div>
            """, unsafe_allow_html=True)
            _authenticator.login()
            st.stop()

        _user_email = st.session_state.get('email', '')
        if ALLOWED_EMAILS and _user_email not in ALLOWED_EMAILS:
            st.error(f"❌ 접근 권한이 없습니다. ({_user_email})\n\n관리자에게 문의하세요.")
            if st.button("로그아웃"):
                _authenticator.logout()
            st.stop()

        # 우상단 사용자 정보 표시
        with st.sidebar:
            _user_name = st.session_state.get('name', _user_email)
            st.markdown(f"**👤 {_user_name}**")
            st.markdown(f"<small>{_user_email}</small>", unsafe_allow_html=True)
            if st.button("로그아웃"):
                _authenticator.logout()
                st.rerun()

    except ImportError:
        pass  # 로컬 환경: streamlit-google-auth 미설치 → 인증 없이 실행

st.markdown("""
<style>
    .main-title { font-size:2rem; font-weight:700; color:#1f4e79; margin-bottom:0.2rem; }
    .sub-title  { font-size:1rem; color:#555; margin-bottom:2rem; }
    .warn-box   { background:#fff8e1; border-radius:8px; padding:0.8rem 1.2rem;
                  border-left:4px solid #f9a825; margin-bottom:0.5rem; }
    .info-box   { background:#e3f2fd; border-radius:8px; padding:0.8rem 1.2rem;
                  border-left:4px solid #1565c0; margin-bottom:0.5rem; }
</style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-title">📦 소포수령증 자동화</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">소포수령증 PDF를 업로드하면 매출집계 엑셀을 자동으로 만들어 드립니다.</p>', unsafe_allow_html=True)
st.divider()


# ══════════════════════════════════════════════════════════════════
# STEP 1 — PDF 업로드
# ══════════════════════════════════════════════════════════════════
st.markdown("### 📄 STEP 1 — 소포수령증 PDF 업로드")
st.caption("쇼피(MY/PH/SG/TH/TW/VN), 라자다 파일을 한꺼번에 올려주세요. *큐텐재팬은 STEP 2에서 진행해주세요.")

uploaded_files = st.file_uploader(
    "PDF 파일 선택 (여러 개 동시 선택 가능)",
    type=["pdf"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

if uploaded_files:
    st.markdown("**업로드된 파일:**")
    cols = st.columns(2)
    for i, f in enumerate(uploaded_files):
        pdf_type = detect_pdf_type(f.name)
        icon  = {"shopee":"🛍️","lazada":"🟠","qoo10":"🇯🇵","unknown":"❓"}.get(pdf_type,"📄")
        label = {"shopee":"쇼피","lazada":"라자다","qoo10":"큐텐재팬","unknown":"미확인"}.get(pdf_type,"")
        cols[i%2].markdown(f"{icon} `{f.name}` — {label}")

st.divider()


# ══════════════════════════════════════════════════════════════════
# STEP 2 — 큐텐재팬 정보 입력
# ══════════════════════════════════════════════════════════════════
st.markdown("### 🇯🇵 STEP 2 — 큐텐재팬 정보 입력")
st.markdown(
    '<div class="warn-box">'
    '큐텐재팬 PDF는 이미지 형식이라 <b>자동 추출이 되지 않습니다.</b><br>'
    'PDF를 열어서 아래에 직접 입력해 주세요.'
    '</div>',
    unsafe_allow_html=True,
)
st.write("")

col1, col2, col3 = st.columns(3)
with col1:
    qoo10_amount = st.number_input("JPY 합계 금액", min_value=0, value=0, format="%d",
                                    help="예: 3802685")
with col2:
    qoo10_qty = st.number_input("발송 건수", min_value=0, value=0, help="예: 386")
with col3:
    qoo10_tracking = st.text_input("발송번호", placeholder="예: K2512244647017 외")

col4, col5, col6 = st.columns(3)
with col4:
    qoo10_period_start = st.text_input("거래기간 시작일", placeholder="예: 2025-12-01")
with col5:
    qoo10_period_end = st.text_input("거래기간 종료일", placeholder="예: 2025-12-31")
with col6:
    qoo10_write_date_input = st.text_input("발행일 (작성일자)", placeholder="예: 2026-01-05",
                                            help="소포수령증에 기재된 작성일자. 환율 기준일로 사용됩니다.")

st.divider()


# ══════════════════════════════════════════════════════════════════
# STEP 3 — 환율 입력
# ══════════════════════════════════════════════════════════════════
st.markdown("### 💱 STEP 3 — 환율")
st.markdown(
    "📌 **소포수령증 발행일(작성일자)** 기준 환율을 적용합니다 (주말·공휴일이면 직전 영업일)  \n"
    "👉 환율 조회: [서울외국환중개(SMBS)](http://www.smbs.biz/ExRate/StdExRate.jsp)"
)
st.write("")

rate_mode = st.radio(
    "환율 입력 방식",
    [AUTO_RATE_LABEL, "📊 SMBS 엑셀 업로드", "✏️ 직접 입력"],
    horizontal=True,
    label_visibility="collapsed",
)

manual_rates = {cur: 0.0 for cur in CURRENCIES}
smbs_excel_files = []

if rate_mode == AUTO_RATE_LABEL:
    _fr = load_fixed_rates()
    if _fr:
        _curs   = ", ".join(_fr.keys())
        _missing = [c for c in CURRENCIES if c not in _fr]
        _period  = next(iter(_fr.values())).get('period', '')
        st.markdown(
            '<div class="info-box">'
            '✅ <b>환율 자동 적용</b> — 환율을 수동으로 입력하지 않아도 됩니다.<br>'
            '소포수령증 <b>발행일에 맞는 날짜의 환율</b>이 자동으로 적용됩니다.<br>'
            f'<small>내장 환율 기간: {_period}<br>적용 통화: {_curs}</small>'
            '</div>',
            unsafe_allow_html=True,
        )
        if _missing:
            st.caption(f"⚠️ 내장 환율 없는 통화: {', '.join(_missing)} → 0 처리 (필요 시 다른 모드 사용)")
    else:
        st.warning(
            "⚠️ 내장 환율 파일(`data/fixed_rates_2025.json`)을 찾을 수 없습니다.  \n"
            "다른 입력 방식을 선택하거나, 데이터 파일을 추가해 주세요."
        )

elif rate_mode == "📊 SMBS 엑셀 업로드":
    st.markdown(
        '<div class="info-box">'
        '<b>SMBS 엑셀 다운로드 방법:</b><br>'
        '1. <a href="http://www.smbs.biz/ExRate/StdExRate.jsp" target="_blank">SMBS 사이트</a> 접속<br>'
        '2. 통화 선택 → 기간 설정 → 조회 → <b>엑셀 저장</b> 클릭<br><br>'
        '<b>📁 업로드 방식 (둘 다 지원)</b><br>'
        '• <b>방식 A</b> — 통화별 엑셀 파일을 따로 저장 후 <b>여러 파일 동시 업로드</b><br>'
        '• <b>방식 B</b> — 엑셀 하나에 <b>여러 시트(탭)</b>로 통화별 데이터 정리 후 업로드<br>'
        '&nbsp;&nbsp;&nbsp;(시트 이름에 통화코드 포함 권장: 예 <i>MYR, JPY, SGD</i>…)'
        '</div>',
        unsafe_allow_html=True,
    )
    smbs_excel_files = st.file_uploader(
        "SMBS 엑셀 파일 업로드",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="smbs_excel",
        label_visibility="collapsed",
    )
    if smbs_excel_files:
        st.caption(f"✅ {len(smbs_excel_files)}개 파일 업로드됨")

else:
    st.caption("거래기간 마지막날 환율을 직접 입력하세요 (0이면 해당 통화 미사용)")
    c1, c2, c3, c4 = st.columns(4)
    manual_rates['MYR'] = c1.number_input("🇲🇾 MYR",        value=0.0, format="%.2f")
    manual_rates['PHP'] = c2.number_input("🇵🇭 PHP",        value=0.0, format="%.2f")
    manual_rates['SGD'] = c3.number_input("🇸🇬 SGD",        value=0.0, format="%.2f")
    manual_rates['THB'] = c4.number_input("🇹🇭 THB",        value=0.0, format="%.2f")
    c5, c6, c7, _ = st.columns(4)
    manual_rates['TWD'] = c5.number_input("🇹🇼 TWD",        value=0.0, format="%.2f")
    manual_rates['VND'] = c6.number_input("🇻🇳 VND",        value=0.0, format="%.2f")
    manual_rates['JPY'] = c7.number_input("🇯🇵 JPY (100엔)", value=0.0, format="%.2f")

st.divider()


# ══════════════════════════════════════════════════════════════════
# STEP 4 — 처리 시작
# ══════════════════════════════════════════════════════════════════
st.markdown("### ⚡ STEP 4 — 처리 시작")

process_btn = st.button(
    "🚀 엑셀 파일 생성하기",
    type="primary",
    use_container_width=True,
    disabled=not bool(uploaded_files),
)
if not uploaded_files:
    st.caption("⬆️ PDF 파일을 먼저 업로드해 주세요")


# ── SMBS 엑셀 파싱 함수 ──────────────────────────────────────────
def _parse_smbs_sheet(ws, fallback_currency=None) -> tuple:
    """
    단일 시트에서 날짜/환율 데이터 추출.
    반환: (currency_code, daily_list) 또는 (None, [])
    """
    daily = []
    currency = None

    for row in ws.iter_rows(values_only=True):
        if not row or row[0] is None:
            continue

        cell0 = str(row[0]).strip()

        # 날짜 행 감지: YYYY.MM.DD 또는 YYYY-MM-DD
        date_match = re.match(r'(\d{4})[.\-](\d{2})[.\-](\d{2})', cell0)
        if date_match:
            date_str = f"{date_match.group(1)}.{date_match.group(2)}.{date_match.group(3)}"
            try:
                rate   = float(str(row[2]).replace(',', '')) if row[2] is not None else 0.0
                change = float(str(row[3]).replace(',', '')) if row[3] is not None else 0.0
                cross  = float(str(row[4]).replace(',', '')) if row[4] is not None else 0.0
            except (ValueError, TypeError):
                continue

            # 통화 코드 추출 (두 번째 열: 통화명)
            if currency is None and row[1]:
                cur_name = str(row[1])
                for key, code in SMBS_NAME_TO_CODE.items():
                    if key in cur_name:
                        currency = code
                        break

            daily.append({'date': date_str, 'rate': rate,
                          'change': change, 'cross': cross})
        else:
            # 헤더/통화명 행에서 통화 코드 추출 시도
            if currency is None:
                for val in row:
                    if val is None:
                        continue
                    val_str = str(val)
                    for key, code in SMBS_NAME_TO_CODE.items():
                        if key in val_str:
                            currency = code
                            break
                    if currency:
                        break

    # 시트 이름에서 통화 코드 추출 시도
    if currency is None and fallback_currency:
        fname_upper = fallback_currency.upper()
        for code in CURRENCIES:
            if code in fname_upper:
                currency = code
                break

    return currency, daily


def _build_currency_entry(currency, daily):
    """daily 리스트로 통화 환율 딕셔너리 생성"""
    rates_only = [d['rate'] for d in daily if d['rate'] > 0]
    avg = round(sum(rates_only) / len(rates_only), 2) if rates_only else 0.0
    return {
        'period':        f"{daily[0]['date']} ~ {daily[-1]['date']}",
        'currency':      currency,
        'currency_name': CURRENCY_NAMES.get(currency, currency),
        'average':       avg,
        'min':           min(rates_only) if rates_only else 0.0,
        'max':           max(rates_only) if rates_only else 0.0,
        'min_date':      '', 'max_date': '',
        'range':         0.0, 'cross_rate': 0.0,
        'daily':         daily,
    }


def parse_smbs_excel_files(excel_files) -> dict:
    """
    SMBS에서 다운로드한 엑셀 파일들을 파싱하여 통화별 일별 환율 딕셔너리 반환.

    지원 형식:
    1. 통화별 별도 파일 (각 파일에 시트 1개)
    2. 단일 파일에 여러 시트 (시트마다 통화 1개)

    반환: {'MYR': {'period':..., 'currency':'MYR', 'average':..., 'daily':[...]}, ...}
    """
    import openpyxl
    from io import BytesIO

    result = {}

    for ef in excel_files:
        try:
            file_bytes = ef.read()
            wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)

            if len(wb.sheetnames) > 1:
                # ── 다중 시트: 시트마다 통화 1개 ──
                for sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    currency, daily = _parse_smbs_sheet(ws, fallback_currency=sheet_name)

                    # 시트명에서 통화 코드 직접 추출 (예: "MYR", "JPY 환율" 등)
                    if currency is None:
                        sn_upper = sheet_name.upper()
                        for code in CURRENCIES:
                            if code in sn_upper:
                                currency = code
                                break

                    if daily and currency:
                        result[currency] = _build_currency_entry(currency, daily)
                    elif daily and not currency:
                        st.caption(f"  ⚠️ `{ef.name}` → 시트 '{sheet_name}' 통화 코드 인식 불가 (건너뜀)")
            else:
                # ── 단일 시트: 파일 1개 = 통화 1개 ──
                ws = wb.active
                currency, daily = _parse_smbs_sheet(ws, fallback_currency=ef.name)

                # 파일명에서 통화 코드 추출 시도
                if currency is None:
                    fname_upper = ef.name.upper()
                    for code in CURRENCIES:
                        if code in fname_upper:
                            currency = code
                            break

                if daily and currency:
                    result[currency] = _build_currency_entry(currency, daily)
                elif daily and not currency:
                    st.caption(f"  ⚠️ `{ef.name}` 통화 코드 인식 불가 (건너뜀)")

        except Exception as e:
            st.warning(f"⚠️ {ef.name} 파싱 오류: {e}")

    return result


def _build_rate_dict(cur, avg, start='', end=''):
    """수동 입력값으로 환율 딕셔너리 생성 (daily 없음)"""
    return {
        'period':        f"{start} ~ {end}",
        'currency':      cur,
        'currency_name': CURRENCY_NAMES.get(cur, cur),
        'average':       avg,
        'min': avg, 'min_date': '', 'max': avg, 'max_date': '',
        'range': 0.0, 'cross_rate': 0.0, 'daily': [],
    }


# ── 처리 실행 ────────────────────────────────────────────────────
if process_btn and uploaded_files:

    progress_bar = st.progress(0, text="처리 시작...")
    status_area  = st.empty()

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)

            # ── PDF 임시 저장 ──
            progress_bar.progress(10, text="📥 PDF 저장 중...")
            pdf_paths = []
            for f in uploaded_files:
                dest = tmpdir / f.name
                dest.write_bytes(f.read())
                pdf_paths.append(dest)

            # ── PDF 파싱 ──
            progress_bar.progress(25, text="📄 PDF 파싱 중...")
            shopee_results = []
            lazada_result  = None
            qoo10_result   = None
            parse_log      = []

            for pdf_path in pdf_paths:
                pdf_type_hint = detect_pdf_type(str(pdf_path))
                result = parse_pdf(str(pdf_path))

                if result is None:
                    if pdf_type_hint == 'qoo10':
                        parse_log.append(
                            f"⚠️ `{pdf_path.name}` → 큐텐 자동 추출 불가 "
                            f"(STEP 2에서 직접 입력해주세요)"
                        )
                    else:
                        parse_log.append(f"⚠️ `{pdf_path.name}` — 파싱 실패")
                    continue

                ptype = result.get("type")
                if ptype == "shopee":
                    shopee_results.append(result)
                    parse_log.append(
                        f"✅ `{pdf_path.name}` → 쇼피 {result.get('currency','')} "
                        f"{result.get('total_qty',0):,}건 / {result.get('total_amount',0):,.2f} "
                        f"[{result.get('period_start','')}~{result.get('period_end','')}]"
                    )
                elif ptype == "lazada":
                    lazada_result = result
                    n = len(result.get("items", []))
                    parse_log.append(
                        f"✅ `{pdf_path.name}` → 라자다 {n}개국 "
                        f"[{result.get('period_start','')}~{result.get('period_end','')}]"
                    )
                elif ptype == "qoo10":
                    qoo10_result = result
                    parse_log.append(
                        f"✅ `{pdf_path.name}` → 큐텐 OCR 성공 "
                        f"{result.get('qty',0):,}건 / JPY {result.get('amount',0):,.0f} "
                        f"[{result.get('period_start','')}~{result.get('period_end','')}]"
                    )

            # ── 큐텐 수동 입력 적용 ──
            if qoo10_amount > 0:
                base = qoo10_result or {
                    "type": "qoo10", "carrier": "국제로지스틱",
                    "destination": "JP", "currency": "JPY",
                    "period_start": qoo10_period_start,
                    "period_end":   qoo10_period_end,
                    "write_date":   qoo10_write_date_input,
                    "tracking_no":  "",
                }
                base['amount']       = float(qoo10_amount)
                base['qty']          = int(qoo10_qty) if qoo10_qty > 0 else base.get('qty', 0)
                base['tracking_no']  = qoo10_tracking or base.get('tracking_no', '')
                base['period_start'] = qoo10_period_start or base.get('period_start', '')
                base['period_end']   = qoo10_period_end   or base.get('period_end', '')
                # 발행일: 수동 입력 우선, 그 다음 OCR 결과
                base['write_date']   = qoo10_write_date_input or base.get('write_date', '')
                qoo10_result = base
                wd_display = base['write_date'] or '미입력'
                parse_log.append(
                    f"📝 큐텐 수동 입력 적용 — {int(qoo10_qty):,}건 / {int(qoo10_amount):,} JPY "
                    f"(발행일: {wd_display})"
                )

            # 파싱 결과 표시
            with status_area.container():
                st.markdown("**📋 파싱 결과**")
                for log in parse_log:
                    st.markdown(log)

            # ── 거래기간 파악 ──
            all_starts, all_ends = [], []
            for sd in shopee_results:
                if sd.get('period_start'): all_starts.append(sd['period_start'])
                if sd.get('period_end'):   all_ends.append(sd['period_end'])
            if lazada_result:
                if lazada_result.get('period_start'): all_starts.append(lazada_result['period_start'])
                if lazada_result.get('period_end'):   all_ends.append(lazada_result['period_end'])
            if qoo10_result:
                if qoo10_result.get('period_start'): all_starts.append(qoo10_result['period_start'])
                if qoo10_result.get('period_end'):   all_ends.append(qoo10_result['period_end'])

            # 연월 추정
            year, month = None, None
            for p in pdf_paths:
                m = re.search(r"(\d{4})(\d{2})\d{2}", p.name)
                if m:
                    year, month = int(m.group(1)), int(m.group(2))
                    break
            if not year:
                today = datetime.today()
                month = today.month - 1 or 12
                year  = today.year if today.month > 1 else today.year - 1

            fetch_start = min(all_starts) if all_starts else f'{year}-{month:02d}-01'
            fetch_end   = max(all_ends)   if all_ends   else f'{year}-{month:02d}-{calendar.monthrange(year,month)[1]:02d}'

            # ── 환율 처리 ──
            progress_bar.progress(50, text="💱 환율 처리 중...")

            rates = {}

            # 환율 소스 결정: 자동(내장) → SMBS 업로드 → (둘 다 아니면) 직접 입력
            rate_source = None
            if rate_mode == AUTO_RATE_LABEL:
                rate_source = load_fixed_rates()
            elif rate_mode == "📊 SMBS 엑셀 업로드" and smbs_excel_files:
                rate_source = parse_smbs_excel_files(smbs_excel_files)

            if rate_source is not None:
                from modules.exchange_rate import get_rate_for_date
                rate_log = []
                missing  = []

                for cur in CURRENCIES:
                    if cur in rate_source and rate_source[cur].get('daily'):
                        rates[cur] = rate_source[cur]
                        # 발행일(write_date) 기준으로 그날 환율 자동 선택 (없으면 직전 영업일)
                        if cur == 'JPY' and qoo10_result:
                            wd = (qoo10_result.get('write_date', '')
                                  or qoo10_result.get('period_end', fetch_end))
                        elif cur in ('MYR','PHP','SGD','THB','TWD','VND'):
                            sd = next((s for s in shopee_results if s.get('currency') == cur), None)
                            if sd:
                                wd = sd.get('write_date', '') or sd.get('period_end', fetch_end)
                            elif lazada_result:
                                wd = (lazada_result.get('write_date', '')
                                      or lazada_result.get('period_end', fetch_end))
                            else:
                                wd = fetch_end
                        else:
                            wd = fetch_end
                        r = get_rate_for_date(rate_source[cur], wd)
                        rate_log.append(f"✅ {cur}: **{r:.2f}** (발행일 {wd} 기준)")
                    else:
                        rates[cur] = _build_rate_dict(cur, 0.0, fetch_start, fetch_end)
                        rate_log.append(f"⚠️ {cur}: 환율 데이터 없음 → 0")
                        missing.append(cur)

                if missing:
                    with status_area.container():
                        st.markdown("**📋 파싱 결과**")
                        for log in parse_log:
                            st.markdown(log)
                        st.warning(f"환율 데이터 없는 통화: **{', '.join(missing)}** → 0으로 처리됩니다.")

            else:
                # 직접 입력
                rates    = {cur: _build_rate_dict(cur, manual_rates.get(cur, 0.0), fetch_start, fetch_end)
                            for cur in CURRENCIES}
                rate_log = [f"✏️ {cur}: **{manual_rates.get(cur,0.0):.2f}**" for cur in CURRENCIES]

            progress_bar.progress(75, text="📊 엑셀 생성 중...")

            # ── 엑셀 생성 ──
            output_path = tmpdir / f"매출집계_{year}{month:02d}.xlsx"
            generate_excel(
                shopee_results=shopee_results,
                lazada_result=lazada_result,
                qoo10_result=qoo10_result,
                rates=rates,
                output_path=str(output_path),
                year=year,
                month=month,
            )

            progress_bar.progress(100, text="✅ 완료!")
            excel_bytes = output_path.read_bytes()

        # ── 결과 ──
        st.success(f"✅ 엑셀 생성 완료! — {year}년 {month:02d}월")

        with st.expander("💱 적용된 환율 (소포수령증 발행일 기준)"):
            for log in rate_log:
                st.markdown(log)

        st.download_button(
            label=f"⬇️  매출집계_{year}{month:02d}.xlsx  다운로드",
            data=excel_bytes,
            file_name=f"매출집계_{year}{month:02d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

        from openpyxl import load_workbook
        import io
        wb = load_workbook(io.BytesIO(excel_bytes), data_only=True)
        with st.expander("📋 생성된 시트 목록"):
            cols = st.columns(3)
            for i, name in enumerate(wb.sheetnames):
                cols[i%3].markdown(f"• {name}")

    except Exception as e:
        progress_bar.progress(100, text="오류 발생")
        st.error(f"❌ 오류: {e}")
        st.exception(e)


# ── 하단 안내 ────────────────────────────────────────────────────
st.divider()
with st.expander("📌 파일명 규칙 안내"):
    st.markdown("""
| 파일명 패턴 | 플랫폼 |
|---|---|
| `유엠(UM)_MY_*.pdf` | 쇼피 말레이시아 |
| `유엠(UM)_PH_*.pdf` | 쇼피 필리핀 |
| `유엠(UM)_SG_*.pdf` | 쇼피 싱가폴 |
| `유엠(UM)_TH_*.pdf` | 쇼피 태국 |
| `유엠(UM)_TW_*.pdf` | 쇼피 대만 |
| `유엠(UM)_VN_*.pdf` | 쇼피 베트남 |
| `라자다_*.pdf` | 라자다 |
| `큐텐재팬_*.pdf` | 큐텐재팬 (STEP 2에서 수동 입력) |
""")
