"""
서울외국환중개(SMBS) 환율 자동 수집 모듈
URL: http://www.smbs.biz/ExRate/StdExRate.jsp

환율 적용 기준: 거래기간 마지막날(period_end) 환율
"""

import requests
from bs4 import BeautifulSoup
import re
import calendar
from typing import Optional
from datetime import datetime, timedelta


SMBS_CURRENCY_NAMES = {
    'MYR': '말레이시아 링깃 (MYR)',
    'PHP': '필리핀 페소 (PHP)',
    'SGD': '싱가포르 달러 (SGD)',
    'THB': '태국 바트 (THB)',
    'TWD': '대만 달러 (TWD)',
    'VND': '베트남 동 (VND)',
    'JPY': '일본 엔 (JPY) (100)',
    'BRL': '브라질 헤알 (BRL)',
}

SMBS_BASE = 'http://www.smbs.biz'
SMBS_URL  = f'{SMBS_BASE}/ExRate/StdExRate.jsp'

HEADERS = {
    'User-Agent': (
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/120.0.0.0 Safari/537.36'
    ),
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
    'Accept-Encoding': 'gzip, deflate',
    'Connection': 'keep-alive',
    'Referer': SMBS_URL,
}


def fetch_rates_for_date_range(start_str: str, end_str: str, currency: str) -> Optional[dict]:
    """
    SMBS에서 임의 날짜 범위의 환율 데이터 수집.
    start_str, end_str: 'YYYY-MM-DD' 형식
    실패 시 None 반환.
    """
    # YYYYMMDD 형식으로 변환
    start_date = start_str.replace('-', '').replace('.', '')[:8]
    end_date   = end_str.replace('-', '').replace('.', '')[:8]

    # 기간 표시용 문자열
    sy, sm, sd = start_date[:4], start_date[4:6], start_date[6:8]
    ey, em, ed = end_date[:4], end_date[4:6], end_date[6:8]
    period_str = f"{sy}년 {sm}월 {sd}일 ~ {ey}년 {em}월 {ed}일"

    # GET → POST 순으로 시도
    result = _try_get(start_date, end_date, currency, period_str)
    if result:
        return result
    result = _try_post(start_date, end_date, currency, period_str)
    if result:
        return result
    return None


def fetch_rates_for_month(year: int, month: int, currency: str) -> Optional[dict]:
    """
    SMBS에서 월별 환율 데이터 수집.
    실패 시 None 반환.
    """
    last_day   = calendar.monthrange(year, month)[1]
    start_str  = f'{year}-{month:02d}-01'
    end_str    = f'{year}-{month:02d}-{last_day:02d}'
    return fetch_rates_for_date_range(start_str, end_str, currency)


def _try_get(start_date, end_date, currency, period_str):
    try:
        params = {
            'yyyymmdd1': start_date,
            'yyyymmdd2': end_date,
            'curCd': currency,
        }
        resp = requests.get(SMBS_URL, params=params, headers=HEADERS, timeout=15)
        resp.raise_for_status()
        resp.encoding = resp.apparent_encoding or 'utf-8'
        return _parse_html(resp.text, currency, period_str)
    except Exception:
        return None


def _try_post(start_date, end_date, currency, period_str):
    try:
        session = requests.Session()
        session.headers.update(HEADERS)
        session.get(SMBS_URL, timeout=10)  # 쿠키 획득

        form_data = {
            'yyyymmdd1': start_date,
            'yyyymmdd2': end_date,
            'curCd': currency,
            'gubun': '1',
        }
        resp = session.post(SMBS_URL, data=form_data, timeout=15)
        resp.raise_for_status()
        resp.encoding = resp.apparent_encoding or 'utf-8'
        return _parse_html(resp.text, currency, period_str)
    except Exception:
        return None


def _parse_html(html: str, currency: str, period_str: str) -> Optional[dict]:
    """SMBS HTML → 환율 딕셔너리 파싱"""
    soup = BeautifulSoup(html, 'html.parser')
    daily_rates = []
    avg_rate = min_val = max_val = range_val = cross_avg = 0.0
    min_date = max_date = ''

    for table in soup.find_all('table'):
        for row in table.find_all('tr'):
            cells = [td.get_text(strip=True) for td in row.find_all(['td', 'th'])]
            if not cells:
                continue

            # 일별 환율 행: 첫 셀이 날짜(YYYY.MM.DD)
            if re.match(r'\d{4}\.\d{2}\.\d{2}', cells[0]):
                try:
                    rate   = float(cells[2].replace(',', ''))
                    change = float(cells[3].replace(',', '')) if len(cells) > 3 else 0.0
                    cross  = float(cells[4].replace(',', '')) if len(cells) > 4 else 0.0
                    daily_rates.append({
                        'date': cells[0], 'rate': rate,
                        'change': change, 'cross': cross,
                    })
                except (ValueError, IndexError):
                    pass

            # 평균환율 통계 행
            else:
                float_vals = []
                date_vals  = []
                for c in cells:
                    clean = c.replace(',', '')
                    try:
                        float_vals.append(float(clean))
                    except ValueError:
                        if re.match(r'\d{4}\.\d{2}\.\d{2}', c):
                            date_vals.append(c)
                if len(float_vals) >= 5 and avg_rate == 0.0:
                    avg_rate  = float_vals[0]
                    min_val   = float_vals[1]
                    max_val   = float_vals[2]
                    range_val = float_vals[3]
                    cross_avg = float_vals[4]
                    if len(date_vals) >= 1: min_date = date_vals[0]
                    if len(date_vals) >= 2: max_date = date_vals[1]

    if not daily_rates:
        return None

    # 평균 직접 계산 (파싱 실패 시)
    if avg_rate == 0.0:
        rs = [d['rate'] for d in daily_rates]
        avg_rate  = round(sum(rs) / len(rs), 2)
        min_val   = min(rs)
        max_val   = max(rs)
        range_val = round(max_val - min_val, 2)
        min_date  = next((d['date'] for d in daily_rates if d['rate'] == min_val), '')
        max_date  = next((d['date'] for d in daily_rates if d['rate'] == max_val), '')

    return {
        'period':        period_str,
        'currency':      currency,
        'currency_name': SMBS_CURRENCY_NAMES.get(currency, currency),
        'average':       avg_rate,
        'min':           min_val,  'min_date': min_date,
        'max':           max_val,  'max_date': max_date,
        'range':         range_val,
        'cross_rate':    cross_avg,
        'daily':         daily_rates,
    }


def get_rate_for_date(rate_data: dict, date_str: str) -> float:
    """
    특정 날짜의 환율 반환.
    없으면 가장 가까운 이전 영업일 환율 사용.
    rate_data가 없거나 daily가 비어있으면 average 반환 (수동입력값).
    """
    if not rate_data:
        return 0.0

    daily = rate_data.get('daily', [])

    # daily 데이터가 없으면 average 반환 (수동입력 모드)
    if not daily:
        return rate_data.get('average', 0.0)

    if not date_str:
        return rate_data.get('average', 0.0)

    # 날짜 정규화 (YYYY-MM-DD 또는 YYYY.MM.DD → YYYY.MM.DD)
    date_norm = date_str.replace('-', '.').replace('/', '.')
    # YYYYMMDD 형식이면 변환
    if re.match(r'^\d{8}$', date_norm.replace('.', '')):
        d = date_norm.replace('.', '')
        date_norm = f"{d[:4]}.{d[4:6]}.{d[6:8]}"

    # 정확히 일치하는 날짜 찾기
    for d in daily:
        if d['date'] == date_norm:
            return d['rate']

    # 없으면 가장 가까운 이전 영업일 환율 반환
    sorted_daily = sorted(daily, key=lambda x: x['date'])
    prev_rate = sorted_daily[0]['rate'] if sorted_daily else 0.0
    for d in sorted_daily:
        if d['date'] <= date_norm:
            prev_rate = d['rate']
        else:
            break
    return prev_rate


def get_period_end_rate(rate_data: dict, period_end: str) -> float:
    """
    거래기간 마지막날 환율 반환.
    마지막날이 주말/공휴일이면 가장 가까운 이전 영업일 환율 사용.
    """
    return get_rate_for_date(rate_data, period_end)


def fetch_all_currencies(year: int, month: int, currencies: list) -> dict:
    """월 기준으로 전체 통화 환율 수집"""
    result = {}
    failed = []
    for cur in currencies:
        data = fetch_rates_for_month(year, month, cur)
        result[cur] = data
        if data:
            print(f'  ✅ {cur} 환율 수집 완료 (평균 {data["average"]})')
        else:
            failed.append(cur)
    if failed:
        print(f'  ⚠️  수집 실패: {", ".join(failed)} → 직접 입력 필요')
    return result


def fetch_all_currencies_for_range(start_str: str, end_str: str, currencies: list) -> dict:
    """날짜 범위 기준으로 전체 통화 환율 수집"""
    result = {}
    failed = []
    for cur in currencies:
        data = fetch_rates_for_date_range(start_str, end_str, cur)
        result[cur] = data
        if data:
            print(f'  ✅ {cur} 환율 수집 완료 (기간: {start_str}~{end_str})')
        else:
            failed.append(cur)
    if failed:
        print(f'  ⚠️  수집 실패: {", ".join(failed)} → 직접 입력 필요')
    return result
