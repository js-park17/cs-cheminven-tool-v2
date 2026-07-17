# -*- coding: utf-8 -*-
"""
화학물질정보처리시스템(K-REACH, https://kreach.mcee.go.kr) CAS번호 기반 공개검색 연동.

- 로그인 없이 접근 가능한 '화학물질통합검색' 결과 페이지를
  (recordCountPerPage/searchExcelYn/searchCasnoKeyword 파라미터의 단순 GET 요청)
  CAS번호로 조회하여 KE번호, 인체급성/인체만성/생태 유해성 여부·고시번호,
  사고대비물질 번호, 중점관리물질, 잔류성오염물질, 등록대상 여부를 가져온다.
- 함량기준(%)은 이 목록 화면에는 없고 '정보보기' 상세 팝업에서만 제공되어
  이 모듈만으로는 채울 수 없다 (MSDS 텍스트에서 보조 추출한 값으로 대체 가능).
- 공공기관 서버이므로 호출 간 지연을 두고, 동일 CAS 재조회를 캐싱한다.
"""
import re
import time
import functools

try:
    import requests
except ImportError:
    requests = None

SEARCH_URL = "https://kreach.mcee.go.kr/repwrt/mttr/kr/mttrList.do"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
}
REQUEST_DELAY_SEC = 0.6  # 공공기관 서버 과다호출 방지
TIMEOUT_SEC = 10

TAG_RE = re.compile(r'<[^>]+>')
TR_RE = re.compile(r'<tr[^>]*>(.*?)</tr>', re.DOTALL | re.IGNORECASE)
TD_RE = re.compile(r'<t[dh][^>]*>(.*?)</t[dh]>', re.DOTALL | re.IGNORECASE)
CAS_RE = re.compile(r'\b\d{2,7}-\d{2}-\d\b')


def _strip_tags(html_fragment):
    text = TAG_RE.sub(' ', html_fragment)
    text = text.replace('&nbsp;', ' ').replace('&amp;', '&')
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def _split_type_and_no(cell_text):
    """'인체급성 97-1-203' / '인체급성·인체만성·생태 97-1-271(1)' 형태 파싱.
    반환: (types:set[str], notice_no:str)"""
    cell_text = cell_text.strip()
    if not cell_text:
        return set(), ""
    m = re.match(r'^([가-힣·\s]+?)\s+([\d].*)$', cell_text)
    if m:
        type_part = m.group(1)
        no_part = m.group(2).strip()
    else:
        type_part = cell_text
        no_part = ""
    types = set()
    if '인체급성' in type_part:
        types.add('인체급성')
    if '인체만성' in type_part:
        types.add('인체만성')
    if '생태' in type_part:
        types.add('생태')
    return types, no_part


@functools.lru_cache(maxsize=512)
def search_by_cas(cas_no):
    """CAS번호로 K-REACH 화학물질통합검색을 조회. 실패/미등록시 status로 구분."""
    result = {
        "cas": cas_no, "status": "error", "ke_no": "", "name_kr": "", "name_en": "",
        "acute_no": "", "chronic_no": "", "eco_no": "",
        "accident_prep_no": "", "restricted_raw": "", "priority_mgmt_no": "",
        "pop_raw": "", "registration_status": "", "raw_row_text": "",
        "detail_note": "",
    }
    if requests is None:
        result["detail_note"] = "requests 패키지가 설치되어 있지 않습니다."
        return result

    params = {"recordCountPerPage": 10, "searchExcelYn": "N", "searchCasnoKeyword": cas_no}
    try:
        resp = requests.get(SEARCH_URL, params=params, headers=HEADERS, timeout=TIMEOUT_SEC)
        resp.raise_for_status()
    except Exception as e:
        result["detail_note"] = f"K-REACH 조회 실패: {e}"
        return result
    finally:
        time.sleep(REQUEST_DELAY_SEC)

    html = resp.text
    rows = TR_RE.findall(html)
    data_row = None
    for row_html in rows:
        if cas_no not in row_html:
            continue
        cells = [_strip_tags(c) for c in TD_RE.findall(row_html)]
        if len(cells) < 8:
            continue
        if CAS_RE.match(cells[0].strip()):
            data_row = cells
            break

    if data_row is None:
        result["status"] = "not_found"
        result["detail_note"] = "K-REACH 검색결과 없음(신규화학물질이거나 CAS 표기 상이 가능)"
        return result

    try:
        result["name_en"] = data_row[1]
        result["name_kr"] = data_row[2]
        result["ke_no"] = data_row[3]
        hazard_cell = data_row[4]
        types, notice_no = _split_type_and_no(hazard_cell)
        if '인체급성' in types:
            result["acute_no"] = notice_no
        if '인체만성' in types:
            result["chronic_no"] = notice_no
        if '생태' in types:
            result["eco_no"] = notice_no
        result["accident_prep_no"] = data_row[5]
        result["restricted_raw"] = data_row[6]
        result["priority_mgmt_no"] = data_row[7]
        result["pop_raw"] = data_row[8] if len(data_row) > 8 else ""
        result["registration_status"] = data_row[9] if len(data_row) > 9 else ""
        result["raw_row_text"] = " | ".join(data_row)
        result["status"] = "ok"
    except IndexError:
        result["status"] = "parse_error"
        result["detail_note"] = "K-REACH 응답 표 구조가 예상과 달라 자동파싱 실패"
    return result


def search_many(cas_list, progress_cb=None):
    """CAS번호 리스트를 순회 조회. 중복은 캐시로 재사용."""
    out = {}
    unique = [c for c in dict.fromkeys(cas_list) if c]
    for i, cas in enumerate(unique):
        out[cas] = search_by_cas(cas)
        if progress_cb:
            progress_cb(i + 1, len(unique), cas)
    return out
