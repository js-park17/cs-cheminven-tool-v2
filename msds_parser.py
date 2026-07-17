# -*- coding: utf-8 -*-
"""
MSDS(물질안전보건자료) PDF 파서
- 국내 K-MSDS 16항목 표준 서식(제조사별 표기 차이 3종 이상 지원)에서
  6호서식/화학물질인벤토리 작성에 필요한 항목을 최대한 자동 추출한다.
- 추출이 불확실/불가능한 항목은 needs_review 플래그와 원문 근거를 남겨 수동 검토를 유도한다.
"""
import re
import io

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

CAS_RE = re.compile(r'\b(\d{2,7}-\d{2}-\d)\b')
WATER_CAS = "7732-18-5"

# ---------------------------------------------------------------------------
# GHS 유해성분류 표준 코드 매핑
# ---------------------------------------------------------------------------
GHS_CODE_MAP = {
    "폭발성물질": "2.1", "인화성가스": "2.2", "에어로졸": "2.3", "산화성가스": "2.4",
    "고압가스": "2.5", "인화성액체": "2.6", "인화성고체": "2.7", "자기반응성물질": "2.8",
    "자연발화성액체": "2.9", "자연발화성고체": "2.10", "자기발열성물질": "2.11",
    "물반응성물질": "2.12", "산화성액체": "2.13", "산화성고체": "2.14", "유기과산화물": "2.15",
    "금속부식성물질": "2.16",
    "급성독성(경구)": "3.1", "급성독성(경피)": "3.1", "급성독성(흡입)": "3.1", "급성독성": "3.1",
    "피부부식성/피부자극성": "3.2", "피부부식성/자극성물질": "3.2", "피부부식성/자극성": "3.2",
    "심한눈손상성/눈자극성": "3.3", "심한눈손상또는눈자극성물질": "3.3", "심한눈손상/눈자극성": "3.3",
    "호흡기과민성": "3.4", "피부과민성": "3.4",
    "생식세포변이원성": "3.5", "발암성": "3.6", "생식독성": "3.7",
    "표적장기전신독성물질(1회노출)": "3.8", "특정표적장기독성-1회노출": "3.8", "표적장기-1회노출": "3.8",
    "표적장기전신독성물질(반복노출)": "3.9", "특정표적장기독성-반복노출": "3.9", "표적장기-반복노출": "3.9",
    "흡인유해성": "3.10",
    "수생환경유해성": "4.1", "만성수생환경유해성": "4.1", "급성수생환경유해성": "4.1", "오존층유해성": "4.2",
}

def _norm_hazard_key(name):
    return re.sub(r'[\s·,.:]', '', name)

def ghs_code(name):
    return GHS_CODE_MAP.get(_norm_hazard_key(name))

# ---------------------------------------------------------------------------
# CAS -> 국문 관용명 (표에서 영문명만 확인되는 경우의 보조 사전 + 흔한 산업물질)
# ---------------------------------------------------------------------------
CAS_KOREAN_NAME = {
    "71-43-2": "벤젠", "108-88-3": "톨루엔", "1330-20-7": "크실렌(혼합자일렌)",
    "106-42-3": "p-자일렌", "64-19-7": "초산(아세트산)", "67-64-1": "아세톤",
    "64-17-5": "에틸알코올(에탄올)", "67-56-1": "메틸알코올(메탄올)",
    "67-63-0": "이소프로필알콜(IPA)", "7664-93-9": "황산", "7647-01-0": "염산",
    "7664-39-3": "불산(플루오르화수소)", "1336-21-6": "암모니아수",
    "7727-37-9": "질소", "7732-18-5": "물", "78-93-3": "메틸에틸케톤(MEK)",
    "7553-56-2": "요오드", "12027-06-4": "암모늄요오드화합물", "7664-41-7": "암모니아",
}

def guess_korean_name(cas, fallback_name=""):
    return CAS_KOREAN_NAME.get(cas, fallback_name)

# ---------------------------------------------------------------------------
# 섹션 분리
# ---------------------------------------------------------------------------
SECTION_ANCHORS = [
    ("s1", "화학제품과 회사에 관한 정보"),
    ("s2", "유해"),
    ("s3", "구성성분의 명칭"),
    ("s4", "응급조치"),
    ("s5", "폭발"),
    ("s6", "누출"),
    ("s7", "취급 및"),
    ("s8", "노출"),
    ("s9", "물리"),
    ("s10", "안정성"),
    ("s11", "독성에 관한"),
    ("s12", "환경에"),
    ("s13", "폐기"),
    ("s14", "운송"),
    ("s15", "법적 규제"),
    ("s16", "기타 참고"),
]

def extract_pdf_text(file_bytes_or_path):
    if pdfplumber is None:
        raise RuntimeError("pdfplumber 가 설치되어 있지 않습니다.")
    if isinstance(file_bytes_or_path, (bytes, bytearray)):
        f = io.BytesIO(file_bytes_or_path)
    else:
        f = file_bytes_or_path
    full = ""
    with pdfplumber.open(f) as pdf:
        for page in pdf.pages:
            full += (page.extract_text() or "") + "\n"
    return full

def _anchor_positions(full_text):
    positions = {}
    search_from = 0
    for key, kw in SECTION_ANCHORS:
        idx = full_text.find(kw, search_from)
        positions[key] = idx if idx != -1 else None
        if idx != -1:
            search_from = idx + len(kw)
    return positions

def get_section_text(full_text, key):
    order = [k for k, _ in SECTION_ANCHORS]
    positions = _anchor_positions(full_text)
    if positions.get(key) is None:
        return ""
    idx = order.index(key)
    start = positions[key]
    end = len(full_text)
    for nxt in order[idx + 1:]:
        if positions.get(nxt) is not None:
            end = positions[nxt]
            break
    return full_text[start:end]

# ---------------------------------------------------------------------------
# 공통: 라벨을 포함하는 줄을 찾아 라벨 뒤 값(괄호 단위 표기 제거)을 반환
# ---------------------------------------------------------------------------
def _find_labeled_value(section_text, label, allow_multi=False, exclude_prefixes=None):
    results = []
    exclude_prefixes = exclude_prefixes or []
    for raw_line in section_text.split('\n'):
        idx = raw_line.find(label)
        if idx == -1:
            continue
        if any(raw_line.find(ex) != -1 and raw_line.find(ex) <= idx for ex in exclude_prefixes):
            continue
        rest = raw_line[idx + len(label):]
        rest = re.sub(r'^\s*(\([^)]*\)\s*)+', '', rest)  # 괄호 단위 표기 제거
        rest = rest.lstrip(' :：').strip()
        if rest:
            if allow_multi:
                results.append((rest, raw_line))
            else:
                return rest, raw_line
    if allow_multi:
        return results
    return None, None

def _first_num(text):
    if text is None:
        return None
    m = re.search(r'([\-]?\d+(?:\.\d+)?)', text)
    return float(m.group(1)) if m else None

# ---------------------------------------------------------------------------
# 1. 제품명 / 용도(제품분류) / 최초작성일자(제정일자)
# ---------------------------------------------------------------------------
def parse_product_name(full_text):
    m = re.search(r'제품명\s*[:：]?\s*([^\n]+)', full_text)
    if m:
        return m.group(1).strip()
    return ""

def parse_usage(full_text):
    """제품의 권고 용도 추출 (제품분류란에 사용). 두 서식 모두 지원:
    A) '권고 용도와 사용상의 제한 : 반도체 제조용' (한 줄, 콜론 뒤 값)
    B) '○ 권고용도 반도체 제조용' (별도 줄, 콜론 없음)"""
    m = re.search(r'권고\s*용도와\s*사용상의\s*제한\s*[:：]\s*([^\n]+)', full_text)
    if m:
        usage = m.group(1).strip()
        usage = re.split(r'\s*(?:나\.|다\.|사용상의\s*제한)', usage)[0].strip()
        return usage
    m = re.search(r'○\s*권고\s*용도\s*[:：]?\s*([^\n]+)', full_text)
    if m:
        return m.group(1).strip()
    return ""

def parse_enact_date(full_text):
    """최초 작성일자(제정일자) 추출"""
    m = re.search(r'최초\s*작성일자[^\d]{0,6}(\d{4}[.\-]\s?\d{1,2}[.\-]\s?\d{1,2})', full_text)
    if m:
        return re.sub(r'\s+', '', m.group(1)).rstrip('.')
    return ""

def parse_msds_meta(full_text):
    """MSDS등록번호(문서번호) / 개정횟수(버전) / 최종개정일자 추출"""
    meta = {"msds_reg_no": "", "revision_no": "", "revision_date": ""}
    m = re.search(r'\b([A-Z]{2}\d{5}-\d{10})\b', full_text)
    if m:
        meta["msds_reg_no"] = m.group(1)
    m = re.search(r'개정\s*(?:회수|횟수)[^\d]{0,10}(\d+)', full_text)
    if m:
        meta["revision_no"] = m.group(1)
    m = re.search(r'개정일자[\s\S]{0,40}?(\d{4}[.\-]\s?\d{1,2}[.\-]\s?\d{1,2})', full_text)
    if m:
        meta["revision_date"] = re.sub(r'\s+', '', m.group(1)).rstrip('.')
    return meta

# ---------------------------------------------------------------------------
# 3. 구성성분 (CAS / 함유량 / 영문명)
# ---------------------------------------------------------------------------
RANGE_PCT_RE = re.compile(r'(\d+(?:\.\d+)?)\s*[~\-]\s*(\d+(?:\.\d+)?)\s*%?')
SINGLE_PCT_RE = re.compile(r'(\d+(?:\.\d+)?)\s*%')
WORDY_PCT_RE = re.compile(r'(\d+(?:\.\d+)?)\s*(이상|이하|초과|미만)')

def parse_composition(section3_text):
    """반환: [{cas, name_en, content_pct}] , CAS 등장 순서를 보존."""
    rows = []
    cas_list = list(CAS_RE.finditer(section3_text))
    for i, m in enumerate(cas_list):
        cas = m.group(1)
        start = m.end()
        end = cas_list[i + 1].start() if i + 1 < len(cas_list) else min(len(section3_text), m.end() + 200)
        window = section3_text[start:end]

        content = ""
        rm = RANGE_PCT_RE.search(window)
        if rm:
            content = f"{rm.group(1)}~{rm.group(2)}%"
        else:
            sm = SINGLE_PCT_RE.search(window)
            if sm:
                content = f"{sm.group(1)}%"
            else:
                wm = WORDY_PCT_RE.search(window)
                if wm:
                    content = f"{wm.group(1)}{wm.group(2)}"

        pct_pos = rm.start() if rm else (len(window))
        candidate_zone = window[:pct_pos]
        words = re.findall(r'[A-Za-z][A-Za-z\-]{2,}(?:\s[A-Za-z][A-Za-z\-]{2,})*', candidate_zone)
        words = [w for w in words if len(w) > 2]
        name_en = max(words, key=len).strip() if words else ""

        rows.append({"cas": cas, "name_en": name_en, "content_pct": content})
    return rows

# ---------------------------------------------------------------------------
# 9. 물리화학적 특성
# ---------------------------------------------------------------------------
def parse_physchem(section9_text):
    result = {
        "state": "", "specific_gravity": None, "flash_point_c": None,
        "flash_point_raw": "", "explosion_lower": None, "explosion_upper": None,
        "vapor_pressure_mmhg": None, "ph": None,
    }

    val, _ = _find_labeled_value(section9_text, "성상")
    if val is None:
        val, _ = _find_labeled_value(section9_text, "외관")
    if val:
        if '기체' in val or '가스' in val:
            result["state"] = "기체"
        elif '고체' in val or '분말' in val or '고상' in val:
            result["state"] = "고체"
        elif '액체' in val or '액상' in val or '액화가스' in val:
            result["state"] = "액체" if '액화가스' not in val else "기체"

    val, _ = _find_labeled_value(section9_text, "비중")
    if val and ('자료없음' in val or '자료 없음' in val or '해당없음' in val or '해당 없음' in val):
        val = None
    if val is None:
        val, _ = _find_labeled_value(section9_text, "밀도", exclude_prefixes=["증기밀도"])
        if val and ('자료없음' in val or '자료 없음' in val or '해당없음' in val or '해당 없음' in val):
            val = None
    result["specific_gravity"] = _first_num(val) if val else None

    val, _ = _find_labeled_value(section9_text, "인화점")
    if val:
        result["flash_point_raw"] = val
        if '자료없음' in val or '자료 없음' in val:
            result["flash_point_c"] = None
        elif '불연성' in val or '해당없음' in val or '해당 없음' in val:
            result["flash_point_c"] = "불연성"
        else:
            cm = re.search(r'([\-]?\d+(?:\.\d+)?)\s*℃', val)
            result["flash_point_c"] = float(cm.group(1)) if cm else _first_num(val)

    val, _ = _find_labeled_value(section9_text, "pH")
    if val and not any(x in val for x in ('자료없음', '자료 없음', '해당없음', '해당 없음')):
        result["ph"] = _first_num(val)

    m = re.search(r'(인화\s*(?:또는|,)?\s*폭발\s*범위의?\s*상한\s*/\s*하한)[^\n]*', section9_text)
    if m:
        line = m.group(0)
        line_wo_label = line[len(m.group(1)):]
        line_wo_label = re.sub(r'\([^)]*\)', '', line_wo_label)
        if '자료없음' not in line_wo_label and '해당없음' not in line_wo_label:
            nums = re.findall(r'[\-]?\d+(?:\.\d+)?', line_wo_label)
            nums = [float(n) for n in nums]
            if len(nums) >= 2:
                result["explosion_lower"] = min(nums[0], nums[1])
                result["explosion_upper"] = max(nums[0], nums[1])
            elif len(nums) == 1:
                result["explosion_upper"] = nums[0]

    for raw_line in section9_text.split('\n'):
        idx = raw_line.find('증기압')
        if idx == -1:
            continue
        rest = raw_line[idx + len('증기압'):]
        unit_m = re.match(r'\s*(\([^)]*\))?', rest)
        unit_hint = unit_m.group(1) or "" if unit_m else ""
        value_part = rest[unit_m.end():] if unit_m else rest
        value_part = value_part.lstrip(' :：').strip()
        if not value_part or '자료없음' in value_part or '해당없음' in value_part:
            break
        num = _first_num(value_part)
        if num is None:
            break
        unit_text = (unit_hint + " " + value_part).lower()
        if 'mmhg' in unit_text:
            mmhg = num
        elif 'atm' in unit_text:
            mmhg = num * 760.0
        elif 'kpa' in unit_text or '㎪' in unit_text:
            mmhg = num * 7.50062
        elif 'hpa' in unit_text or 'mbar' in unit_text:
            mmhg = num * 0.750062
        elif 'pa' in unit_text:
            mmhg = num * 0.00750062
        else:
            mmhg = num
        result["vapor_pressure_mmhg"] = round(mmhg, 2)
        break

    return result

# ---------------------------------------------------------------------------
# 2. 유해위험성 분류 (독성구분 항목/등급)
# ---------------------------------------------------------------------------
def parse_hazard_classification(section2_text):
    """반환: [(항목명, GHS코드 또는 None, 등급)]"""
    items = []

    block_m = re.search(r'분류항목\s*구분(.*?)(?:나\.|○\s*그림문자)', section2_text, re.DOTALL)
    if block_m:
        block = block_m.group(1)
        for line in block.split('\n'):
            line = line.strip()
            if not line.startswith('-'):
                continue
            body = line[1:].strip()
            gm = re.match(r'^(.*\S)\s+([0-9]+[A-Za-z]{0,2})$', body)
            if gm:
                name = gm.group(1).strip()
                grade = gm.group(2).strip()
                items.append((name, ghs_code(name), grade))

    if not items:
        for m in re.finditer(r'^([가-힣][가-힣A-Za-z/\(\)\s\-]{1,30}?)\s*[:：]?\s*(?:\(([\d.]+)\))?\s*구분\s*([0-9A-Za-z]+)\s*$',
                              section2_text, re.MULTILINE):
            name = m.group(1).strip()
            code = m.group(2) or ghs_code(name)
            items.append((name, code, m.group(3)))

    return items

# ---------------------------------------------------------------------------
# 15. 법적 규제 현황 (화학물질관리법) - MSDS 원문 보조 파싱 (K-REACH가 1차 소스)
# ---------------------------------------------------------------------------
def parse_chemcontrol_regulation(section15_text):
    result = {
        "hazard_chem_no": "", "accident_prep_no": "", "restricted": False,
        "prohibited": False, "raw_note": "", "regulation_summary": "",
    }
    text = section15_text

    def _clean(v):
        v = v.strip()
        v = re.sub(r'\s*-\s*$', '', v).strip()
        return v

    def _is_empty(v):
        return re.fullmatch(r'[\s\-,]*', v or '') is not None

    m = re.search(r'(?:^|\n)\s*\d\s*유해화학물질\s+([^\n]+)', text)
    if m:
        val = _clean(m.group(1))
        if not _is_empty(val):
            result["hazard_chem_no"] = val

    m = re.search(r'(?:^|\n)\s*\d\s*사고대비물질\s+([^\n]+)', text)
    if m:
        val = _clean(m.group(1))
        if not _is_empty(val):
            result["accident_prep_no"] = val

    m = re.search(r'(?:^|\n)\s*\d\s*제한물질\s+([^\n]+)', text)
    if m:
        result["restricted"] = not _is_empty(_clean(m.group(1)))

    m = re.search(r'(?:^|\n)\s*\d\s*금지물질\s+([^\n]+)', text)
    if m:
        result["prohibited"] = not _is_empty(_clean(m.group(1)))

    if not result["hazard_chem_no"] and not result["accident_prep_no"]:
        m = re.search(r'화학물질관리법에\s*의한\s*규제\s*[:：]?\s*([^\n]+)', text)
        if m:
            val = m.group(1).strip()
            result["raw_note"] = val
            if '해당없음' not in val.replace(" ", "") and '해당 없음' not in val:
                num_m = re.search(r'\d{1,4}-\d-\d{1,4}', val)
                if num_m:
                    result["hazard_chem_no"] = val

    parts = []
    if result["hazard_chem_no"]:
        parts.append(f"유해화학물질({result['hazard_chem_no'].split(',')[0]})")
    if result["accident_prep_no"]:
        parts.append(f"사고대비물질({result['accident_prep_no'].split(',')[0]})")
    if result["restricted"]:
        parts.append("제한물질")
    if result["prohibited"]:
        parts.append("금지물질")
    result["regulation_summary"] = ", ".join(parts) if parts else "해당없음"
    return result

# ---------------------------------------------------------------------------
# 위험노출수준: 끝점농도(응급노출기준) - ERPG-2 > AEGL-2 > PAC-2 > IDLH 우선순위
# ---------------------------------------------------------------------------
ENDPOINT_PATTERNS = [
    ("ERPG-2", re.compile(r'ERPG[\s\-]?2\D{0,10}([\d,]+(?:\.\d+)?)\s*(ppm|mg/m3|mg/㎥|㎎/㎥)', re.IGNORECASE)),
    ("AEGL-2", re.compile(r'AEGL[\s\-]?2\D{0,15}([\d,]+(?:\.\d+)?)\s*(ppm|mg/m3|mg/㎥|㎎/㎥)', re.IGNORECASE)),
    ("PAC-2", re.compile(r'PAC[\s\-]?2\D{0,10}([\d,]+(?:\.\d+)?)\s*(ppm|mg/m3|mg/㎥|㎎/㎥)', re.IGNORECASE)),
    ("IDLH", re.compile(r'IDLH\D{0,10}([\d,]+(?:\.\d+)?)\s*(ppm|mg/m3|mg/㎥|㎎/㎥)', re.IGNORECASE)),
]

def parse_endpoint_concentration(full_text):
    """ERPG-2 > AEGL-2 > PAC-2 > IDLH 순으로 최초로 발견되는 끝점농도를 위험노출수준으로 채택"""
    for label, pattern in ENDPOINT_PATTERNS:
        m = pattern.search(full_text)
        if m:
            return f"{label}: {m.group(1)} {m.group(2)}"
    return ""

# ---------------------------------------------------------------------------
# 허용농도값: 11.독성에 관한 정보 - TWA > LD50 > LC50 우선순위
# ---------------------------------------------------------------------------
TWA_RE = re.compile(r'TWA\s*[:：]?\s*([\d,]+(?:\.\d+)?)\s*(ppm|mg/m3|mg/㎥|㎎/㎥)', re.IGNORECASE)
LD50_INLINE_RE = re.compile(
    r'\bLD\s*50\s*([\d,]+(?:\.\d+)?)\s*(㎎/㎏|mg/kg|mg/㎏|㎍/㎏)\s*(?:\(([^)]+)\)|,?\s*([A-Za-z가-힣]+))?')
LD50_BROKEN_RE = re.compile(
    r'\bLD\s+([\d,]+(?:\.\d+)?)\s*(㎎/㎏|mg/kg|mg/㎏)([^\n]{0,40})\n\s*(?:mix\s*)?50\b')
LC50_INLINE_RE = re.compile(
    r'\bLC\s*50\s*([\d,]+(?:\.\d+)?)\s*(ppm|㎎/[㎥ℓLl]|mg/m3|mg/L)\s*(?:\(([^)]+)\)|,?\s*([A-Za-z가-힣]+))?')
LC50_BROKEN_RE = re.compile(
    r'\bLC\s+([\d,]+(?:\.\d+)?)\s*(ppm|㎎/[㎥ℓLl]|mg/m3|mg/L)([^\n]{0,40})\n\s*(?:mix\s*)?50\b')

def parse_toxicity_reference_value(section11_text):
    """TWA > LD50 > LC50 순으로 11.독성에 관한 정보에서 값 채택"""
    m = TWA_RE.search(section11_text)
    if m:
        return f"TWA {m.group(1)} {m.group(2)}"

    m = LD50_INLINE_RE.search(section11_text)
    if m:
        species = (m.group(3) or m.group(4) or "").strip()
        tail = f" ({species})" if species else ""
        return f"LD50 {m.group(1)}{m.group(2)}{tail}"
    m = LD50_BROKEN_RE.search(section11_text)
    if m:
        species_m = re.search(r'[A-Za-z가-힣]+', m.group(3) or "")
        tail = f" ({species_m.group(0)})" if species_m else ""
        return f"LD50 {m.group(1)}{m.group(2)}{tail}"

    m = LC50_INLINE_RE.search(section11_text)
    if m:
        species = (m.group(3) or m.group(4) or "").strip()
        tail = f" ({species})" if species else ""
        return f"LC50 {m.group(1)}{m.group(2)}{tail}"
    m = LC50_BROKEN_RE.search(section11_text)
    if m:
        species_m = re.search(r'[A-Za-z가-힣]+', m.group(3) or "")
        tail = f" ({species_m.group(0)})" if species_m else ""
        return f"LC50 {m.group(1)}{m.group(2)}{tail}"
    return ""

# ---------------------------------------------------------------------------
# 부식성 여부
# ---------------------------------------------------------------------------
def infer_corrosive(hazard_items, physchem):
    for name, code, grade in hazard_items:
        key = _norm_hazard_key(name)
        if '부식성' in key:
            return "유"
    ph = physchem.get("ph")
    if ph is not None:
        if ph <= 2 or ph >= 11.5:
            return "유"
    return "무"

# ---------------------------------------------------------------------------
# 종합 파서
# ---------------------------------------------------------------------------
def parse_msds(file_bytes, filename=""):
    full_text = extract_pdf_text(file_bytes)
    product_name = parse_product_name(full_text)
    meta = parse_msds_meta(full_text)
    usage = parse_usage(full_text)
    enact_date = parse_enact_date(full_text)

    sec2 = get_section_text(full_text, "s2")
    sec3 = get_section_text(full_text, "s3")
    sec9 = get_section_text(full_text, "s9")
    sec11 = get_section_text(full_text, "s11")
    sec15 = get_section_text(full_text, "s15")

    composition = parse_composition(sec3)
    physchem = parse_physchem(sec9)
    hazards = parse_hazard_classification(sec2)
    endpoint_conc = parse_endpoint_concentration(full_text)
    toxicity_ref = parse_toxicity_reference_value(sec11)
    corrosive = infer_corrosive(hazards, physchem)

    components = [c for c in composition if c["cas"] != WATER_CAS]
    if not components and composition:
        components = composition

    is_mixture = len(components) > 1
    if is_mixture:
        regulation = {
            "hazard_chem_no": "", "accident_prep_no": "",
            "regulation_summary": "다성분 혼합물 - 성분별 화관법 규제여부 수동확인 필요",
        }
    else:
        regulation = parse_chemcontrol_regulation(sec15)

    rows = []
    for comp in components:
        kr_name = guess_korean_name(comp["cas"], comp["name_en"] or product_name)
        name_needs_review = kr_name == comp["name_en"] and comp["name_en"] != ""
        rows.append({
            "source_file": filename,
            "product_name": product_name,
            "물질명": kr_name,
            "CAS번호": comp["cas"],
            "함량(%)": comp["content_pct"],
            "물질상태": physchem["state"],
            "비중": physchem["specific_gravity"],
            "인화점(℃)": physchem["flash_point_c"],
            "폭발하한": physchem["explosion_lower"],
            "폭발상한": physchem["explosion_upper"],
            "증기압(mmHg)": physchem["vapor_pressure_mmhg"],
            "부식성": corrosive,
            "고유번호_유해화학물질": regulation["hazard_chem_no"],
            "고유번호_사고대비물질": regulation["accident_prep_no"],
            "규제구분요약": regulation["regulation_summary"],
            "독성구분_항목": "\n".join([h[0] + (f"({h[1]})" if h[1] else "") for h in hazards]),
            "독성구분_등급": "\n".join([h[2] for h in hazards]),
            "위험노출수준": endpoint_conc,
            "허용농도값": toxicity_ref,
            "msds_reg_no": meta["msds_reg_no"],
            "revision_no": meta["revision_no"],
            "revision_date": meta["revision_date"],
            "enact_date": enact_date,
            "usage": usage,
            "다성분혼합물": is_mixture,
            "needs_review": name_needs_review or physchem["specific_gravity"] is None or is_mixture,
        })

    if not rows:
        rows.append({
            "source_file": filename, "product_name": product_name,
            "물질명": product_name, "CAS번호": "", "함량(%)": "",
            "물질상태": physchem["state"], "비중": physchem["specific_gravity"],
            "인화점(℃)": physchem["flash_point_c"], "폭발하한": physchem["explosion_lower"],
            "폭발상한": physchem["explosion_upper"], "증기압(mmHg)": physchem["vapor_pressure_mmhg"],
            "부식성": corrosive, "고유번호_유해화학물질": "", "고유번호_사고대비물질": "",
            "규제구분요약": "성분 추출 실패 - 수동 입력 필요",
            "독성구분_항목": "\n".join([h[0] for h in hazards]),
            "독성구분_등급": "\n".join([h[2] for h in hazards]),
            "위험노출수준": endpoint_conc,
            "허용농도값": toxicity_ref,
            "msds_reg_no": meta["msds_reg_no"],
            "revision_no": meta["revision_no"],
            "revision_date": meta["revision_date"],
            "enact_date": enact_date,
            "usage": usage,
            "다성분혼합물": False,
            "needs_review": True,
        })
    return rows
