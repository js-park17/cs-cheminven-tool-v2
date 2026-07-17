# -*- coding: utf-8 -*-
"""6호서식(공식 양식) 및 화학물질인벤토리(내부관리용/K-REACH연동) xlsx 작성"""
import io
import math
import xlsxwriter


def _fmt(wb, **kwargs):
    return wb.add_format(kwargs)


def _is_finite_number(val):
    """xlsxwriter write_number는 NaN/Inf를 기본적으로 거부하므로 사전 검증한다."""
    if isinstance(val, bool):
        return False
    if not isinstance(val, (int, float)):
        return False
    if isinstance(val, float) and (math.isnan(val) or math.isinf(val)):
        return False
    return True


def build_6ho_form(rows):
    """공식 별지 제6호서식과 동일한 컬럼 구조로 xlsx 생성. rows: list[dict] (msds_parser 출력 + '연번' 필요)."""
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet('유해화학물질')

    header_fmt = _fmt(wb, bold=True, align='center', valign='vcenter', bg_color='#D9D9D9',
                       border=1, text_wrap=True)
    cell_fmt = _fmt(wb, border=1, align='left', valign='vcenter', text_wrap=True)
    review_fmt = _fmt(wb, border=1, align='left', valign='vcenter', text_wrap=True, bg_color='#FFF2CC')
    num_fmt = _fmt(wb, border=1, align='center', valign='vcenter', num_format='0.####')

    headers_top = [
        ('연번', 0, 0), ('유해화학물질명', 1, 1), ('CAS\n번호', 2, 2), ('고유번호', 3, 3),
        ('물질\n상태', 4, 4), ('농도\n(%)', 5, 5), ('비중', 6, 6),
    ]
    for txt, c1, c2 in headers_top:
        ws.merge_range(0, c1, 1, c2, txt, header_fmt)
    ws.merge_range(0, 7, 0, 8, '폭발한계', header_fmt)
    ws.write(1, 7, '하한', header_fmt)
    ws.write(1, 8, '상한', header_fmt)
    ws.merge_range(0, 9, 0, 10, '독성구분', header_fmt)
    ws.write(1, 9, '항목', header_fmt)
    ws.write(1, 10, '구분', header_fmt)
    for txt, col in [('위험노출수준', 11), ('허용\n농도값', 12), ('증기압\n(20℃,\nmmHg)', 13),
                      ('부식성\n(유,무)', 14)]:
        ws.merge_range(0, col, 1, col, txt, header_fmt)
    ws.merge_range(0, 15, 1, 15, '비고', header_fmt)

    col_order = [
        '연번', '물질명', 'CAS번호', '고유번호', '물질상태', '함량(%)', '비중',
        '폭발하한', '폭발상한', '독성구분_항목', '독성구분_등급', '위험노출수준', '허용농도값',
        '증기압(mmHg)', '부식성', '비고',
    ]
    r = 2
    for row in rows:
        needs_review = row.get('needs_review')
        fmt = review_fmt if needs_review else cell_fmt
        for c, key in enumerate(col_order):
            val = row.get(key, '')
            if val is None:
                val = ''
            if key in ('비중', '폭발하한', '폭발상한', '증기압(mmHg)') and _is_finite_number(val):
                ws.write_number(r, c, val, num_fmt)
            else:
                if key in ('비중', '폭발하한', '폭발상한', '증기압(mmHg)') and isinstance(val, float) and math.isnan(val):
                    val = ''  # NaN은 빈 칸으로 처리
                ws.write(r, c, val, fmt)
        r += 1

    ws.set_column('B:B', 20)
    ws.set_column('D:D', 20)
    ws.set_column('J:J', 24)
    ws.set_column('P:P', 22)
    ws.freeze_panes(2, 0)

    if any(row.get('needs_review') for row in rows):
        note_fmt = _fmt(wb, italic=True, font_color='#7F6000')
        ws.write(r + 1, 0, '※ 노란색 음영 행은 자동 추출값이 불확실하거나(다성분 혼합물 등) 확인이 필요합니다. 검토 후 수정해 주세요.', note_fmt)

    wb.close()
    output.seek(0)
    return output


def build_inventory(rows):
    """화학물질인벤토리(내부관리용) - MSDS 전체 추출정보를 원본파일 단위로 정리."""
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet('화학물질인벤토리')

    header_fmt = _fmt(wb, bold=True, align='center', valign='vcenter', bg_color='#D9D9D9', border=1, text_wrap=True)
    cell_fmt = _fmt(wb, border=1, align='left', valign='vcenter', text_wrap=True)
    review_fmt = _fmt(wb, border=1, align='left', valign='vcenter', text_wrap=True, bg_color='#FFF2CC')

    headers = [
        ('연번', 6), ('출처 MSDS 파일', 26), ('제품명(MSDS)', 20), ('구성성분(국문)', 18),
        ('CAS번호', 14), ('함량(%)', 12), ('물질상태', 8), ('비중', 8), ('인화점(℃)', 12),
        ('폭발하한(%)', 10), ('폭발상한(%)', 10), ('증기압(mmHg)', 12), ('부식성', 8),
        ('다성분혼합물', 10), ('화관법 규제구분', 32), ('유해화학물질 지정번호', 26),
        ('사고대비물질 번호', 20), ('GHS 독성구분(항목)', 30), ('GHS 독성구분(등급)', 14),
        ('노출기준(TWA)', 14), ('응급노출기준(ERPG/PAC)', 20), ('검토필요', 8),
    ]
    for c, (title, width) in enumerate(headers):
        ws.write(0, c, title, header_fmt)
        ws.set_column(c, c, width)

    keys = [
        None, 'source_file', 'product_name', '물질명', 'CAS번호', '함량(%)', '물질상태', '비중',
        '인화점(℃)', '폭발하한', '폭발상한', '증기압(mmHg)', '부식성', '다성분혼합물',
        '규제구분요약', '고유번호_유해화학물질', '고유번호_사고대비물질', '독성구분_항목',
        '독성구분_등급', '위험노출수준', '허용농도값', 'needs_review',
    ]
    r = 1
    for i, row in enumerate(rows, start=1):
        fmt = review_fmt if row.get('needs_review') else cell_fmt
        for c, key in enumerate(keys):
            if key is None:
                ws.write(r, c, i, fmt)
                continue
            val = row.get(key, '')
            if val is None:
                val = ''
            if isinstance(val, bool):
                val = '예' if val else '아니오'
            if isinstance(val, float) and math.isnan(val):
                val = ''
            ws.write(r, c, val, fmt)
        r += 1

    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, r - 1, len(headers) - 1)
    wb.close()
    output.seek(0)
    return output


def build_kreach_inventory(rows):
    """SK인천석유화학 화학물질 인벤토리 양식과 동일한 컬럼 구조로 xlsx 생성.
    rows: list[dict] - msds_parser 출력 + kreach_lookup 결과가 병합된 딕셔너리."""
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet('화학물질인벤토리')

    header_fmt = _fmt(wb, bold=True, align='center', valign='vcenter', bg_color='#D9D9D9',
                       border=1, text_wrap=True)
    title_fmt = _fmt(wb, bold=True, align='center', valign='vcenter', font_size=14)
    cell_fmt = _fmt(wb, border=1, align='left', valign='vcenter', text_wrap=True)
    review_fmt = _fmt(wb, border=1, align='left', valign='vcenter', text_wrap=True, bg_color='#FFF2CC')

    ws.merge_range(0, 0, 0, 35, '화학물질 인벤토리 (K-REACH 연동)', title_fmt)

    top_single = [
        ('순번', 0), ('제품명', 1), ('물질명', 2), ('Cas No.', 3), ('KE-NO.', 4),
        ('MSDS 작성 함량(%)', 5), ('최소 함량(%)', 6), ('최고 함량(%)', 7),
    ]
    for txt, col in top_single:
        ws.merge_range(1, col, 4, col, txt, header_fmt)

    ws.merge_range(1, 8, 2, 32, '화평법 · 화관법', header_fmt)
    sub_headers = [
        (8, 'MSDS등록번호'), (9, '기존화학물질'),
        (10, '인체급성\n고시번호'), (11, '인체급성'), (12, '인체급성\n함량기준(%)'),
        (13, '인체만성\n고시번호'), (14, '인체만성'), (15, '인체만성 함량기준(%)'),
        (16, '생태'), (17, '생태독성'), (18, '생태 함량기준(%)'),
        (19, '사고대비물질 고시번호'), (20, '사고대비물질 대상'), (21, '사고대비물질 함량기준(%)'),
        (22, '취급제한물질'), (23, '취급제한물질 고시번호'), (24, '취급제한물질 함량기준(%)'),
        (25, '취급금지물질'), (26, '취급금지물질 고시번호'), (27, '취급금지물질 함량기준(%)'),
        (28, '버전'), (29, '개정일자'), (30, '제정일자'), (31, '비고'), (32, '제품분류'),
    ]
    for col, txt in sub_headers:
        ws.write(3, col, txt, header_fmt)
        ws.write(4, col, '', header_fmt)

    for txt, col in [('유해화학물질 대상', 33), ('유해화학물질 구분', 34), ('K-REACH 조회 비고', 35)]:
        ws.merge_range(1, col, 4, col, txt, header_fmt)

    col_keys = [
        'seq', 'product_name', '물질명', 'CAS번호', 'ke_no', '함량(%)', '최소함량', '최고함량',
        'msds_reg_no', 'ke_existing',
        'acute_no', 'acute_flag', 'acute_pct',
        'chronic_no', 'chronic_flag', 'chronic_pct',
        'eco_no', 'eco_flag', 'eco_pct',
        'accident_prep_no', 'accident_prep_flag', 'accident_prep_pct',
        'restricted_flag', 'restricted_no', 'restricted_pct',
        'prohibited_flag', 'prohibited_no', 'prohibited_pct',
        'version', 'revision_date', 'enact_no', 'note', 'product_category',
        'hazard_target', 'hazard_type', 'kreach_note',
    ]
    r = 5
    for row in rows:
        needs_review = row.get('kreach_status') not in ('ok',)
        fmt = review_fmt if needs_review else cell_fmt
        for c, key in enumerate(col_keys):
            val = row.get(key, '')
            if val is None:
                val = ''
            if isinstance(val, float) and math.isnan(val):
                val = ''
            ws.write(r, c, val, fmt)
        r += 1

    ws.set_column('B:C', 20)
    ws.set_column('AJ:AJ', 30)
    ws.freeze_panes(5, 0)

    note_fmt = _fmt(wb, italic=True, font_color='#7F6000')
    ws.write(r + 1, 0,
             '※ 함량기준(%) 열은 K-REACH 목록화면에서 제공되지 않아 MSDS 원문에서 보조 추출한 값이며, '
             '정확한 규제 함량기준은 K-REACH "정보보기" 상세화면에서 재확인이 필요합니다.', note_fmt)
    ws.write(r + 2, 0,
             '※ 노란색 음영 행은 K-REACH 자동조회에 실패했거나 검색결과가 없는 경우입니다(신규화학물질 가능성 포함).', note_fmt)

    wb.close()
    output.seek(0)
    return output
