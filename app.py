# -*- coding: utf-8 -*-
"""
6호서식(유해화학물질 목록) 작성 툴
- MSDS(PDF) 여러 개를 업로드하면 CAS/물질명/비중/인화점/폭발한계/증기압 등 물성값과
  K-REACH(화학물질정보처리시스템) 공개검색을 통한 화관법 규제정보를 자동 반영하여
  1) 별지 제6호서식(유해화학물질만 발췌), 2) 제품명 기준 화학물질인벤토리(K-REACH 연동)
  두 가지 결과물을 내려받을 수 있다.
"""
import re

import streamlit as st
import pandas as pd

import msds_parser as mp
import xlsx_writers as xw
import kreach_lookup as kr


def dedup_ho6_by_cas(rows):
    """6호서식 행을 CAS번호 기준으로 병합한다.
    - 같은 CAS가 여러 MSDS에서 추출되어도 6호서식에는 1행만 남긴다.
    - 함량(%)은 관측된 최소~최대 범위로 병합한다.
    - 비고에는 근거가 된 모든 제품명을 나열하고, 위험노출수준이 기술지침 표(1차 공식 출처)가
      아닌 보충조사/미확정 값인 경우 그 출처를 함께 기입한다.
    - CAS가 없는 행(추출 실패 등)은 병합하지 않고 그대로 각각 유지한다.
    """
    groups = {}
    order = []
    for r in rows:
        cas = (r.get("CAS번호") or "").strip()
        key = cas if cas else f"__nocasid_{id(r)}"
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(r)

    merged = []
    for key in order:
        group = groups[key]
        base = dict(group[0])
        product_names = [g.get("product_name", "") for g in group if g.get("product_name")]
        product_names = list(dict.fromkeys(product_names))  # 순서 보존 중복제거

        if len(group) > 1:
            pct_nums = []
            for g in group:
                txt = str(g.get("함량(%)", "") or "")
                pct_nums += [float(n) for n in re.findall(r"\d+(?:\.\d+)?", txt)]
            if pct_nums:
                lo, hi = min(pct_nums), max(pct_nums)
                base["함량(%)"] = f"{lo:g}%" if lo == hi else f"{lo:g}~{hi:g}%"

            sg_vals = {g.get("비중") for g in group
                       if isinstance(g.get("비중"), (int, float)) and str(g.get("비중")) != "nan"}
            source_files = list(dict.fromkeys(g.get("source_file", "") for g in group if g.get("source_file")))
            base["source_file"] = ", ".join(source_files)
            base["needs_review"] = bool(base.get("needs_review")) or any(g.get("needs_review") for g in group) \
                or len(sg_vals) > 1

        remark_parts = [", ".join(product_names)] if product_names else []
        ep_status = base.get("위험노출수준_상태", "")
        ep_source = base.get("위험노출수준_출처", "")
        if ep_status in ("supplementary", "na", "no_data", "msds_text") and ep_source:
            remark_parts.append(f"[위험노출수준 출처: {ep_source}]")
        base["비고"] = " / ".join(remark_parts)

        merged.append(base)

    for i, r in enumerate(merged, start=1):
        r["연번"] = i
    return merged


st.set_page_config(page_title="6호서식 작성 툴", layout="wide")
st.title("🧪 유해화학물질 목록(별지 제6호서식) 자동 작성 툴")
st.markdown("##### MSDS(PDF) 업로드 → 물성값 자동추출 + K-REACH 규제정보 자동조회")

with st.sidebar:
    st.header("ℹ️ 사용 안내")
    st.markdown(
        "1. 각 물질의 MSDS PDF 파일을 모두 선택해 업로드하세요.\n"
        "2. **[MSDS 분석 + K-REACH 조회 시작]** 버튼을 누르면 CAS번호로 K-REACH를 자동 조회합니다.\n"
        "3. 결과 표를 검토하고, 필요하면 '6호서식 대상' 여부를 직접 수정하세요.\n"
        "4. 완료 후 별지6호서식과 K-REACH 연동 인벤토리를 각각 내려받으세요."
    )
    st.divider()
    st.caption(
        "※ 6호서식은 K-REACH 조회 결과 인체급성·인체만성·생태·사고대비물질 중 하나라도 "
        "해당되는 '유해화학물질'만 자동으로 골라 작성됩니다. K-REACH 조회에 실패한 물질은 "
        "안전을 위해 6호서식에 포함하고 노란색으로 표시하니 직접 확인해 주세요.\n\n"
        "※ 다성분 혼합물(제품) MSDS는 성분별로 각각 CAS 조회를 수행합니다."
    )

uploaded_files = st.file_uploader(
    "MSDS PDF 파일 업로드 (여러 개 선택 가능)", type=["pdf"], accept_multiple_files=True
)

if uploaded_files:
    run = st.button("🔍 MSDS 분석 + K-REACH 조회 시작")

    if run:
        all_rows = []
        errors = []
        with st.spinner("MSDS 분석 중..."):
            for uf in uploaded_files:
                try:
                    rows = mp.parse_msds(uf.read(), uf.name)
                    all_rows.extend(rows)
                except Exception as e:
                    errors.append((uf.name, str(e)))
        for fname, err in errors:
            st.error(f"'{fname}' 처리 중 오류: {err}")

        if all_rows:
            cas_list = [r.get("CAS번호", "") for r in all_rows if r.get("CAS번호")]
            progress = st.progress(0.0, text="K-REACH 조회 준비 중...")

            def _cb(i, total, cas):
                progress.progress(i / total, text=f"K-REACH 조회 중... ({i}/{total}) {cas}")

            kreach_results = kr.search_many(cas_list, progress_cb=_cb) if cas_list else {}
            progress.empty()

            for i, row in enumerate(all_rows, start=1):
                row["연번"] = i
                cas = row.get("CAS번호", "")
                kres = kreach_results.get(cas, {})

                parts = []
                for key in ("acute_no", "chronic_no", "eco_no"):
                    v = kres.get(key)
                    if v and v not in parts:
                        parts.append(v)
                if kres.get("accident_prep_no"):
                    parts.append(f"사고대비 {kres['accident_prep_no']}")
                row["고유번호"] = "\n".join(parts)

                # 물질명 최종 확정: K-REACH 국문명(name_kr)이 있으면 최우선 사용,
                # 없으면 msds_parser가 CAS 기준으로 해석해 둔 값을 그대로 유지한다.
                # (제품명은 어떤 단계에서도 물질명으로 사용하지 않는다.)
                if kres.get("name_kr"):
                    row["물질명"] = kres["name_kr"]
                if not row.get("물질명"):
                    row["물질명"] = f"[물질명 미확인] CAS {cas}" if cas else "[물질명 미확인]"
                    row["needs_review"] = True

                row["비고"] = row.get("product_name", "")

                is_regulated = bool(
                    kres.get("acute_no") or kres.get("chronic_no") or kres.get("eco_no")
                    or kres.get("accident_prep_no") or kres.get("restricted_raw")
                )
                kreach_uncertain = kres.get("status") not in ("ok", "not_found")
                row["kreach_status"] = kres.get("status", "error")
                row["kreach_note"] = kres.get("detail_note", "")
                row["kreach_ke_no"] = kres.get("ke_no", "")
                row["kreach_name_kr"] = kres.get("name_kr", "")
                row["kreach_acute_no"] = kres.get("acute_no", "")
                row["kreach_chronic_no"] = kres.get("chronic_no", "")
                row["kreach_eco_no"] = kres.get("eco_no", "")
                row["kreach_accident_prep_no"] = kres.get("accident_prep_no", "")
                row["kreach_restricted_raw"] = kres.get("restricted_raw", "")
                row["6호서식대상"] = bool(is_regulated or kreach_uncertain)
                row["needs_review"] = bool(row.get("needs_review")) or kreach_uncertain

            st.session_state["all_rows"] = all_rows

    if "all_rows" in st.session_state:
        all_rows = st.session_state["all_rows"]
        df = pd.DataFrame(all_rows)

        st.write("---")
        hazard_count = int(df["6호서식대상"].sum())
        c1, c2, c3 = st.columns(3)
        c1.metric("업로드 MSDS 수", f"{len(uploaded_files)}개")
        c2.metric("추출된 물질(행) 수", f"{len(df)}개")
        c3.metric("6호서식 대상(유해화학물질)", f"{hazard_count}개")

        st.write("### 📋 추출 결과 검토 및 수정")
        st.caption("'6호서식대상' 열을 체크/해제하면 6호서식 다운로드에 포함될 물질을 직접 조정할 수 있습니다.")

        display_cols = [
            "연번", "6호서식대상", "source_file", "product_name", "물질명", "CAS번호", "고유번호",
            "물질상태", "함량(%)", "비중", "폭발하한", "폭발상한", "위험노출수준", "위험노출수준_출처",
            "허용농도값", "증기압(mmHg)", "부식성", "독성구분_항목", "독성구분_등급",
            "kreach_status", "kreach_note",
        ]
        show_df = df[display_cols].rename(columns={"source_file": "출처파일", "product_name": "제품명"})

        edited_df = st.data_editor(
            show_df,
            column_config={
                "6호서식대상": st.column_config.CheckboxColumn(),
                "비중": st.column_config.NumberColumn(format="%.4f"),
                "폭발하한": st.column_config.NumberColumn(format="%.2f"),
                "폭발상한": st.column_config.NumberColumn(format="%.2f"),
                "증기압(mmHg)": st.column_config.NumberColumn(format="%.2f"),
            },
            disabled=["연번", "출처파일", "위험노출수준_출처", "kreach_status", "kreach_note"],
            hide_index=True,
            use_container_width=True,
        )

        merged_rows = []
        for i, row in enumerate(all_rows):
            m = dict(row)
            m["6호서식대상"] = bool(edited_df.iloc[i]["6호서식대상"])
            m["물질명"] = edited_df.iloc[i]["물질명"]
            m["고유번호"] = edited_df.iloc[i]["고유번호"]
            m["비중"] = edited_df.iloc[i]["비중"]
            m["함량(%)"] = edited_df.iloc[i]["함량(%)"]
            m["needs_review"] = row.get("needs_review", False)
            merged_rows.append(m)

        st.write("---")
        st.write("### 📥 결과물 다운로드")

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            ho6_rows_raw = [r for r in merged_rows if r.get("6호서식대상")]
            ho6_rows = dedup_ho6_by_cas(ho6_rows_raw)
            if len(ho6_rows) < len(ho6_rows_raw):
                st.caption(
                    f"※ 동일 CAS번호 {len(ho6_rows_raw)}건 → {len(ho6_rows)}건으로 병합했습니다 "
                    "(같은 물질이 여러 MSDS에서 확인된 경우 1행으로 통합)."
                )
            ho6_bytes = xw.build_6ho_form(ho6_rows)
            st.download_button(
                label=f"[별지6호서식] 유해화학물질목록 다운로드 ({len(ho6_rows)}건)",
                data=ho6_bytes.getvalue(),
                file_name="1-나-1_유해화학물질목록(별지6호서식).xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        with col_dl2:
            sorted_rows = sorted(merged_rows, key=lambda r: (r.get("product_name") or ""))
            kreach_rows = []
            for i, row in enumerate(sorted_rows, start=1):
                content_pct = str(row.get("함량(%)", "") or "")
                nums = [float(n) for n in re.findall(r"\d+(?:\.\d+)?", content_pct)]
                min_pct = min(nums) if nums else ""
                max_pct = max(nums) if nums else ""

                def _threshold_from_text(text):
                    tm = re.search(r"(\d+(?:\.\d+)?)\s*%\s*이상", text or "")
                    return float(tm.group(1)) if tm else ""

                acute_pct = _threshold_from_text(row.get("고유번호_유해화학물질", "")) if row.get("kreach_acute_no") else ""
                accident_pct = _threshold_from_text(row.get("고유번호_사고대비물질", "")) if row.get("kreach_accident_prep_no") else ""

                hazard_types = []
                if row.get("kreach_acute_no"):
                    hazard_types.append("인체급성")
                if row.get("kreach_chronic_no"):
                    hazard_types.append("인체만성")
                if row.get("kreach_eco_no"):
                    hazard_types.append("생태")
                if row.get("kreach_accident_prep_no"):
                    hazard_types.append("사고대비물질")

                kreach_rows.append({
                    "seq": i,
                    "product_name": row.get("product_name", ""),
                    "물질명": row.get("kreach_name_kr") or row.get("물질명", ""),
                    "CAS번호": row.get("CAS번호", ""),
                    "ke_no": row.get("kreach_ke_no", ""),
                    "함량(%)": row.get("함량(%)", ""),
                    "최소함량": min_pct,
                    "최고함량": max_pct,
                    "msds_reg_no": row.get("msds_reg_no", ""),
                    "ke_existing": row.get("kreach_ke_no", ""),
                    "acute_no": row.get("kreach_acute_no", ""),
                    "acute_flag": "O" if row.get("kreach_acute_no") else "",
                    "acute_pct": acute_pct,
                    "chronic_no": row.get("kreach_chronic_no", ""),
                    "chronic_flag": "O" if row.get("kreach_chronic_no") else "",
                    "chronic_pct": "",
                    "eco_no": row.get("kreach_eco_no", ""),
                    "eco_flag": "O" if row.get("kreach_eco_no") else "",
                    "eco_pct": "",
                    "accident_prep_no": row.get("kreach_accident_prep_no", ""),
                    "accident_prep_flag": "O" if row.get("kreach_accident_prep_no") else "",
                    "accident_prep_pct": accident_pct,
                    "restricted_flag": "O" if row.get("kreach_restricted_raw") else "",
                    "restricted_no": row.get("kreach_restricted_raw", ""),
                    "restricted_pct": "",
                    "prohibited_flag": "",
                    "prohibited_no": "",
                    "prohibited_pct": "",
                    "version": row.get("revision_no", ""),
                    "revision_date": row.get("revision_date", ""),
                    "enact_no": row.get("enact_date", ""),
                    "note": "",
                    "product_category": row.get("usage", ""),
                    "hazard_target": "유해화학물질 대상" if hazard_types else "-",
                    "hazard_type": ", ".join(hazard_types) if hazard_types else "-",
                    "kreach_status": row.get("kreach_status", "error"),
                    "kreach_note": row.get("kreach_note", ""),
                })

            kreach_bytes = xw.build_kreach_inventory(kreach_rows)
            st.download_button(
                label=f"[화학물질인벤토리-KREACH연동] 다운로드 ({len(kreach_rows)}건, 제품명순)",
                data=kreach_bytes.getvalue(),
                file_name="화학물질인벤토리_KREACH연동.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.info("파일을 업로드한 뒤 위 버튼을 눌러 분석을 시작하세요.")
else:
    st.info("왼쪽 안내를 참고하여 MSDS PDF 파일을 업로드하면 분석을 시작할 수 있습니다.")
