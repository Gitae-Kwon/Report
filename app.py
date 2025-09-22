# app.py
# -*- coding: utf-8 -*-
import io
import re
from collections import defaultdict

import streamlit as st
import pandas as pd

# 비교 로직은 모듈에서 사용
from compare_weekly_reports import (
    load_to_dataframe,
    build_report,
    write_excel,
)

st.set_page_config(page_title="주간 보고서 비교 (PDF/Excel)", layout="wide")
st.title("📊 주간 보고서 비교 (PDF/Excel 지원)")

tab_compare, tab_pdf2xl = st.tabs(["✅ 주간 비교", "📄 PDF → 🧾 Excel 변환"])

# =========================================================
# ① 주간 비교 탭
# =========================================================
with tab_compare:
    st.subheader("주간 보고서 비교")

    col1, col2 = st.columns(2)
    with col1:
        prev_file = st.file_uploader("전주 파일 업로드 (PDF/XLSX)", type=["pdf", "xlsx", "xls"], key="prev")
    with col2:
        curr_file = st.file_uploader("금주 파일 업로드 (PDF/XLSX)", type=["pdf", "xlsx", "xls"], key="curr")

    # 컬럼명은 정규화 결과 기준으로 고정 (오타/불일치 방지)
    project_col = "프로젝트명"
    launch_col  = "런칭"
    work_col    = "금주 진행 업무"

    if st.button("비교 실행", type="primary", use_container_width=True):
        if not prev_file or not curr_file:
            st.error("전주/금주 파일을 모두 업로드하세요.")
            st.stop()

        try:
            prev_df = load_to_dataframe(prev_file)
            curr_df = load_to_dataframe(curr_file)
        except Exception as e:
            st.exception(e)
            st.stop()

        st.success("비교 완료 ✅")

        # 컬럼 확인(디버깅용)
        with st.expander("전주/금주 컬럼 확인", expanded=False):
            st.write("전주 컬럼:", list(prev_df.columns))
            st.write("금주 컬럼:", list(curr_df.columns))

        merged, modified, added, removed = build_report(
            prev_df, curr_df,
            project_col=project_col, launch_col=launch_col, work_col=work_col
        )

        st.markdown("### 요약 결과 (Summary)")
        st.dataframe(merged[[c for c in merged.columns if c in
                             [project_col, f"{launch_col}_curr", f"{work_col}_curr",
                              f"{launch_col}_prev", f"{work_col}_prev", "_merge", "STATUS"]]],
                     use_container_width=True)

        st.markdown("### 변경된 항목 (Modified) ↪︎")
        if len(modified):
            st.dataframe(modified, use_container_width=True)
        else:
            st.info("변경된 항목 없음")

        st.markdown("### 추가/삭제 항목 (Added / Removed)")
        ar = []
        if len(added):
            a = added.copy(); a["변경유형"] = "ADDED"; ar.append(a)
        if len(removed):
            r = removed.copy(); r["변경유형"] = "REMOVED"; ar.append(r)
        added_removed = pd.concat(ar, ignore_index=True) if ar else pd.DataFrame(
            columns=[project_col, launch_col, work_col, "변경유형"]
        )
        st.dataframe(added_removed, use_container_width=True)

        # 결과 엑셀 생성 & 다운로드
        out_path = "weekly_diff_report.xlsx"
        write_excel(out_path, merged, modified, added, removed,
                    project_col=project_col, launch_col=launch_col, work_col=work_col)

        with open(out_path, "rb") as f:
            st.download_button("📥 결과 엑셀 다운로드",
                               f,
                               file_name="주간비교결과.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# =========================================================
# ② PDF → Excel 변환 탭
# =========================================================
with tab_pdf2xl:
    st.subheader("PDF의 표를 추출하여 엑셀로 저장")

    pdf_file = st.file_uploader("PDF 업로드", type=["pdf"], key="pdf2xl")

    def _strip(s):
        if pd.isna(s):
            return ""
        return str(s).replace("\n", " ").strip()

    def _make_unique_columns(cols):
        cnt = defaultdict(int)
        out = []
        for c in cols:
            c = "" if c is None else str(c)
            cnt[c] += 1
            out.append(c if cnt[c] == 1 else f"{c}_{cnt[c]}")
        return out

    def _norm_key(s: str) -> str:
        return re.sub(r"\s+", "", s or "")

    def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
        """(선택) 보고서 스타일에 맞춰 일부 헤더 표준화"""
        rename_map = {}
        for col in df.columns:
            key = _norm_key(col)
            if ("프로젝트" in key and "명" in key) or key in ("프로젝트명", "프로젝트"):
                rename_map[col] = "프로젝트명"
            elif ("런칭" in key) or ("오픈" in key):
                rename_map[col] = "런칭"
            elif ("금주" in key) and (("업무" in key) or ("진행" in key)):
                rename_map[col] = "금주 진행 업무"
        return df.rename(columns=rename_map)

    def read_pdf_to_dataframe(file_like) -> pd.DataFrame:
        import pdfplumber
        frames = []
        with pdfplumber.open(file_like) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables() or []
                for tbl in tables:
                    if not tbl or len(tbl) < 2:
                        continue
                    header = [_strip(h) for h in tbl[0]]
                    header = _make_unique_columns(header)
                    rows = [[_strip(x) for x in r] for r in tbl[1:]]
                    df = pd.DataFrame(rows, columns=header)
                    df = normalize_columns(df)  # 필요 없으면 주석 처리
                    df = df.dropna(how="all")
                    if len(df):
                        frames.append(df)
        if not frames:
            return pd.DataFrame()
        # 표 구조가 달라도 합치도록 outer concat
        return pd.concat(frames, ignore_index=True, sort=False)

    if pdf_file is not None:
        try:
            df_pdf = read_pdf_to_dataframe(pdf_file)
        except Exception as e:
            st.exception(e)
            st.stop()

        if df_pdf.empty:
            st.warning("표를 찾지 못했습니다. PDF 레이아웃을 확인해주세요.")
        else:
            st.success(f"표 추출 완료! (행 {len(df_pdf)})")
            st.dataframe(df_pdf, use_container_width=True)

            # 엑셀로 다운로드
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                df_pdf.to_excel(writer, sheet_name="Extracted", index=False)
            buf.seek(0)
            st.download_button("📥 엑셀로 다운로드",
                               data=buf,
                               file_name="pdf_extracted.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
