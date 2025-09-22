# app.py
# -*- coding: utf-8 -*-
import io
import re
from collections import defaultdict
import html

import streamlit as st
import pandas as pd

from compare_weekly_reports import (
    load_to_dataframe,
    build_report,
    write_excel,
)

st.set_page_config(page_title="주간 보고서 비교 (PDF/Word/Excel)", layout="wide")
st.title("📊 주간 보고서 비교 (PDF/Word/Excel 지원)")

tab_compare, tab_convert = st.tabs(["✅ 주간 비교", "📄 PDF/Word → 🧾 Excel 변환"])


def diff_markup_to_html(s: str) -> str:
    if s is None:
        return ""
    s = html.escape(str(s))
    s = s.replace("\n", "<br>")
    s = re.sub(r'\[-(.*?)-\]', r'<span style="color:#0000ff;text-decoration:line-through;">\1</span>', s)
    s = re.sub(r'\[\+(.*?)\+\]', r'<span style="color:#ff0000;font-weight:600;">\1</span>', s)
    return s


# ====================== ① 주간 비교 ======================
with tab_compare:
    st.subheader("주간 보고서 비교")

    col1, col2 = st.columns(2)
    with col1:
        prev_file = st.file_uploader("전주 파일 업로드 (PDF/Word/Excel)",
                                     type=["pdf", "docx", "xlsx", "xls"], key="prev")
    with col2:
        curr_file = st.file_uploader("금주 파일 업로드 (PDF/Word/Excel)",
                                     type=["pdf", "docx", "xlsx", "xls"], key="curr")

    project_col = "프로젝트명"; launch_col = "런칭"; work_col = "금주 진행 업무"

    if st.button("비교 실행", type="primary", use_container_width=True):
        if not prev_file or not curr_file:
            st.error("전주/금주 파일을 모두 업로드하세요."); st.stop()

        prev_df = load_to_dataframe(prev_file)
        curr_df = load_to_dataframe(curr_file)
        st.success("비교 완료 ✅")

        merged, modified, added, removed = build_report(
            prev_df, curr_df, project_col, launch_col, work_col
        )

        # Summary (줄바꿈 유지)
        st.markdown("### 요약 결과 (Summary)")
        summary_cols = [project_col, f"{launch_col}_prev", f"{launch_col}_curr",
                        f"{work_col}_prev", f"{work_col}_curr", "STATUS"]
        summary_df = merged[[c for c in summary_cols if c in merged.columns]].copy()
        for col in [f"{work_col}_prev", f"{work_col}_curr"]:
            if col in summary_df.columns:
                summary_df[col] = summary_df[col].astype(str).str.replace("\n", "<br>", regex=False)
        st.markdown(summary_df.to_html(escape=False, index=False), unsafe_allow_html=True)

        # Modified (줄바꿈 + 하이라이트)
        st.markdown("### 변경된 항목 (Modified)")
        if len(modified):
            mod_view = modified.copy()
            for col in [f"{work_col}_prev", f"{work_col}_curr"]:
                if col in mod_view.columns:
                    mod_view[col] = mod_view[col].astype(str).map(lambda x: html.escape(x).replace("\n", "<br>"))
            if "업무_diff" in mod_view.columns:
                mod_view["업무_diff"] = mod_view["업무_diff"].map(diff_markup_to_html)
            st.markdown(mod_view.to_html(escape=False, index=False), unsafe_allow_html=True)
        else:
            st.info("변경된 항목 없음")

        # Added / Removed
        st.markdown("### 추가/삭제 항목 (Added / Removed)")
        ar = []
        if len(added):   a = added.copy();   a["변경유형"] = "ADDED";   ar.append(a)
        if len(removed): r = removed.copy(); r["변경유형"] = "REMOVED"; ar.append(r)
        added_removed = pd.concat(ar, ignore_index=True) if ar else pd.DataFrame(
            columns=[project_col, launch_col, work_col, "변경유형"]
        )
        st.dataframe(added_removed, use_container_width=True)

        # 화면 느낌 그대로 엑셀 저장 (부분서식/줄바꿈)
        out_path = "weekly_diff_report.xlsx"
        write_excel(out_path, merged, modified, added, removed,
                    project_col=project_col, launch_col=launch_col, work_col=work_col)
        with open(out_path, "rb") as f:
            st.download_button("📥 결과 엑셀 다운로드", f,
                file_name="주간비교결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ====================== ② 파일 변환(줄바꿈 보존) ======================
with tab_convert:
    st.subheader("PDF/Word의 표를 추출하여 엑셀로 저장 (줄바꿈 보존)")

    src_file = st.file_uploader("PDF/Word 파일 업로드", type=["pdf", "docx"], key="pdfdoc2xl")

    # 변환용 로컬 유틸 (줄바꿈 보존)
    def _strip_keep_nl_local(s):
        if pd.isna(s): return ""
        txt = str(s).replace("\r", "\n")
        return "\n".join([ln.strip() for ln in txt.split("\n")])

    def _make_unique_columns(cols):
        cnt = defaultdict(int); out = []
        for c in cols:
            key = ("" if c is None else str(c)).replace("\r", " ").replace("\n", " ").strip()
            cnt[key] += 1
            out.append(key if cnt[key] == 1 else f"{key}_{cnt[key]}")
        return out

    def _norm_key(s: str) -> str:
        return re.sub(r"\s+", "", s or "")

    def normalize_columns_local(df: pd.DataFrame) -> pd.DataFrame:
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
                    if not tbl or len(tbl) < 2: continue
                    header = _make_unique_columns(tbl[0])
                    rows = [[_strip_keep_nl_local(x) for x in r] for r in tbl[1:]]
                    df = pd.DataFrame(rows, columns=header)
                    df = normalize_columns_local(df)
                    frames.append(df)
        return pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()

    def read_docx_to_dataframe(file_like) -> pd.DataFrame:
        from docx import Document
        frames = []
        doc = Document(file_like)
        for table in doc.tables:
            rows = []
            for row in table.rows:
                rows.append([_strip_keep_nl_local(cell.text) for cell in row.cells])
            if len(rows) > 1:
                header = _make_unique_columns(rows[0])
                data = rows[1:]
                df = pd.DataFrame(data, columns=header)
                df = normalize_columns_local(df)
                frames.append(df)
        return pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()

    if src_file is not None:
        try:
            name = src_file.name.lower()
            if name.endswith(".pdf"):
                df_conv = read_pdf_to_dataframe(src_file)
            elif name.endswith(".docx"):
                df_conv = read_docx_to_dataframe(src_file)
            else:
                st.error("⚠️ .doc(구버전)은 지원하지 않습니다. .docx 또는 PDF로 업로드하세요."); st.stop()
        except Exception as e:
            st.exception(e); st.stop()

        if df_conv.empty:
            st.warning("표를 찾지 못했습니다.")
        else:
            st.success(f"표 추출 완료! (행 {len(df_conv)})")
            st.dataframe(df_conv, use_container_width=True)

            # 엑셀로 저장(랩 적용, xlsxwriter 사용)
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="xlsxwriter") as wr:
                df_conv.to_excel(wr, sheet_name="Extracted", index=False)
                ws = wr.sheets["Extracted"]
                wrap = wr.book.add_format({"text_wrap": True, "valign": "top"})
                ws.set_column(0, len(df_conv.columns)-1, 45, wrap)
            buf.seek(0)
            st.download_button("📥 엑셀로 다운로드",
                               data=buf,
                               file_name="doc_or_pdf_extracted.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
