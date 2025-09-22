# app.py
import streamlit as st
import pandas as pd
from compare_weekly_reports import load_to_dataframe, build_report, write_excel

st.set_page_config(page_title="주간 보고서 비교 도구", layout="wide")
st.title("📊 주간 보고서 비교 (PDF/Excel 지원)")

with st.sidebar:
    st.markdown("### 업로드")
    prev_file = st.file_uploader("전주 파일 (PDF/XLSX)", type=["pdf", "xlsx", "xls"], key="prev")
    curr_file = st.file_uploader("금주 파일 (PDF/XLSX)", type=["pdf", "xlsx", "xls"], key="curr")
    project_col = st.text_input("프로젝트명 컬럼", "프로젝트명")
    launch_col = st.text_input("런칭 컬럼", "런칭")
    work_col = st.text_input("금주 진행 업무 컬럼", "금주 진행 업무")

run = st.button("비교 실행")

if run:
    if not prev_file or not curr_file:
        st.error("전주/금주 파일을 모두 업로드하세요.")
        st.stop()

    prev_df = load_to_dataframe(prev_file)
    curr_df = load_to_dataframe(curr_file)

    merged, modified, added, removed = build_report(
        prev_df, curr_df,
        project_col=project_col, launch_col=launch_col, work_col=work_col
    )

    st.success("비교 완료 ✅")

    # 📌 화면에서 결과 미리보기
    st.subheader("요약 결과 (Summary)")
    st.dataframe(merged, use_container_width=True)

    st.subheader("변경된 항목 (Modified)")
    if len(modified):
        st.dataframe(modified, use_container_width=True)
    else:
        st.info("변경된 항목 없음")

    st.subheader("추가/삭제 항목 (Added / Removed)")
    ar = []
    if len(added):
        a = added.copy(); a["변경유형"] = "ADDED"; ar.append(a)
    if len(removed):
        r = removed.copy(); r["변경유형"] = "REMOVED"; ar.append(r)
    added_removed = pd.concat(ar, ignore_index=True) if ar else pd.DataFrame(columns=[project_col, launch_col, work_col, "변경유형"])
    st.dataframe(added_removed, use_container_width=True)

    # 📌 다운로드 버튼
    out_path = "weekly_diff_report.xlsx"
    write_excel(out_path, merged, modified, added, removed,
                project_col=project_col, launch_col=launch_col, work_col=work_col)

    with open(out_path, "rb") as f:
        st.download_button("📥 결과 엑셀 다운로드", f, file_name="주간비교결과.xlsx")
