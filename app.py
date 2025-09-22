import streamlit as st
import pandas as pd
import pdfplumber
from compare_weekly_reports import build_report, write_excel

st.title("주간 보고서 비교 도구")

prev_file = st.file_uploader("전주 파일 업로드 (PDF/Excel)", type=["pdf", "xlsx"])
curr_file = st.file_uploader("금주 파일 업로드 (PDF/Excel)", type=["pdf", "xlsx"])

if prev_file and curr_file:
    if st.button("비교 실행"):
        # 1. 파일 읽기 → DataFrame 변환
        prev_df = load_to_dataframe(prev_file)  # pdf/xlsx 구분해서 처리
        curr_df = load_to_dataframe(curr_file)

        # 2. 비교 실행
        merged, modified, added, removed = build_report(
            prev_df, curr_df,
            project_col="프로젝트명",
            launch_col="런칭",
            work_col="금주 진행 업무"
        )

        # 3. 엑셀 파일 생성
        out_path = "weekly_diff_report.xlsx"
        write_excel(out_path, merged, modified, added, removed,
                    project_col="프로젝트명",
                    launch_col="런칭",
                    work_col="금주 진행 업무")

        with open(out_path, "rb") as f:
            st.download_button("결과 엑셀 다운로드", f, file_name="주간비교결과.xlsx")
