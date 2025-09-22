# compare_weekly_reports.py
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def build_report(prev_df, curr_df, project_col, launch_col, work_col):
    # 테스트용 최소 버전 (실제 로직은 확장 가능)
    merged = pd.DataFrame()
    modified = pd.DataFrame()
    added = pd.DataFrame()
    removed = pd.DataFrame()
    return merged, modified, added, removed

def write_excel(out_path, merged, modified, added, removed,
                project_col, launch_col, work_col):
    # 그냥 빈 엑셀 파일 하나 만들어주는 더미
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        pd.DataFrame({"dummy":["no data"]}).to_excel(writer, sheet_name="Summary", index=False)
