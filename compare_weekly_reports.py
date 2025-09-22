# compare_weekly_reports.py
import pandas as pd
import difflib
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def load_to_dataframe(file_or_path, sheet=None):
    """엑셀(.xlsx/.xls) 또는 PDF(.pdf)을 DataFrame으로 변환"""
    name = getattr(file_or_path, "name", str(file_or_path))
    if name.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(file_or_path, sheet_name=sheet)
    elif name.lower().endswith(".pdf"):
        import pdfplumber
        frames = []
        with pdfplumber.open(file_or_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for tbl in tables or []:
                    if not tbl or len(tbl) < 2:
                        continue
                    header, rows = tbl[0], tbl[1:]
                    df = pd.DataFrame(rows, columns=header)
                    frames.append(df)
        if frames:
            df = pd.concat(frames, ignore_index=True)
        else:
            df = pd.DataFrame(columns=["프로젝트명","런칭","금주 진행 업무"])
        return df
    else:
        raise ValueError("지원하지 않는 파일 형식입니다. (.xlsx, .xls, .pdf)")

def build_report(prev_df, curr_df, project_col, launch_col, work_col):
    """전주 vs 금주 비교"""
    prev_df = prev_df.copy()
    curr_df = curr_df.copy()

    merged = curr_df.merge(prev_df, on=project_col, how="outer",
                           suffixes=("_curr", "_prev"), indicator=True)

    def status_row(row):
        if row["_merge"] == "left_only":
            return "ADDED"
        if row["_merge"] == "right_only":
            return "REMOVED"
        if row[f"{launch_col}_curr"] != row[f"{launch_col}_prev"] or \
           row[f"{work_col}_curr"] != row[f"{work_col}_prev"]:
            return "MODIFIED"
        return "UNCHANGED"

    merged["STATUS"] = merged.apply(status_row, axis=1)

    modified = merged[merged["STATUS"]=="MODIFIED"][
        [project_col, f"{launch_col}_prev", f"{launch_col}_curr",
         f"{work_col}_prev", f"{work_col}_curr"]
    ]

    added = merged[merged["STATUS"]=="ADDED"][
        [project_col, f"{launch_col}_curr", f"{work_col}_curr"]
    ]
    removed = merged[merged["STATUS"]=="REMOVED"][
        [project_col, f"{launch_col}_prev", f"{work_col}_prev"]
    ]

    return merged, modified, added, removed

def write_excel(out_path, merged, modified, added, removed,
                project_col, launch_col, work_col):
    """결과를 엑셀로 저장"""
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        merged.to_excel(writer, sheet_name="Summary", index=False)
        modified.to_excel(writer, sheet_name="Modified", index=False)
        added.to_excel(writer, sheet_name="Added", index=False)
        removed.to_excel(writer, sheet_name="Removed", index=False)

    wb = load_workbook(out_path)
    if "Modified" in wb.sheetnames:
        ws = wb["Modified"]
        yellow = PatternFill(start_color="FFF8B4", end_color="FFF8B4", fill_type="solid")
        for row in range(2, ws.max_row+1):
            ws.cell(row=row, column=3).fill = yellow
            ws.cell(row=row, column=5).fill = yellow
    wb.save(out_path)
