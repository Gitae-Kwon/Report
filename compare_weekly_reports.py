# compare_weekly_reports.py
# -*- coding: utf-8 -*-
import pandas as pd
import difflib
import re
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


# ---------------------------
# 내부 유틸 함수
# ---------------------------
def _make_unique_columns(cols):
    """컬럼명이 중복되면 name, name_2, name_3 식으로 바꿔줌"""
    counts = defaultdict(int)
    out = []
    for c in cols:
        c = "" if c is None else str(c)
        counts[c] += 1
        if counts[c] == 1:
            out.append(c)
        else:
            out.append(f"{c}_{counts[c]}")
    return out


def _read_pdf_df(file_like) -> pd.DataFrame:
    """
    pdfplumber로 PDF 표를 읽어서
    ['프로젝트명','런칭','금주 진행 업무'] 3개 컬럼만 표준화해 합칩니다.
    """
    import pdfplumber

    normalized_frames = []

    def _norm_key(s: str) -> str:
        return re.sub(r"\s+", "", s or "")

    with pdfplumber.open(file_like) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables() or []
            for tbl in tables:
                if not tbl or len(tbl) < 2:
                    continue
                header_raw = tbl[0]
                rows = tbl[1:]

                # 1) 헤더 유니크 처리 + 개행/공백 제거
                header = [ ("" if h is None else str(h).replace("\n", " ").strip()) for h in header_raw ]
                header = _make_unique_columns(header)

                df = pd.DataFrame(rows, columns=header)

                # 2) 헤더 매핑
                col_map = {}
                for col in df.columns:
                    key = _norm_key(col)
                    if ("프로젝트" in key and "명" in key) or key in ("프로젝트", "프로젝트명"):
                        col_map[col] = "프로젝트명"
                    elif ("런칭" in key) or ("오픈" in key) or ("오픈일" in key):
                        col_map[col] = "런칭"
                    elif ("금주" in key and ("업무" in key or "진행" in key)) or key in ("금주진행업무","금주업무"):
                        col_map[col] = "금주 진행 업무"

                df = df.rename(columns=col_map)

                # 3) 필요한 컬럼만 뽑기
                std = pd.DataFrame({
                    "프로젝트명": df["프로젝트명"] if "프로젝트명" in df.columns else pd.Series(dtype=str),
                    "런칭": df["런칭"] if "런칭" in df.columns else pd.Series(dtype=str),
                    "금주 진행 업무": df["금주 진행 업무"] if "금주 진행 업무" in df.columns else pd.Series(dtype=str),
                })

                # 완전히 빈 행 제거
                std = std.replace({None: "", pd.NA: "", "nan": ""})
                std = std[~(std["프로젝트명"].astype(str).str.strip().eq("") &
                            std["런칭"].astype(str).str.strip().eq("") &
                            std["금주 진행 업무"].astype(str).str.strip().eq(""))]

                if len(std):
                    normalized_frames.append(std)

    if not normalized_frames:
        return pd.DataFrame(columns=["프로젝트명", "런칭", "금주 진행 업무"])

    out = pd.concat(normalized_frames, ignore_index=True)
    return out


# ---------------------------
# 외부에서 쓰는 주요 함수
# ---------------------------
def load_to_dataframe(file_or_path, sheet=None):
    """엑셀(.xlsx/.xls) 또는 PDF(.pdf)을 DataFrame으로 변환"""
    name = getattr(file_or_path, "name", str(file_or_path)).lower()
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(file_or_path, sheet_name=sheet)
    elif name.endswith(".pdf"):
        return _read_pdf_df(file_or_path)
    else:
        raise ValueError("지원하지 않는 파일 형식입니다. (.xlsx, .xls, .pdf)")


def make_inline_diff(a: str, b: str) -> str:
    """텍스트 차이를 보기 쉽게 표시"""
    if pd.isna(a) and pd.isna(b):
        return ""
    a = "" if pd.isna(a) else str(a)
    b = "" if pd.isna(b) else str(b)

    sm = difflib.SequenceMatcher(a=a, b=b)
    pieces = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            pieces.append(b[j1:j2])
        elif tag == "insert":
            pieces.append(f"[+{b[j1:j2]}+]")
        elif tag == "delete":
            pieces.append(f"[-{a[i1:i2]}-]")
        elif tag == "replace":
            pieces.append(f"[-{a[i1:i2]}-][+{b[j1:j2]}+]")
    return "".join(pieces)


def build_report(prev_df, curr_df, project_col, launch_col, work_col):
    """전주 vs 금주 비교"""
    prev_df = prev_df.copy()
    curr_df = curr_df.copy()

    merged = curr_df.merge(prev_df, on=project_col, how="outer",
                           suffixes=("_curr", "_prev"), indicator=True)

    def status_row(r):
        if r["_merge"] == "left_only":
            return "ADDED"
        if r["_merge"] == "right_only":
            return "REMOVED"
        if r.get(f"{launch_col}_curr") != r.get(f"{launch_col}_prev") or \
           r.get(f"{work_col}_curr") != r.get(f"{work_col}_prev"):
            return "MODIFIED"
        return "UNCHANGED"

    merged["STATUS"] = merged.apply(status_row, axis=1)

    modified = merged[merged["STATUS"]=="MODIFIED"][
        [project_col, f"{launch_col}_prev", f"{launch_col}_curr",
         f"{work_col}_prev", f"{work_col}_curr"]
    ].copy()
    modified["업무_diff"] = [
        make_inline_diff(a, b)
        for a, b in zip(modified[f"{work_col}_prev"], modified[f"{work_col}_curr"])
    ]

    added = merged[merged["STATUS"]=="ADDED"][
        [project_col, f"{launch_col}_curr", f"{work_col}_curr"]
    ].copy()

    removed = merged[merged["STATUS"]=="REMOVED"][
        [project_col, f"{launch_col}_prev", f"{work_col}_prev"]
    ].copy()

    return merged, modified, added, removed


def write_excel(out_path, merged, modified, added, removed,
                project_col, launch_col, work_col):
    """결과를 엑셀로 저장"""
    with pd.ExcelWriter(out_path, engine="openpyxl") as wr:
        merged.to_excel(wr, "Summary", index=False)
        modified.to_excel(wr, "Modified", index=False)
        added.to_excel(wr, "Added", index=False)
        removed.to_excel(wr, "Removed", index=False)

    wb = load_workbook(out_path)
    if "Modified" in wb.sheetnames:
        ws = wb["Modified"]
        yellow = PatternFill(start_color="FFF8B4", end_color="FFF8B4", fill_type="solid")
        for r in range(2, ws.max_row+1):
            ws.cell(row=r, column=3).fill = yellow  # launch_curr
            ws.cell(row=r, column=5).fill = yellow  # work_curr
    wb.save(out_path)
