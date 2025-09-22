# compare_weekly_reports.py
# -*- coding: utf-8 -*-
import re
from collections import defaultdict
import difflib
from typing import Optional

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

import pdfplumber
from docx import Document


# =========================
# 공통 유틸
# =========================
def _strip(s):
    if s is None or pd.isna(s):
        return ""
    return str(s).replace("\n", " ").strip()


def _make_unique_columns(cols):
    """중복 헤더를 name, name_2, name_3... 로 유니크하게 변경"""
    counts = defaultdict(int)
    out = []
    for c in cols:
        c = "" if c is None else str(c)
        counts[c] += 1
        out.append(c if counts[c] == 1 else f"{c}_{counts[c]}")
    return out


def _norm_key(s: str) -> str:
    return re.sub(r"\s+", "", s or "")


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    다양한 표기 변형을 표준 컬럼명으로 통일:
      - 프로젝트명
      - 런칭   (런칭, 런칭(예정), 오픈, 오픈일 등)
      - 금주 진행 업무 (금주진행업무, 금주 진행업무, 금주진행 등)
    """
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


def normalize_values(df: pd.DataFrame) -> pd.DataFrame:
    """문자열 컬럼의 개행/여백을 정리"""
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].map(_strip)
    return df


def make_inline_diff(a: str, b: str) -> str:
    """업무 텍스트의 변경점을 [-삭제-][+추가+] 형태로 표시"""
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


# =========================
# 파일 Reader
# =========================
def _read_pdf_df(file_like) -> pd.DataFrame:
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
                df = normalize_columns(df)
                frames.append(df)

    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def _read_docx_df(file_like) -> pd.DataFrame:
    doc = Document(file_like)
    frames = []
    for table in doc.tables:
        rows = []
        for row in table.rows:
            rows.append([_strip(cell.text) for cell in row.cells])
        if len(rows) > 1:
            header = _make_unique_columns(rows[0])
            data = rows[1:]
            df = pd.DataFrame(data, columns=header)
            df = normalize_columns(df)
            frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def _read_excel_df(file_like) -> pd.DataFrame:
    return pd.read_excel(file_like)


# =========================
# 공통 로더
# =========================
def load_to_dataframe(file_or_path, sheet: Optional[str] = None) -> pd.DataFrame:
    """
    엑셀(.xlsx/.xls), PDF(.pdf), Word(.doc/.docx) → DataFrame 변환
      - 컬럼명/값 정규화
      - 프로젝트명 없는 행 제거
      - 프로젝트명 중복 시 마지막 행 우선
    """
    name = getattr(file_or_path, "name", str(file_or_path)).lower()
    if name.endswith(".pdf"):
        df = _read_pdf_df(file_or_path)
    elif name.endswith((".doc", ".docx")):
        df = _read_docx_df(file_or_path)
    elif name.endswith((".xlsx", ".xls")):
        df = _read_excel_df(file_or_path)
    else:
        raise ValueError("지원하지 않는 파일 형식입니다. (.pdf, .doc, .docx, .xls, .xlsx)")

    # 정규화
    df = normalize_columns(df)
    for k in ["프로젝트명", "런칭", "금주 진행 업무"]:
        if k not in df.columns:
            df[k] = pd.Series(dtype=str)
    df = df[["프로젝트명", "런칭", "금주 진행 업무"]]
    df = normalize_values(df)

    # 프로젝트명 없는 행 제거
    df = df[~df["프로젝트명"].str.strip().eq("")]

    # 중복 프로젝트명 처리 (마지막 행만 유지)
    if len(df):
        df = df.groupby("프로젝트명", as_index=False).last()

    return df


# =========================
# 비교 & 엑셀 출력
# =========================
def build_report(prev_df: pd.DataFrame,
                 curr_df: pd.DataFrame,
                 project_col: str,
                 launch_col: str,
                 work_col: str):
    """
    동일 컬럼끼리 비교:
      - 런칭 ↔ 런칭
      - 금주 진행 업무 ↔ 금주 진행 업무
    상태: ADDED / REMOVED / MODIFIED / UNCHANGED
    """
    for df in (prev_df, curr_df):
        for k in [project_col, launch_col, work_col]:
            if k not in df.columns:
                df[k] = pd.Series(dtype=str)

    prev_df = prev_df[[project_col, launch_col, work_col]].copy()
    curr_df = curr_df[[project_col, launch_col, work_col]].copy()

    merged = curr_df.merge(
        prev_df,
        on=project_col,
        how="outer",
        suffixes=("_curr", "_prev"),
        indicator=True,
    )

    def status_row(r):
        if r["_merge"] == "left_only":
            return "ADDED"
        if r["_merge"] == "right_only":
            return "REMOVED"
        launch_changed = (r.get(f"{launch_col}_curr") != r.get(f"{launch_col}_prev"))
        work_changed = (r.get(f"{work_col}_curr") != r.get(f"{work_col}_prev"))
        return "MODIFIED" if (launch_changed or work_changed) else "UNCHANGED"

    merged["STATUS"] = merged.apply(status_row, axis=1)

    modified = merged[merged["STATUS"] == "MODIFIED"][
        [project_col, f"{launch_col}_prev", f"{launch_col}_curr",
         f"{work_col}_prev", f"{work_col}_curr"]
    ].copy()
    modified["업무_diff"] = [
        make_inline_diff(a, b)
        for a, b in zip(modified[f"{work_col}_prev"], modified[f"{work_col}_curr"])
    ]

    added = merged[merged["STATUS"] == "ADDED"][
        [project_col, f"{launch_col}_curr", f"{work_col}_curr"]
    ].copy().rename(columns={
        f"{launch_col}_curr": launch_col,
        f"{work_col}_curr": work_col
    })

    removed = merged[merged["STATUS"] == "REMOVED"][
        [project_col, f"{launch_col}_prev", f"{work_col}_prev"]
    ].copy().rename(columns={
        f"{launch_col}_prev": launch_col,
        f"{work_col}_prev": work_col
    })

    return merged, modified, added, removed


def write_excel(out_path: str,
                merged: pd.DataFrame,
                modified: pd.DataFrame,
                added: pd.DataFrame,
                removed: pd.DataFrame,
                project_col: str,
                launch_col: str,
                work_col: str):
    """
    결과 엑셀 생성:
      - Summary: 전/금주 값 + STATUS
      - Modified: 변경 행만 (현재값 하이라이트 + 업무_diff)
      - Added, Removed
    """
    with pd.ExcelWriter(out_path, engine="openpyxl") as wr:
        summary_cols = [
            project_col,
            f"{launch_col}_prev", f"{launch_col}_curr",
            f"{work_col}_prev", f"{work_col}_curr",
            "STATUS"
        ]
        keep = [c for c in summary_cols if c in merged.columns]
        merged[keep].to_excel(wr, "Summary", index=False)
        modified.to_excel(wr, "Modified", index=False)
        added.to_excel(wr, "Added", index=False)
        removed.to_excel(wr, "Removed", index=False)

    wb = load_workbook(out_path)
    if "Modified" in wb.sheetnames:
        ws = wb["Modified"]
        yellow = PatternFill(start_color="FFF8B4", end_color="FFF8B4", fill_type="solid")
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=3).fill = yellow  # launch_curr
            ws.cell(row=r, column=5).fill = yellow  # work_curr
    wb.save(out_path)
