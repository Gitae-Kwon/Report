# compare_weekly_reports.py
# -*- coding: utf-8 -*-
import re
from collections import defaultdict
import difflib
from typing import Optional

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


# =========================
# 공통 유틸
# =========================
def _strip(s):
    if pd.isna(s):
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
    """공백/개행 제거하여 키 비교 용도로 사용"""
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
        # 프로젝트명
        if ("프로젝트" in key and "명" in key) or key in ("프로젝트명", "프로젝트"):
            rename_map[col] = "프로젝트명"
            continue
        # 런칭 / 런칭(예정) / 오픈
        if ("런칭" in key) or ("오픈" in key):
            rename_map[col] = "런칭"
            continue
        # 금주 진행 업무
        if ("금주" in key) and (("업무" in key) or ("진행" in key)):
            rename_map[col] = "금주 진행 업무"
            continue
    df = df.rename(columns=rename_map)
    return df


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
# 입력 로더
# =========================
def _read_pdf_df(file_like) -> pd.DataFrame:
    """
    pdfplumber로 PDF 표를 읽어서 표준 컬럼(프로젝트명/런칭/금주 진행 업무)만 추출 후 결합.
    - 헤더 중복, 변형 헤더명 처리
    - 완전 빈 행 제거
    """
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

                df = normalize_columns(df)
                keep = ["프로젝트명", "런칭", "금주 진행 업무"]
                # 없는 컬럼은 빈 시리즈로 추가해도 concat 가능
                for k in keep:
                    if k not in df.columns:
                        df[k] = pd.Series(dtype=str)
                df = df[keep]

                # 완전 빈 행 제거
                mask_all_empty = (
                    df["프로젝트명"].str.strip().eq("") &
                    df["런칭"].str.strip().eq("") &
                    df["금주 진행 업무"].str.strip().eq("")
                )
                df = df[~mask_all_empty]
                if len(df):
                    frames.append(df)

    if not frames:
        return pd.DataFrame(columns=["프로젝트명", "런칭", "금주 진행 업무"])

    out = pd.concat(frames, ignore_index=True)
    return out


def load_to_dataframe(file_or_path, sheet: Optional[str] = None) -> pd.DataFrame:
    """
    엑셀(.xlsx/.xls) 또는 PDF(.pdf)을 DataFrame으로 변환 후:
      - 컬럼명/값 정규화
      - 프로젝트명 없는 행 제거
      - 프로젝트명 중복 시 마지막 행 유지
    """
    name = getattr(file_or_path, "name", str(file_or_path)).lower()

    if name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(file_or_path, sheet_name=sheet)
    elif name.endswith(".pdf"):
        df = _read_pdf_df(file_or_path)
    else:
        raise ValueError("지원하지 않는 파일 형식입니다. (.xlsx, .xls, .pdf)")

    # 정규화
    df = normalize_columns(df)
    # 필요한 컬럼 강제 확보
    for k in ["프로젝트명", "런칭", "금주 진행 업무"]:
        if k not in df.columns:
            df[k] = pd.Series(dtype=str)
    df = df[["프로젝트명", "런칭", "금주 진행 업무"]]
    df = normalize_values(df)

    # 프로젝트명 없는 행 제거
    df = df[~df["프로젝트명"].str.strip().eq("")]

    # 프로젝트명 중복 시 마지막 행 우선 사용
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
    동일 컬럼끼리만 비교:
      - 런칭 ↔ 런칭
      - 금주 진행 업무 ↔ 금주 진행 업무
    상태: ADDED / REMOVED / MODIFIED / UNCHANGED
    """
    # 필요한 컬럼만 보장
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
        # 둘 다 있는 경우에만 값 비교
        launch_changed = (r.get(f"{launch_col}_curr") != r.get(f"{launch_col}_prev"))
        work_changed = (r.get(f"{work_col}_curr") != r.get(f"{work_col}_prev"))
        return "MODIFIED" if (launch_changed or work_changed) else "UNCHANGED"

    merged["STATUS"] = merged.apply(status_row, axis=1)

    modified = merged[merged["STATUS"] == "MODIFIED"][
        [project_col, f"{launch_col}_prev", f"{launch_col}_curr",
         f"{work_col}_prev", f"{work_col}_curr"]
    ].copy()

    # 업무 텍스트 인라인 diff
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
        # Summary에는 표준 5컬럼 + STATUS만 표시
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

    # 하이라이트(Modified 시트의 현재값 컬럼)
    wb = load_workbook(out_path)
    if "Modified" in wb.sheetnames:
        ws = wb["Modified"]
        yellow = PatternFill(start_color="FFF8B4", end_color="FFF8B4", fill_type="solid")
        # Modified 컬럼 순서: 프로젝트명, launch_prev, launch_curr, work_prev, work_curr, 업무_diff
        # 현재값: launch_curr(3), work_curr(5)
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=3).fill = yellow
            ws.cell(row=r, column=5).fill = yellow
    wb.save(out_path)
