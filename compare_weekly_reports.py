# compare_weekly_reports.py
# -*- coding: utf-8 -*-
import re
import difflib
from collections import defaultdict
from typing import Optional, List, Tuple

import pandas as pd

import pdfplumber
from docx import Document


# =============== 공통 유틸 ===============
def _strip_header(s):
    """헤더용: 개행은 공백으로, 양끝 공백 제거"""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return str(s).replace("\r", " ").replace("\n", " ").strip()

def _strip_keep_nl(s):
    """본문용: 개행(\n) 보존, 라인별 trim"""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    txt = str(s).replace("\r", "\n")
    lines = [ln.strip() for ln in txt.split("\n")]
    return "\n".join(lines)

def _make_unique_columns(cols):
    counts = defaultdict(int)
    out = []
    for c in cols:
        c = _strip_header(c)
        counts[c] += 1
        out.append(c if counts[c] == 1 else f"{c}_{counts[c]}")
    return out

def _norm_key(s: str) -> str:
    return re.sub(r"\s+", "", s or "")

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """여러 표현 → 표준 컬럼명 통일"""
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

def normalize_values_keep_nl(df: pd.DataFrame) -> pd.DataFrame:
    """문자열 컬럼 줄바꿈 보존 + trim"""
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].map(_strip_keep_nl)
    return df

def make_inline_diff(a: str, b: str) -> str:
    """[-삭제-][+추가+] 마크업 생성 (화면/엑셀 렌더의 소스)"""
    a = "" if pd.isna(a) else str(a)
    b = "" if pd.isna(b) else str(b)
    if a == b:
        return a
    sm = difflib.SequenceMatcher(a=a, b=b)
    pieces: List[str] = []
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


# =============== 리더 ===============
def _read_pdf_df(file_like) -> pd.DataFrame:
    frames = []
    with pdfplumber.open(file_like) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables() or []
            for tbl in tables:
                if not tbl or len(tbl) < 2:
                    continue
                header = _make_unique_columns(tbl[0])
                rows = [[_strip_keep_nl(x) for x in r] for r in tbl[1:]]
                df = pd.DataFrame(rows, columns=header)
                df = normalize_columns(df)
                frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

def _read_docx_df(file_like) -> pd.DataFrame:
    doc = Document(file_like)  # .docx만 지원
    frames = []
    for table in doc.tables:
        rows = []
        for row in table.rows:
            rows.append([_strip_keep_nl(cell.text) for cell in row.cells])
        if len(rows) > 1:
            header = _make_unique_columns(rows[0])
            data = rows[1:]
            df = pd.DataFrame(data, columns=header)
            df = normalize_columns(df)
            frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

def _read_excel_df(file_like, sheet: Optional[str] = None) -> pd.DataFrame:
    if hasattr(file_like, "seek"):
        try: file_like.seek(0)
        except Exception: pass
    df = pd.read_excel(file_like, sheet_name=(sheet if sheet is not None else 0))
    if isinstance(df, dict):
        for v in df.values():
            if isinstance(v, pd.DataFrame) and not v.empty:
                return v
        return list(df.values())[0]
    # 엑셀도 개행 문자 보존
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].map(_strip_keep_nl)
    return df


# =============== 로더(공통) ===============
def load_to_dataframe(file_or_path, sheet: Optional[str] = None) -> pd.DataFrame:
    name = getattr(file_or_path, "name", str(file_or_path)).lower()
    if hasattr(file_or_path, "seek"):
        try: file_or_path.seek(0)
        except Exception: pass

    if name.endswith(".pdf"):
        df = _read_pdf_df(file_or_path)
    elif name.endswith(".docx"):
        df = _read_docx_df(file_or_path)
    elif name.endswith((".xlsx", ".xls")):
        df = _read_excel_df(file_or_path, sheet=sheet)
    else:
        raise ValueError("지원하지 않는 파일 형식입니다. (.pdf, .docx, .xls, .xlsx)")

    df = normalize_columns(df)
    for k in ["프로젝트명", "런칭", "금주 진행 업무"]:
        if k not in df.columns:
            df[k] = pd.Series(dtype=str)
    df = df[["프로젝트명", "런칭", "금주 진행 업무"]]
    df = normalize_values_keep_nl(df)

    # 프로젝트명 없는 행 제거 + 중복 마지막 유지
    df = df[~df["프로젝트명"].str.strip().eq("")]
    if len(df):
        df = df.groupby("프로젝트명", as_index=False).last()
    return df


# =============== 비교 ===============
def build_report(prev_df: pd.DataFrame,
                 curr_df: pd.DataFrame,
                 project_col: str,
                 launch_col: str,
                 work_col: str):
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
        if r["_merge"] == "left_only":  return "ADDED"
        if r["_merge"] == "right_only": return "REMOVED"
        launch_changed = (r.get(f"{launch_col}_curr") != r.get(f"{launch_col}_prev"))
        work_changed   = (r.get(f"{work_col}_curr")   != r.get(f"{work_col}_prev"))
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
        f"{work_col}_curr":  work_col
    })
    removed = merged[merged["STATUS"] == "REMOVED"][
        [project_col, f"{launch_col}_prev", f"{work_col}_prev"]
    ].copy().rename(columns={
        f"{launch_col}_prev": launch_col,
        f"{work_col}_prev":  work_col
    })
    return merged, modified, added, removed


# =============== 엑셀 내보내기 (화면 느낌 최대 재현) ===============
def _split_diff_for_rich(diff_text: str) -> List[Tuple[str, str]]:
    """
    '...[-삭제-][+추가+]...' → [("equal","..."),("del","삭제"),("add","추가"), ...]
    """
    parts: List[Tuple[str, str]] = []
    i = 0
    s = diff_text or ""
    pattern = re.compile(r'(\[-.*?-\]|\[\+.*?\+\])', re.S)
    for m in pattern.finditer(s):
        if m.start() > i:
            parts.append(("equal", s[i:m.start()]))
        token = m.group(0)
        if token.startswith("[-"):
            parts.append(("del", token[2:-2]))
        else:
            parts.append(("add", token[2:-2]))
        i = m.end()
    if i < len(s):
        parts.append(("equal", s[i:]))
    return parts

def write_excel(out_path: str,
                merged: pd.DataFrame,
                modified: pd.DataFrame,
                added: pd.DataFrame,
                removed: pd.DataFrame,
                project_col: str,
                launch_col: str,
                work_col: str):
    """
    - Summary/Modified/Added/Removed 시트 생성
    - 줄바꿈 보존 (wrap)
    - Modified! 업무_diff는 부분 서식(삭제=빨강+취소선 / 추가=초록+볼드)로 표시
    """
    # 1) xlsxwriter 엔진 사용 (부분서식 지원)
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as wr:
        wb  = wr.book

        # 공통 포맷
        wrap = wb.add_format({"text_wrap": True, "valign": "top"})
        bold = wb.add_format({"bold": True})
        # diff segment 포맷
        fmt_equal = wb.add_format({})
        fmt_del   = wb.add_format({"font_color": "#c62828", "font_strikeout": True})
        fmt_add   = wb.add_format({"font_color": "#1b5e20", "bold": True})

        # --- Summary
        summary_cols = [
            project_col,
            f"{launch_col}_prev", f"{launch_col}_curr",
            f"{work_col}_prev",  f"{work_col}_curr",
            "STATUS"
        ]
        keep = [c for c in summary_cols if c in merged.columns]
        merged[keep].to_excel(wr, sheet_name="Summary", index=False)
        ws = wr.sheets["Summary"]
        ws.set_column(0, len(keep)-1, 40, wrap)

        # --- Modified (일단 DataFrame으로 쓰고 diff는 부분서식으로 다시 씀)
        modified.to_excel(wr, sheet_name="Modified", index=False)
        ws = wr.sheets["Modified"]
        ws.set_column(0, len(modified.columns)-1, 45, wrap)

        # 업무_diff 부분서식 재작성
        if "업무_diff" in modified.columns:
            diff_col_idx = list(modified.columns).index("업무_diff")
            for r, diff in enumerate(modified["업무_diff"], start=2):  # 1-based + header
                segments = _split_diff_for_rich(diff)
                rich: List = []
                for kind, text in segments:
                    # 엑셀 줄바꿈은 '\n'
                    chunk = text.replace("\r", "\n")
                    fmt = fmt_equal if kind == "equal" else (fmt_del if kind == "del" else fmt_add)
                    rich.extend([fmt, chunk])
                # 최소 하나는 필요
                if not rich:
                    rich = [fmt_equal, ""]
                ws.write_rich_string(r-1, diff_col_idx, *rich, wrap)

        # --- Added / Removed
        if not added.empty:
            added.to_excel(wr, sheet_name="Added", index=False)
            wr.sheets["Added"].set_column(0, len(added.columns)-1, 45, wrap)
        if not removed.empty:
            removed.to_excel(wr, sheet_name="Removed", index=False)
            wr.sheets["Removed"].set_column(0, len(removed.columns)-1, 45, wrap)
