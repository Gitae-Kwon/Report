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


# =============== 엑셀 내보내기 (부분 서식 + 줄바꿈) ===============
def _split_diff_for_rich(diff_text: str) -> List[Tuple[str, str]]:
    """'...[-삭제-]동일[+추가+]...' → [("equal","..."),("del","삭제"),("equal","동일"),("add","추가"), ...]"""
    parts: List[Tuple[str, str]] = []
    s = diff_text or ""
    i = 0
    pat = re.compile(r'(\[-.*?-\]|\[\+.*?\+\])', re.S)
    for m in pat.finditer(s):
        if m.start() > i:
            parts.append(("equal", s[i:m.start()]))
        tok = m.group(0)
        if tok.startswith("[-"):
            parts.append(("del", tok[2:-2]))
        else:
            parts.append(("add", tok[2:-2]))
        i = m.end()
    if i < len(s):
        parts.append(("equal", s[i:]))
    return parts

def _rich_from_segments_for_prev(segments, fmt_equal, fmt_del):
    """Prev 셀: equal + del만 표시 (add는 prev에 없음)"""
    rich = []
    for kind, text in segments:
        chunk = (text or "").replace("\r", "\n")
        if kind == "equal":
            rich += [fmt_equal, chunk]
        elif kind == "del":
            rich += [fmt_del, chunk]
    return rich or [fmt_equal, ""]

def _rich_from_segments_for_curr(segments, fmt_equal, fmt_add):
    """Curr 셀: equal + add만 표시 (del은 curr에 없음)"""
    rich = []
    for kind, text in segments:
        chunk = (text or "").replace("\r", "\n")
        if kind == "equal":
            rich += [fmt_equal, chunk]
        elif kind == "add":
            rich += [fmt_add, chunk]
    return rich or [fmt_equal, ""]

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
    - 본문 줄바꿈 보존 (wrap)
    - Prev/Curr 본문에 부분 서식으로 변경점 하이라이트
    - 런칭 변경 시 Curr를 노란 배경
    """
    import xlsxwriter  # ensure installed

    with pd.ExcelWriter(out_path, engine="xlsxwriter") as wr:
        wb = wr.book

        # 공통 포맷
        wrap = wb.add_format({"text_wrap": True, "valign": "top"})
        fmt_equal = wb.add_format({})
        fmt_del   = wb.add_format({"font_color": "#0000ff", "font_strikeout": True})
        fmt_add   = wb.add_format({"font_color": "#ff0000", "bold": True})
        fill_yel  = wb.add_format({"bg_color": "#FFF8B4"})

        # ---------- Summary ----------
        summary_cols = [
            project_col,
            f"{launch_col}_prev", f"{launch_col}_curr",
            f"{work_col}_prev",  f"{work_col}_curr",
            "STATUS"
        ]
        keep = [c for c in summary_cols if c in merged.columns]
        merged[keep].to_excel(wr, sheet_name="Summary", index=False)
        ws_sum = wr.sheets["Summary"]
        ws_sum.set_column(0, len(keep)-1, 42, wrap)

        # Summary prev/curr 본문에 부분서식 적용
        if {f"{work_col}_prev", f"{work_col}_curr"}.issubset(merged.columns):
            col_prev = keep.index(f"{work_col}_prev")
            col_curr = keep.index(f"{work_col}_curr")
            col_lprev = keep.index(f"{launch_col}_prev")
            col_lcurr = keep.index(f"{launch_col}_curr") if f"{launch_col}_curr" in keep else None

            for r, row in merged.iterrows():
                diff = make_inline_diff(row.get(f"{work_col}_prev", ""), row.get(f"{work_col}_curr", ""))
                segs = _split_diff_for_rich(diff)

                # prev 본문
                rich_prev = _rich_from_segments_for_prev(segs, fmt_equal, fmt_del)
                ws_sum.write_rich_string(r+1, col_prev, *rich_prev, wrap)

                # curr 본문
                rich_curr = _rich_from_segments_for_curr(segs, fmt_equal, fmt_add)
                ws_sum.write_rich_string(r+1, col_curr, *rich_curr, wrap)

                # 런칭 변경 하이라이트 (curr만)
                if col_lcurr is not None:
                    if str(row.get(f"{launch_col}_prev","")) != str(row.get(f"{launch_col}_curr","")):
                        ws_sum.write(r+1, col_lcurr, row.get(f"{launch_col}_curr",""), fill_yel)

        # ---------- Modified ----------
        modified.to_excel(wr, sheet_name="Modified", index=False)
        ws_mod = wr.sheets["Modified"]
        ws_mod.set_column(0, len(modified.columns)-1, 42, wrap)

        cols = list(modified.columns)
        col_lp = cols.index(f"{launch_col}_prev")
        col_lc = cols.index(f"{launch_col}_curr")
        col_wp = cols.index(f"{work_col}_prev")
        col_wc = cols.index(f"{work_col}_curr")
        col_diff = cols.index("업무_diff") if "업무_diff" in cols else None

        for r, row in modified.iterrows():
            # diff 세그먼트
            segs = _split_diff_for_rich(row.get("업무_diff", "") if col_diff is not None else
                                        make_inline_diff(row.get(f"{work_col}_prev",""), row.get(f"{work_col}_curr","")))

            # prev 본문
            rich_prev = _rich_from_segments_for_prev(segs, fmt_equal, fmt_del)
            ws_mod.write_rich_string(r+1, col_wp, *rich_prev, wrap)

            # curr 본문
            rich_curr = _rich_from_segments_for_curr(segs, fmt_equal, fmt_add)
            ws_mod.write_rich_string(r+1, col_wc, *rich_curr, wrap)

            # diff 컬럼(요약)도 부분서식으로 재작성
            if col_diff is not None:
                rich_all = []
                for kind, text in segs:
                    fmt = fmt_equal if kind=="equal" else (fmt_del if kind=="del" else fmt_add)
                    rich_all += [fmt, (text or "").replace("\r","\n")]
                if not rich_all:
                    rich_all = [fmt_equal, ""]
                ws_mod.write_rich_string(r+1, col_diff, *rich_all, wrap)

            # 런칭 변경 하이라이트 (curr만)
            if str(row.get(f"{launch_col}_prev","")) != str(row.get(f"{launch_col}_curr","")):
                ws_mod.write(r+1, col_lc, row.get(f"{launch_col}_curr",""), fill_yel)

        # ---------- Added / Removed ----------
        if not added.empty:
            added.to_excel(wr, sheet_name="Added", index=False)
            wr.sheets["Added"].set_column(0, len(added.columns)-1, 42, wrap)
        if not removed.empty:
            removed.to_excel(wr, sheet_name="Removed", index=False)
            wr.sheets["Removed"].set_column(0, len(removed.columns)-1, 42, wrap)
