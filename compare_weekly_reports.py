import pandas as pd
import re
from collections import defaultdict

def _make_unique_columns(cols):
    """
    같은 이름의 컬럼이 여러 번 나오면 name, name_2, name_3... 형태로 바꿔줌
    """
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
        # 비교용 키(공백/개행 제거)
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

                # 2) 헤더 매핑: 다양한 변형을 '프로젝트명/런칭/금주 진행 업무'로 맞춤
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

                # 3) 필요한 컬럼만 뽑아 표준 형태로 재구성
                keep_cols = ["프로젝트명", "런칭", "금주 진행 업무"]
                has_any = any(c in df.columns for c in keep_cols)
                if not has_any:
                    continue  # 이 테이블에는 우리가 원하는 컬럼이 없음

                std = pd.DataFrame({
                    "프로젝트명": df["프로젝트명"] if "프로젝트명" in df.columns else "",
                    "런칭": df["런칭"] if "런칭" in df.columns else "",
                    "금주 진행 업무": df["금주 진행 업무"] if "금주 진행 업무" in df.columns else "",
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

    # 표준 3컬럼만 concat → 더 이상 중복 컬럼 이슈 없음
    out = pd.concat(normalized_frames, ignore_index=True)
    return out
