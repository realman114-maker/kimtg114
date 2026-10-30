# -*- coding: utf-8 -*-
"""
대리점 수수료 정산 에이전트 (Streamlit)

본사에서 매월 받는 '지역_정산마스터' 파일(xlsb/xlsx)에서
  - '요약' 시트(대리점별 요약)와 '인센티브' 시트(개통 건별 상세)만 사용하고
  - 담당 대리점(대리점명 기준)만 필터링한 뒤
  - 공란(필요하면 0 포함) 항목을 제거하여
  - 표 형식 PDF + 분석 리포트 + 엑셀 파일로 내려받을 수 있게 합니다.

[공유해서 쓰는 방법]
  관리자가 정산 파일을 한 번 올려 '공유용으로 저장'하면, 접속 코드를 받은 사람은
  파일 업로드 없이 자기 담당 대리점만 골라 PDF/엑셀을 받을 수 있습니다.
  Streamlit secrets 예시:
      ADMIN_PASSWORD = "관리자용-비밀번호"
      [USERS]
      "1234" = ["베스트원", "파트너"]
      "5678" = ["SN네트워크"]
  secrets를 설정하지 않으면 기존처럼 개인용(업로드해서 바로 사용)으로 동작합니다.

실행:  streamlit run app.py
"""
import hmac
import io
import json
import os
import re
from datetime import datetime
from xml.sax.saxutils import escape

import numpy as np
import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (PageBreak, Paragraph, SimpleDocTemplate,
                                Spacer, Table, TableStyle)

# ─────────────────────────────────────────────────────────────
# 설정 (여기만 수정하면 됩니다)
# ─────────────────────────────────────────────────────────────
# 내가 담당하는 대리점명을 파일에 적힌 그대로 적어두면 매달 자동 선택됩니다.
# 예) MY_AGENCIES = ["베스트원", "파트너", "SN네트워크"]
MY_AGENCIES = []

# PDF '개통 상세' 표에 보여줄 열 (상세 시트 기준, 없는 열/전부 공란인 열은 자동 제외)
DETAIL_COLUMNS = [
    "개통일자", "서비스관리번호", "고객명", "요금상품", "약정할인개월명",
    "번들유형구분명", "시군구명", "사은품합계금액", "추정_인센티브(본사)", "해지일",
]
# 개통 상세 표 맨 아래에 합계를 낼 열
DETAIL_SUM_COLUMNS = ["사은품합계금액", "추정_인센티브(본사)"]

# 공유용으로 저장한 정산 파일이 놓일 폴더 (영구 저장되는 디스크 경로로 바꿀 수 있음)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.environ.get("SETTLEMENT_DATA_DIR", os.path.join(BASE_DIR, "shared_data"))

ID_COLS = {"센터명", "대리점명", "合", "합", "합계"}
NO_COMMA_HINTS = ("코드", "번호", "ID", "id")


# ─────────────────────────────────────────────────────────────
# 파일 읽기
# ─────────────────────────────────────────────────────────────
def _clean(name):
    return re.sub(r"\s+", " ", str(name)).strip()


def _engine(filename):
    return "pyxlsb" if filename.lower().endswith(".xlsb") else None


def _header_row(raw):
    """위에서 20행 안에서 '대리점명' 헤더가 있는 행 번호를 찾는다. 없으면 None."""
    for i in range(min(20, len(raw))):
        vals = [_clean(v) for v in raw.iloc[i].tolist() if pd.notna(v)]
        if "대리점명" in vals:
            return i
    return None


def _find_sheets(xl):
    """사용할 시트 2개(요약, 인센티브 상세)의 위치를 찾는다.
    1) 시트 이름이 '요약' / '인센티브...'로 시작하는 시트를 우선 사용
    2) 없으면 '대리점명' 헤더가 있는 앞쪽 시트 2개를 사용
    (월·상품에 따라 맨 앞에 '수정사항' 같은 시트가 끼어 있어도 동작하도록 이름 기준으로 찾음)"""
    names = [str(n).strip() for n in xl.sheet_names]
    s_idx = next((i for i, n in enumerate(names) if n == "요약"), None)
    d_idx = next((i for i, n in enumerate(names) if n.startswith("인센티브")), None)
    if s_idx is not None and d_idx is not None:
        return s_idx, d_idx
    found = []
    for i in range(len(names)):
        raw = xl.parse(sheet_name=i, header=None, nrows=20)
        if _header_row(raw) is not None:
            found.append(i)
        if len(found) == 2:
            return found[0], found[1]
    raise ValueError("'대리점명' 헤더가 있는 시트 2개를 찾지 못했습니다. 시트 목록: " + ", ".join(names[:12]))


def _parse_sheet(xl, idx):
    """시트를 읽어 '대리점명' 헤더 행을 자동으로 찾고 (DataFrame, 그룹 라벨 dict)를 돌려준다."""
    raw = xl.parse(sheet_name=idx, header=None)
    hdr = _header_row(raw)
    if hdr is None:
        raise ValueError(f"'{xl.sheet_names[idx]}' 시트에서 '대리점명' 헤더를 찾지 못했습니다.")

    names, seen = [], {}
    for j, v in enumerate(raw.iloc[hdr].tolist()):
        n = _clean(v) if pd.notna(v) else f"(열{j + 1})"
        if n in seen:
            seen[n] += 1
            n = f"{n}.{seen[n]}"
        else:
            seen[n] = 0
        names.append(n)

    groups = {}
    if hdr > 0:
        g = raw.iloc[hdr - 1].ffill()
        for j, n in enumerate(names):
            groups[n] = _clean(g.iloc[j]) if pd.notna(g.iloc[j]) else ""

    df = raw.iloc[hdr + 1:].copy()
    df.columns = names
    df = df.reset_index(drop=True)
    df = df.replace(r"^\s*$", pd.NA, regex=True)
    df = df[df["대리점명"].notna()].copy()
    df["대리점명"] = df["대리점명"].astype(str).str.strip()
    df = df.infer_objects()
    return df, groups


@st.cache_data(show_spinner="파일을 읽는 중입니다...")
def load_workbook(data, filename):
    xl = pd.ExcelFile(io.BytesIO(data), engine=_engine(filename))
    s_idx, d_idx = _find_sheets(xl)
    sheet_names = [xl.sheet_names[s_idx], xl.sheet_names[d_idx]]
    summary, groups = _parse_sheet(xl, s_idx)
    detail, _ = _parse_sheet(xl, d_idx)

    # 그룹 라벨(본사/담당/영업팀)이 없으면 숫자 열 전체를 하나의 그룹으로 취급
    if not any(groups.get(c) for c in summary.columns if c not in ID_COLS):
        groups = {c: "정산항목" for c in summary.columns
                  if c not in ID_COLS and pd.api.types.is_numeric_dtype(summary[c])}
    return summary, groups, detail, sheet_names


# ─────────────────────────────────────────────────────────────
# 필터 / 정리
# ─────────────────────────────────────────────────────────────
def drop_empty_columns(df, drop_zero):
    keep = []
    for c in df.columns:
        s = df[c]
        if s.isna().all():
            continue
        if drop_zero and pd.api.types.is_numeric_dtype(s) and (s.fillna(0) == 0).all():
            continue
        keep.append(c)
    return df[keep]


def item_columns(summary, groups):
    return [c for c in summary.columns if groups.get(c) and c not in ID_COLS]


def agency_totals(summary, groups):
    """대리점별 구분(본사/담당/영업팀)별 합계와 총합계. 원본 '合' 열은 쓰지 않고 항목을 직접 합산."""
    cols = item_columns(summary, groups)
    num = summary[cols].apply(pd.to_numeric, errors="coerce").fillna(0)
    out = pd.DataFrame({
        "센터명": summary["센터명"] if "센터명" in summary else "",
        "대리점명": summary["대리점명"],
    })
    order = list(dict.fromkeys(groups[c] for c in cols))
    for g in order:
        out[g] = num[[c for c in cols if groups[c] == g]].sum(axis=1)
    out["합계"] = num.sum(axis=1)
    return out.reset_index(drop=True), order


def build_items(summary_f, groups, drop_zero):
    """선택 대리점의 정산 항목을 (대리점, 구분, 항목, 금액) 세로 표로 변환. 공란은 제외."""
    rows = []
    cols = item_columns(summary_f, groups)
    for _, r in summary_f.iterrows():
        for c in cols:
            v = pd.to_numeric(r[c], errors="coerce")
            if pd.isna(v) or (drop_zero and v == 0):
                continue
            rows.append((r["대리점명"], groups[c], c, float(v)))
    return pd.DataFrame(rows, columns=["대리점명", "구분", "항목", "금액"])


def item_detail(detail_f, agency, item):
    """정산 항목 1개에 해당하는 개통 건별 내역을 돌려준다.
    요약 시트의 정산 항목 열은 개통 상세 시트의 같은 이름 열을 대리점별로 합산한 값이므로,
    해당 열의 값이 0/공란이 아닌 건만 추려서 보여준다.
    반환: (표시용 DataFrame, 건별 합계, 건수)  /  상세 시트에 같은 이름의 열이 없으면 None"""
    if item not in detail_f.columns:
        return None
    d = detail_f[detail_f["대리점명"] == agency]
    amt = pd.to_numeric(d[item], errors="coerce").fillna(0)
    keep = amt != 0
    d, amt = d[keep], amt[keep]
    if "개통일자" in d.columns:
        order = d["개통일자"].sort_values().index
        d, amt = d.loc[order], amt.loc[order]
    cols = [c for c in DETAIL_COLUMNS
            if c in d.columns and c not in DETAIL_SUM_COLUMNS and d[c].notna().any()]
    out = pd.DataFrame({c: [fmt_cell(c, v) for v in d[c]] for c in cols})
    out["해당 항목 금액"] = [won(v) for v in amt]
    return out, float(amt.sum()), len(d)


# ─────────────────────────────────────────────────────────────
# 서식 도우미
# ─────────────────────────────────────────────────────────────
def won(v):
    return f"{int(round(float(v))):,}"


def pct(a, b):
    return f"{(a / b * 100):.1f}%" if b else "-"


def fmt_cell(col, v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    if isinstance(v, (int, float, np.integer, np.floating)):
        if float(v).is_integer():
            iv = int(v)
            if ("일자" in col or col.endswith("일")) and len(str(abs(iv))) == 8:
                s = str(iv)
                return f"{s[:4]}-{s[4:6]}-{s[6:]}"
            if any(h in col for h in NO_COMMA_HINTS):
                return str(iv)
            return f"{iv:,}"
        return f"{float(v):,.2f}"
    return str(v)


def month_label(filename):
    m = re.search(r"_(\d{2})(\d{2})_", filename)
    if m and 1 <= int(m.group(2)) <= 12:
        return f"20{m.group(1)}년 {int(m.group(2))}월"
    return ""


# ─────────────────────────────────────────────────────────────
# 분석 리포트 (규칙 기반 자동 생성)
# ─────────────────────────────────────────────────────────────
def top_counts(s, n=5):
    vc = s.dropna().astype(str).value_counts()
    total = int(vc.sum())
    return [f"{k} {v:,}건({pct(v, total)})" for k, v in vc.head(n).items()]


def make_analysis(totals_all, totals_sel, order, items, detail_f):
    sections = []
    total = totals_sel["합계"].sum()
    n = len(totals_sel)

    lines = [f"선택한 {n}개 대리점의 정산 합계는 {won(total)}원이며, 대리점 평균은 {won(total / n)}원입니다."]
    lines.append("구분별 구성: " + ", ".join(
        f"{g} {won(totals_sel[g].sum())}원({pct(totals_sel[g].sum(), total)})" for g in order) + ".")
    all_sum = totals_all.groupby("대리점명")["합계"].sum()
    ranks = all_sum.rank(ascending=False, method="min")
    for _, r in totals_sel.sort_values("합계", ascending=False).iterrows():
        nm = r["대리점명"]
        rk = f"전체 {len(all_sum)}개 대리점 중 {int(ranks[nm])}위" if nm in ranks else ""
        lines.append(f"{nm}: 합계 {won(r['합계'])}원 (선택 대리점 내 비중 {pct(r['합계'], total)}) {rk}".strip())
    sections.append(("1. 정산 총괄", lines))

    lines = []
    if len(items):
        g = items.groupby(["구분", "항목"])["금액"].sum()
        pos = g[g > 0].sort_values(ascending=False)
        neg = g[g < 0].sort_values()
        pos_total = pos.sum()
        if len(pos):
            top5 = pos.head(5)
            lines.append(f"지급 항목은 {len(pos)}개이며, 상위 5개 항목이 지급 합계의 {pct(top5.sum(), pos_total)}를 차지합니다.")
            for i, ((grp, item), v) in enumerate(top5.items(), 1):
                lines.append(f"  {i}) {item} ({grp}): {won(v)}원 ({pct(v, pos_total)})")
        if len(neg):
            lines.append(f"차감 항목은 {len(neg)}개, 합계 {won(neg.sum())}원입니다.")
            for (grp, item), v in neg.head(3).items():
                lines.append(f"  - {item} ({grp}): {won(v)}원")
        else:
            lines.append("차감 항목은 없습니다.")
    else:
        lines.append("표시할 정산 항목이 없습니다.")
    sections.append(("2. 정산 항목 분석", lines))

    lines = []
    if len(detail_f):
        d = detail_f
        lines.append(f"개통 상세는 총 {len(d):,}건입니다. 대리점별: " + ", ".join(top_counts(d["대리점명"], 10)) + ".")
        if "개통일자" in d and d["개통일자"].notna().any():
            dt = pd.to_numeric(d["개통일자"], errors="coerce").dropna().astype(int)
            lines.append(f"개통일자 범위: {fmt_cell('개통일자', dt.min())} ~ {fmt_cell('개통일자', dt.max())}.")
        for col, label in [("요금상품", "요금상품 상위"), ("유선서비스유형명", "서비스 유형 상위"),
                           ("약정할인개월명", "약정 구성"), ("번들유형구분명", "번들 유형 구성"),
                           ("시군구명", "지역(시군구) 상위")]:
            if col in d:
                lines.append(f"{label}: " + ", ".join(top_counts(d[col], 5 if col != "약정할인개월명" else 3)) + ".")
        if "추정_인센티브(본사)" in d:
            inc = pd.to_numeric(d["추정_인센티브(본사)"], errors="coerce").fillna(0)
            lines.append(f"추정 인센티브(본사) 합계 {won(inc.sum())}원, 건당 평균 {won(inc.mean())}원.")
            if "사은품합계금액" in d:
                gift = pd.to_numeric(d["사은품합계금액"], errors="coerce").fillna(0)
                lines.append(f"사은품 합계금액 {won(gift.sum())}원 (추정 인센티브 대비 {pct(gift.sum(), inc.sum())}).")
        if "해지여부" in d:
            y = int((d["해지여부"].astype(str).str.upper() == "Y").sum())
            lines.append(f"해지 건 {y:,}건 ({pct(y, len(d))}).")
    else:
        lines.append("선택한 대리점의 개통 상세 내역이 없습니다.")
    sections.append(("3. 개통 내역 분석", lines))

    sections.append(("4. 참고", [
        "합계는 원본의 '合' 열이 아니라 정산 항목 열을 직접 합산한 값입니다.",
        "공란 항목은 삭제했고, 0원 항목은 설정에 따라 제외했습니다.",
        "분석 문구는 데이터에서 규칙으로 자동 생성한 것이므로 중요한 수치는 원본과 대조해 주세요.",
    ]))
    return sections


# ─────────────────────────────────────────────────────────────
# PDF 생성
# ─────────────────────────────────────────────────────────────
HEADER_BG = colors.HexColor("#1F3A5F")
ZEBRA = colors.HexColor("#F3F5F8")
LINE = colors.HexColor("#C9CFD8")
TOTAL_BG = colors.HexColor("#E3E8EF")
NEG = colors.HexColor("#C62828")


def register_font():
    cands = [
        (os.path.join(BASE_DIR, "fonts", "NanumGothic.ttf"), os.path.join(BASE_DIR, "fonts", "NanumGothicBold.ttf")),
        ("/usr/share/fonts/truetype/nanum/NanumGothic.ttf", "/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf"),
        ("C:/Windows/Fonts/malgun.ttf", "C:/Windows/Fonts/malgunbd.ttf"),
        ("/System/Library/Fonts/Supplemental/AppleGothic.ttf", None),
    ]
    for reg, bold in cands:
        if os.path.exists(reg):
            try:
                pdfmetrics.registerFont(TTFont("KR", reg))
                pdfmetrics.registerFont(TTFont("KR-B", bold if bold and os.path.exists(bold) else reg))
                return "KR", "KR-B"
            except Exception:
                continue
    pdfmetrics.registerFont(UnicodeCIDFont("HYSMyeongJo-Medium"))  # 폰트 파일이 없을 때의 대체 한글 폰트
    return "HYSMyeongJo-Medium", "HYSMyeongJo-Medium"


def _vis_len(s):
    return sum(2 if ord(ch) > 0x2E80 else 1 for ch in str(s))


def _scale(weights, total):
    s = float(sum(weights))
    return [w * total / s for w in weights]


def build_pdf(meta, totals_sel, order, items, detail_f, sections, include_detail):
    font, bold = register_font()
    st_ = {
        "title": ParagraphStyle("t", fontName=bold, fontSize=18, leading=24, textColor=HEADER_BG),
        "h1": ParagraphStyle("h1", fontName=bold, fontSize=12.5, leading=17, textColor=HEADER_BG, spaceBefore=6, spaceAfter=5),
        "h2": ParagraphStyle("h2", fontName=bold, fontSize=10, leading=14, spaceBefore=8, spaceAfter=3),
        "body": ParagraphStyle("b", fontName=font, fontSize=9, leading=13.5, wordWrap="CJK"),
        "meta": ParagraphStyle("m", fontName=font, fontSize=9, leading=13, textColor=colors.HexColor("#555555")),
        "head": ParagraphStyle("hd", fontName=bold, fontSize=7.5, leading=10, textColor=colors.white, alignment=1, wordWrap="CJK"),
        "cell": ParagraphStyle("c", fontName=font, fontSize=7.5, leading=10, wordWrap="CJK"),
        "num": ParagraphStyle("n", fontName=font, fontSize=7.5, leading=10, alignment=2),
        "neg": ParagraphStyle("ng", fontName=font, fontSize=7.5, leading=10, alignment=2, textColor=NEG),
        "cellb": ParagraphStyle("cb", fontName=bold, fontSize=7.5, leading=10, wordWrap="CJK"),
        "numb": ParagraphStyle("nb", fontName=bold, fontSize=7.5, leading=10, alignment=2),
        "negb": ParagraphStyle("ngb", fontName=bold, fontSize=7.5, leading=10, alignment=2, textColor=NEG),
    }

    def P(text, style):
        return Paragraph(escape(str(text)), st_[style])

    def table(header, rows, widths, right=(), total_last=False, align="LEFT"):
        data = [[P(h, "head") for h in header]]
        for i, r in enumerate(rows):
            is_tot = total_last and i == len(rows) - 1
            line = []
            for j, v in enumerate(r):
                txt = "" if v is None else str(v)
                if j in right:
                    key = ("negb" if is_tot else "neg") if txt.startswith("-") else ("numb" if is_tot else "num")
                else:
                    key = "cellb" if is_tot else "cell"
                line.append(P(txt, key))
            data.append(line)
        t = Table(data, colWidths=widths, repeatRows=1, hAlign=align)
        style = [
            ("BACKGROUND", (0, 0), (-1, 0), HEADER_BG),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, ZEBRA]),
            ("GRID", (0, 0), (-1, -1), 0.25, LINE),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 2.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ]
        if total_last and rows:
            style += [("BACKGROUND", (0, len(data) - 1), (-1, len(data) - 1), TOTAL_BG),
                      ("LINEABOVE", (0, len(data) - 1), (-1, len(data) - 1), 0.9, HEADER_BG)]
        t.setStyle(TableStyle(style))
        return t

    page_w, _ = landscape(A4)
    margin = 14 * mm
    usable = page_w - 2 * margin
    story = []

    # 표지/요약 헤더
    story.append(Paragraph("대리점 수수료 정산 리포트", st_["title"]))
    meta_bits = [b for b in [meta.get("month"), meta.get("hq"), f"대리점 {len(totals_sel)}곳",
                             f"생성일 {meta['created']}"] if b]
    story.append(Paragraph(escape("  |  ".join(meta_bits)), st_["meta"]))
    story.append(Spacer(1, 8))

    # 1. 정산 요약
    story.append(Paragraph("1. 대리점별 정산 요약 (원)", st_["h1"]))
    header = ["센터명", "대리점명"] + order + ["합계"]
    rows = [[r["센터명"], r["대리점명"]] + [won(r[g]) for g in order] + [won(r["합계"])]
            for _, r in totals_sel.iterrows()]
    rows.append(["합계", ""] + [won(totals_sel[g].sum()) for g in order] + [won(totals_sel["합계"].sum())])
    widths = _scale([16, 24] + [14] * len(order) + [16], usable)
    story.append(table(header, rows, widths, right=set(range(2, len(header))), total_last=True))

    # 2. 분석 리포트
    story.append(Spacer(1, 10))
    story.append(Paragraph("2. 분석 리포트", st_["h1"]))
    for title, lines in sections:
        story.append(Paragraph(escape(title), st_["h2"]))
        for ln in lines:
            indent = ln.startswith("  ")
            txt = escape(ln.strip())
            story.append(Paragraph(("&nbsp;&nbsp;&nbsp;&nbsp;" if indent else "• ") + txt, st_["body"]))

    # 3. 정산 항목 상세
    story.append(PageBreak())
    story.append(Paragraph("3. 대리점별 정산 항목 상세 (공란 제외, 원)", st_["h1"]))
    for ag in totals_sel["대리점명"]:
        sub = items[items["대리점명"] == ag]
        story.append(Paragraph(escape(ag), st_["h2"]))
        if sub.empty:
            story.append(Paragraph("표시할 정산 항목이 없습니다.", st_["body"]))
            continue
        rows = [[r["구분"], r["항목"], won(r["금액"])] for _, r in sub.iterrows()]
        rows.append(["합계", "", won(sub["금액"].sum())])
        story.append(table(["구분", "항목", "금액"], rows, _scale([14, 60, 20], usable * 0.62),
                           right={2}, total_last=True))

    # 4. 개통 상세
    if include_detail and len(detail_f):
        story.append(PageBreak())
        story.append(Paragraph("4. 개통 상세 내역 (공란 열 제외)", st_["h1"]))
        for ag in totals_sel["대리점명"]:
            d = detail_f[detail_f["대리점명"] == ag]
            if d.empty:
                continue
            cols = [c for c in DETAIL_COLUMNS if c in d.columns and d[c].notna().any()]
            if "개통일자" in d:
                d = d.sort_values("개통일자")
            story.append(Paragraph(escape(f"{ag}  ({len(d):,}건)"), st_["h2"]))
            rows = [[fmt_cell(c, v) for c, v in zip(cols, rec)] for rec in d[cols].itertuples(index=False, name=None)]
            sums = [c for c in DETAIL_SUM_COLUMNS if c in cols]
            if sums:
                tot = ["합계"] + [""] * (len(cols) - 1)
                for c in sums:
                    tot[cols.index(c)] = won(pd.to_numeric(d[c], errors="coerce").fillna(0).sum())
                rows.append(tot)
            weights = []
            for j, c in enumerate(cols):
                sample = [_vis_len(r[j]) for r in rows[:200]] or [0]
                weights.append(min(max(_vis_len(c), float(np.percentile(sample, 90))) + 2, 28))
            right = {j for j, c in enumerate(cols)
                     if pd.api.types.is_numeric_dtype(d[c]) and not any(h in c for h in NO_COMMA_HINTS)
                     and "일자" not in c and not c.endswith("일")}
            story.append(table(cols, rows, _scale(weights, usable), right=right, total_last=bool(sums)))

    def on_page(canvas, doc):
        canvas.saveState()
        canvas.setFont(font, 7.5)
        canvas.setFillColor(colors.HexColor("#777777"))
        canvas.drawString(margin, 9 * mm, f"대리점 수수료 정산 리포트  {meta.get('month', '')}")
        canvas.drawRightString(page_w - margin, 9 * mm, f"- {doc.page} -")
        canvas.restoreState()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=margin, rightMargin=margin,
                            topMargin=12 * mm, bottomMargin=15 * mm, title="대리점 수수료 정산 리포트")
    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────
# 엑셀 내보내기 (PDF에 다 담기지 않는 전체 열은 엑셀로)
# ─────────────────────────────────────────────────────────────
def build_excel(totals_sel, items, summary_f, detail_f):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        totals_sel.to_excel(xw, sheet_name="대리점별_합계", index=False)
        items.to_excel(xw, sheet_name="정산항목_상세", index=False)
        summary_f.to_excel(xw, sheet_name="1_요약(필터)", index=False)
        detail_f.to_excel(xw, sheet_name="2_인센티브(필터)", index=False)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────
# 접근 제어 (공유용)
# ─────────────────────────────────────────────────────────────
def _secret(key, default=None):
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default


def _norm_agencies(v):
    if isinstance(v, str):
        v = v.strip()
        return ["*"] if v == "*" else [v]
    return [str(x).strip() for x in v]


def get_access_config():
    """secrets에서 관리자 비밀번호(ADMIN_PASSWORD)와 사용자별 접속 코드(USERS)를 읽는다."""
    admin = _secret("ADMIN_PASSWORD")
    admin = str(admin) if admin else ""
    users = {}
    raw = _secret("USERS")
    if raw:
        try:
            for code, ags in dict(raw).items():
                users[str(code)] = _norm_agencies(ags)
        except Exception:
            users = {}
    return admin, users


def _same(a, b):
    return hmac.compare_digest(str(a).encode("utf-8"), str(b).encode("utf-8"))


def match_code(code, admin, users):
    """접속 코드로 권한을 판정한다. 관리자는 전체 대리점, 사용자는 지정된 대리점만."""
    if admin and _same(code, admin):
        return {"role": "admin", "agencies": None}
    for c, ags in users.items():
        if _same(code, c):
            return {"role": "user", "agencies": None if "*" in ags else set(ags)}
    return None


def login():
    admin, users = get_access_config()
    if not admin and not users:
        return {"role": "open", "agencies": None}  # 비밀번호 미설정: 개인용(기존 방식)
    if st.session_state.get("auth"):
        return st.session_state["auth"]
    code = st.text_input("접속 코드", type="password")
    if code:
        auth = match_code(code, admin, users)
        if auth:
            st.session_state["auth"] = auth
            st.rerun()
        else:
            st.error("접속 코드가 올바르지 않습니다.")
    return None


# ─────────────────────────────────────────────────────────────
# 공유용 파일 저장/불러오기
# ─────────────────────────────────────────────────────────────
def _shared_paths():
    return os.path.join(DATA_DIR, "latest.bin"), os.path.join(DATA_DIR, "latest.json")


def save_shared(data, filename):
    os.makedirs(DATA_DIR, exist_ok=True)
    bin_p, meta_p = _shared_paths()
    tmp = bin_p + ".tmp"
    with open(tmp, "wb") as f:
        f.write(data)
    os.replace(tmp, bin_p)
    meta = {"filename": filename, "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M")}
    with open(meta_p, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False)


def read_shared():
    """저장된 공유 파일을 (바이트, 파일명, 저장시각)으로 돌려준다. 없으면 None."""
    bin_p, meta_p = _shared_paths()
    try:
        if not (os.path.exists(bin_p) and os.path.exists(meta_p)):
            return None
        with open(meta_p, encoding="utf-8") as f:
            meta = json.load(f)
        with open(bin_p, "rb") as f:
            data = f.read()
        return data, meta.get("filename", "shared.xlsx"), meta.get("saved_at", "")
    except Exception:
        return None


# ─────────────────────────────────────────────────────────────
# 화면
# ─────────────────────────────────────────────────────────────
def show_item_detail(agency, group, item, amount, detail_f):
    """선택한 정산 항목의 개통 건별 세부내역을 화면에 보여준다."""
    st.divider()
    st.subheader(f"{agency} · {item}")
    res = item_detail(detail_f, agency, item)
    if res is None:
        st.info("이 항목은 개통 상세 시트에 같은 이름의 열이 없어 건별 내역을 표시할 수 없습니다. (요약 시트에만 있는 항목)")
        return
    out, total, n = res
    st.caption(f"구분: {group}  ·  해당 건수 {n:,}건  ·  건별 합계 {won(total)}원")
    if abs(total - amount) < 1:
        st.success(f"요약 시트의 금액({won(amount)}원)과 건별 합계가 일치합니다.")
    else:
        st.warning(f"요약 시트 금액({won(amount)}원)과 건별 합계({won(total)}원)가 다릅니다. 차이 {won(total - amount)}원")
    if n == 0:
        st.info("금액이 있는 개통 건이 없습니다. (대리점 단위로만 반영된 항목일 수 있습니다)")
        return
    st.dataframe(out, use_container_width=True, hide_index=True)
    safe = re.sub(r'[\\/:*?"<>|\[\]]', "_", f"{agency}_{item}")
    st.download_button("⬇️ 이 항목 세부내역 CSV", out.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"{safe}.csv", mime="text/csv", key=f"dl_{safe}")


def main():
    st.set_page_config(page_title="대리점 정산 에이전트", page_icon="📄", layout="wide")
    st.title("📄 대리점 수수료 정산 에이전트")

    auth = login()
    if auth is None:
        return
    role, allowed = auth["role"], auth["agencies"]
    if role != "open":
        with st.sidebar:
            st.write("권한: " + ("관리자" if role == "admin" else "사용자"))
            if st.button("로그아웃"):
                st.session_state.pop("auth", None)
                st.session_state.pop("pdf", None)
                st.rerun()

    st.caption("본사 정산마스터 파일에서 담당 대리점만 골라 PDF 표와 분석 리포트로 정리합니다. ('요약'·'인센티브' 시트만 사용)")

    # ── 파일 준비: 관리자/개인은 업로드, 사용자는 저장된 공유 파일 사용
    shared = read_shared()
    up = None
    if role in ("admin", "open"):
        title = "📤 정산마스터 파일 업로드" + (" / 공유용 교체" if role == "admin" else "")
        with st.expander(title, expanded=(shared is None)):
            up = st.file_uploader("정산마스터 파일 (xlsb / xlsx)", type=["xlsb", "xlsx"])
            if up is not None and role == "admin":
                st.caption("저장하면 접속 코드를 받은 사용자가 파일 업로드 없이 이 파일을 사용합니다.")
                if st.button("공유용으로 저장"):
                    try:
                        save_shared(up.getvalue(), up.name)
                        shared = read_shared()
                        st.success("공유용으로 저장했습니다.")
                    except Exception as e:
                        st.error(f"저장하지 못했습니다: {e}")

    if up is not None:
        data, src_name, saved_at = up.getvalue(), up.name, "방금 업로드"
    elif shared:
        data, src_name, saved_at = shared
    else:
        st.info("관리자가 아직 정산 파일을 올리지 않았습니다." if role == "user" else "파일을 올리면 시작합니다.")
        return
    st.caption(f"사용 중인 파일: {src_name} ({saved_at})")

    try:
        summary, groups, detail, sheet_names = load_workbook(data, src_name)
    except ImportError:
        st.error("xlsb를 읽으려면 pyxlsb가 필요합니다. requirements.txt에 pyxlsb가 있는지 확인해 주세요.")
        return
    except Exception as e:
        st.error(f"파일을 읽지 못했습니다: {e}")
        return

    st.success(f"읽기 완료 · 사용 시트: {' / '.join(sheet_names)} · 요약 {len(summary):,}행, 개통 상세 {len(detail):,}행")

    # ── 대리점 선택 (사용자 권한이면 허용된 대리점만 노출)
    agencies_all = list(dict.fromkeys(list(summary["대리점명"]) + list(detail["대리점명"])))
    agencies = agencies_all if allowed is None else [a for a in agencies_all if a in allowed]
    if not agencies:
        st.error("선택할 수 있는 대리점이 없습니다. 접속 코드에 지정된 대리점명이 이번 파일에 있는지 관리자에게 확인해 주세요.")
        return
    default = agencies if role == "user" else [a for a in MY_AGENCIES if a in agencies]
    c1, c2, c3 = st.columns([3, 1, 1])
    picked = c1.multiselect("담당 대리점 (대리점명 기준)", agencies, default=default)
    drop_zero = c2.checkbox("0원 항목도 제외", value=True)
    include_detail = c3.checkbox("PDF에 개통 상세 포함", value=True)
    if not picked:
        st.warning("대리점을 선택해 주세요." + (" (app.py 상단 MY_AGENCIES에 적어두면 자동 선택됩니다)" if role != "user" else ""))
        return

    totals_all, order = agency_totals(summary, groups)
    summary_f = summary[summary["대리점명"].isin(picked)]
    detail_f = detail[detail["대리점명"].isin(picked)]
    totals_sel = totals_all[totals_all["대리점명"].isin(picked)]
    items = build_items(summary_f, groups, drop_zero)
    sections = make_analysis(totals_all, totals_sel, order, items, detail_f)

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("정산 합계(원)", won(totals_sel["합계"].sum()))
    k2.metric("대리점 수", f"{len(totals_sel)}")
    k3.metric("개통 건수", f"{len(detail_f):,}")
    if "추정_인센티브(본사)" in detail_f and len(detail_f):
        k4.metric("건당 추정 인센티브(원)",
                  won(pd.to_numeric(detail_f["추정_인센티브(본사)"], errors="coerce").fillna(0).mean()))

    t1, t2, t3, t4 = st.tabs(["대리점별 요약", "분석 리포트", "정산 항목 상세", "개통 상세"])
    with t1:
        show = totals_sel.copy()
        for c in order + ["합계"]:
            show[c] = show[c].map(won)
        st.dataframe(show, use_container_width=True, hide_index=True)
    with t2:
        for title, lines in sections:
            st.subheader(title)
            for ln in lines:
                st.markdown(("&nbsp;&nbsp;&nbsp;&nbsp;" if ln.startswith("  ") else "- ") + ln.strip())
    with t3:
        ag_pick = "전체"
        if len(totals_sel) > 1:
            ag_pick = st.selectbox("대리점 보기", ["전체"] + list(totals_sel["대리점명"]))
        view = items if ag_pick == "전체" else items[items["대리점명"] == ag_pick]
        view = view.reset_index(drop=True)
        show = view.copy()
        show["금액"] = show["금액"].map(won)
        st.caption("표에서 정산 항목 행을 클릭하면 아래에 개통 건별 세부내역이 나옵니다.")
        chosen = None
        tbl_key = f"items_tbl_{abs(hash((ag_pick, tuple(picked), drop_zero, src_name)))}"
        try:
            ev = st.dataframe(show, use_container_width=True, hide_index=True,
                              on_select="rerun", selection_mode="single-row", key=tbl_key)
            rows = list(getattr(getattr(ev, "selection", None), "rows", []) or [])
            if rows and rows[0] < len(view):
                chosen = view.iloc[rows[0]]
        except TypeError:  # 행 선택을 지원하지 않는 구버전 Streamlit: 선택 상자로 대체
            st.dataframe(show, use_container_width=True, hide_index=True)
            labels = [f"{r['대리점명']} | {r['항목']}" for _, r in view.iterrows()]
            pick = st.selectbox("세부내역을 볼 항목", ["선택 안 함"] + labels)
            if pick != "선택 안 함":
                chosen = view.iloc[labels.index(pick)]
        if chosen is not None:
            show_item_detail(chosen["대리점명"], chosen["구분"], chosen["항목"], float(chosen["금액"]), detail_f)
    with t4:
        cols = [c for c in DETAIL_COLUMNS if c in detail_f.columns and detail_f[c].notna().any()]
        show = pd.DataFrame({c: [fmt_cell(c, v) for v in detail_f[c]] for c in ["대리점명"] + cols})
        st.dataframe(show, use_container_width=True, hide_index=True)

    st.divider()
    meta = {"month": month_label(src_name), "created": datetime.now().strftime("%Y-%m-%d"),
            "hq": ""}
    if "본부명" in detail_f.columns and detail_f["본부명"].notna().any():
        meta["hq"] = str(detail_f["본부명"].mode().iloc[0])

    # 선택 조건이 바뀌면 이전에 만든 PDF는 다시 만들도록 서명으로 구분
    sig = (tuple(picked), drop_zero, include_detail, src_name, saved_at)
    b1, b2 = st.columns(2)
    if b1.button("PDF 리포트 만들기", type="primary"):
        with st.spinner("PDF를 만드는 중입니다..."):
            st.session_state["pdf"] = build_pdf(meta, totals_sel, order, items, detail_f, sections, include_detail)
            st.session_state["pdf_sig"] = sig
    if st.session_state.get("pdf") and st.session_state.get("pdf_sig") == sig:
        fname = f"대리점정산_{meta['month'] or datetime.now().strftime('%Y%m')}.pdf".replace(" ", "")
        b1.download_button("⬇️ PDF 내려받기", st.session_state["pdf"], file_name=fname, mime="application/pdf")

    sf = drop_empty_columns(summary_f, drop_zero)
    df_ = drop_empty_columns(detail_f, drop_zero=False)
    b2.download_button("⬇️ 필터링된 엑셀 내려받기 (전체 열)", build_excel(totals_sel, items, sf, df_),
                       file_name="대리점정산_필터결과.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("※ 고객 정보가 포함되어 있으니 파일 공유·보관에 유의하세요.")


if __name__ == "__main__":
    main()
