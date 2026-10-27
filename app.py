import io
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

# PDF '개통 상세' 표에 보여줄 열 (2번째 시트 기준, 없는 열/전부 공란인 열은 자동 제외)
DETAIL_COLUMNS = [
    "개통일자", "서비스관리번호", "고객명", "요금상품", "약정할인개월명",
    "번들유형구분명", "시군구명", "사은품합계금액", "추정_인센티브(본사)", "해지일",
]
# 개통 상세 표 맨 아래에 합계를 낼 열
DETAIL_SUM_COLUMNS = ["사은품합계금액", "추정_인센티브(본사)"]

ID_COLS = {"센터명", "대리점명", "合", "합", "합계"}
NO_COMMA_HINTS = ("코드", "번호", "ID", "id")


# ─────────────────────────────────────────────────────────────
# 파일 읽기
# ─────────────────────────────────────────────────────────────
def _clean(name):
    return re.sub(r"\s+", " ", str(name)).strip()


def _engine(filename):
    return "pyxlsb" if filename.lower().endswith(".xlsb") else None


def _parse_sheet(xl, idx):
    """시트를 읽어 '대리점명' 헤더 행을 자동으로 찾고 (DataFrame, 그룹 라벨 dict)를 돌려준다."""
    raw = xl.parse(sheet_name=idx, header=None)
    hdr = None
    for i in range(min(20, len(raw))):
        vals = [_clean(v) for v in raw.iloc[i].tolist() if pd.notna(v)]
        if "대리점명" in vals:
            hdr = i
            break
    if hdr is None:
        raise ValueError(f"{idx + 1}번째 시트에서 '대리점명' 헤더를 찾지 못했습니다.")

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
    sheet_names = xl.sheet_names[:2]
    summary, groups = _parse_sheet(xl, 0)
    detail, _ = _parse_sheet(xl, 1)

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
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
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
# 화면
# ─────────────────────────────────────────────────────────────
def check_password():
    try:
        pw = st.secrets.get("APP_PASSWORD")
    except Exception:
        pw = None
    if not pw:
        return True
    if st.session_state.get("authed"):
        return True
    typed = st.text_input("비밀번호", type="password")
    if typed:
        if typed == pw:
            st.session_state["authed"] = True
            st.rerun()
        else:
            st.error("비밀번호가 올바르지 않습니다.")
    return False


def main():
    st.set_page_config(page_title="대리점 정산 에이전트", page_icon="📄", layout="wide")
    st.title("📄 대리점 수수료 정산 에이전트")
    if not check_password():
        return
    st.caption("본사 정산마스터 파일에서 담당 대리점만 골라 PDF 표와 분석 리포트로 정리합니다. (1·2번째 시트만 사용)")

    up = st.file_uploader("정산마스터 파일 업로드 (xlsb / xlsx)", type=["xlsb", "xlsx"])
    if up is None:
        st.info("파일을 올리면 시작합니다.")
        return

    try:
        summary, groups, detail, sheet_names = load_workbook(up.getvalue(), up.name)
    except ImportError:
        st.error("xlsb를 읽으려면 pyxlsb가 필요합니다. requirements.txt에 pyxlsb가 있는지 확인해 주세요.")
        return
    except Exception as e:
        st.error(f"파일을 읽지 못했습니다: {e}")
        return

    st.success(f"읽기 완료 · 사용 시트: {' / '.join(sheet_names)} · 요약 {len(summary):,}행, 개통 상세 {len(detail):,}행")

    agencies = list(dict.fromkeys(list(summary["대리점명"]) + list(detail["대리점명"])))
    default = [a for a in MY_AGENCIES if a in agencies]
    c1, c2, c3 = st.columns([3, 1, 1])
    picked = c1.multiselect("담당 대리점 (대리점명 기준)", agencies, default=default)
    drop_zero = c2.checkbox("0원 항목도 제외", value=True)
    include_detail = c3.checkbox("PDF에 개통 상세 포함", value=True)
    if not picked:
        st.warning("담당 대리점을 선택해 주세요. (app.py 상단 MY_AGENCIES에 적어두면 자동 선택됩니다)")
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
        show = items.copy()
        show["금액"] = show["금액"].map(won)
        st.dataframe(show, use_container_width=True, hide_index=True)
    with t4:
        cols = [c for c in DETAIL_COLUMNS if c in detail_f.columns and detail_f[c].notna().any()]
        show = pd.DataFrame({c: [fmt_cell(c, v) for v in detail_f[c]] for c in ["대리점명"] + cols})
        st.dataframe(show, use_container_width=True, hide_index=True)

    st.divider()
    meta = {"month": month_label(up.name), "created": datetime.now().strftime("%Y-%m-%d"),
            "hq": ""}
    if "본부명" in detail_f.columns and detail_f["본부명"].notna().any():
        meta["hq"] = str(detail_f["본부명"].mode().iloc[0])

    b1, b2 = st.columns(2)
    if b1.button("PDF 리포트 만들기", type="primary"):
        with st.spinner("PDF를 만드는 중입니다..."):
            st.session_state["pdf"] = build_pdf(meta, totals_sel, order, items, detail_f, sections, include_detail)
    if st.session_state.get("pdf"):
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
