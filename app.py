Python
import streamlit as st
import pandas as pd
from pyxlsb import open_workbook
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io
import os

# --- 한글 폰트 등록 ---
try:
    font_path = "C:/Windows/Fonts/malgun.ttf"
    if os.path.exists(font_path):
        pdfmetrics.registerFont(TTFont('MalgunGothic', font_path))
        DEFAULT_FONT = 'MalgunGothic'
    else:
        DEFAULT_FONT = 'Helvetica'
except Exception:
    DEFAULT_FONT = 'Helvetica'

# 페이지 설정
st.set_page_config(page_title="대리점 수수료 정산 프로그램", layout="wide")

st.title("📊 대리점 수수료 정산 및 PDF 리포트 생성기")
st.write("본사 정산 파일(.xlsb)을 업로드하고 대리점명을 선택하여 요약 및 PDF 리포트를 생성하세요.")

# 1. 파일 업로드
uploaded_file = st.file_uploader("본사 수수료 정산 파일(.xlsb)을 업로드하세요", type=["xlsb"])

if uploaded_file is not None:
    try:
        # xlsb 읽기
        with open_workbook(uploaded_file) as wb:
            sheets = wb.sheets
            
            # 첫 번째 시트 읽기
            with wb.get_sheet(sheets[0]) as sheet1:
                data1 = []
                for row in sheet1.rows():
                    data1.append([item.v for item in row])
                df_summary = pd.DataFrame(data1[1:], columns=data1[0]) if len(data1) > 1 else pd.DataFrame()

            # 두 번째 시트 읽기 (존재 시)
            df_incentive = pd.DataFrame()
            if len(sheets) > 1:
                with wb.get_sheet(sheets[1]) as sheet2:
                    data2 = []
                    for row in sheet2.rows():
                        data2.append([item.v for item in row])
                    df_incentive = pd.DataFrame(data2[1:], columns=data2[0]) if len(data2) > 1 else pd.DataFrame()

        st.success("파일을 성공적으로 불러왔습니다!")

        # 디버깅용: 읽어온 컬럼 목록과 데이터 미리보기 표시
        with st.expander("🔍 파일 구조 확인하기 (디버그용)", expanded=False):
            st.write("시트 이름 목록:", sheets)
            st.write("요약 시트 컬럼 목록:", df_summary.columns.tolist())
            st.dataframe(df_summary.head(3))

        # 대리점명 컬럼 찾기 (검색 키워드 확장: 대리점, 지점, 업체, 상호 등)
        agency_col = None
        keywords = ["대리점", "지점", "업체", "상호", "파트너"]
        
        for col in df_summary.columns:
            col_str = str(col).strip()
            if any(kw in col_str for kw in keywords):
                agency_col = col
                break

        # 자동 찾기 실패 시 수동 선택 옵션 제공
        if not agency_col and not df_summary.empty:
            st.warning("자동으로 '대리점' 컬럼을 찾지 못했습니다. 아래에서 직접 대리점 컬럼을 선택해 주세요.")
            agency_col = st.selectbox("대리점 컬럼 선택:", df_summary.columns.tolist())

        if agency_col:
            # 대리점 목록 추출 (결측치 제거 및 문자열 변환)
            agency_list = df_summary[agency_col].dropna().astype(str).str.strip().unique().tolist()
            # 빈 값이나 'nan' 제거
            agency_list = [a for a in agency_list if a and a.lower() != 'nan']

            if len(agency_list) == 0:
                st.error(f"'{agency_col}' 컬럼에 유효한 대리점 데이터가 없습니다.")
            else:
                selected_agency = st.selectbox("조회할 대리점을 선택하세요:", agency_list)

                if selected_agency:
                    # 대리점 데이터 필터링
                    filtered_summary = df_summary[df_summary[agency_col].astype(str).str.strip() == selected_agency]
                    
                    st.subheader(f"📌 [{selected_agency}] 정산 요약")
                    st.dataframe(filtered_summary)

                    filtered_incentive = pd.DataFrame()
                    if not df_incentive.empty:
                        inc_agency_col = None
                        for col in df_incentive.columns:
                            col_str = str(col).strip()
                            if any(kw in col_str for kw in keywords):
                                inc_agency_col = col
                                break
                        if inc_agency_col:
                            filtered_incentive = df_incentive[df_incentive[inc_agency_col].astype(str).str.strip() == selected_agency]
                            st.subheader(f"🎁 [{selected_agency}] 인센티브 세부 내역")
                            st.dataframe(filtered_incentive)

                    # PDF 다운로드 버튼 생성
                    buffer = io.BytesIO()
                    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
                    elements = []

                    styles = getSampleStyleSheet()
                    
                    title_style = ParagraphStyle('ReportTitle', parent=styles['Heading1'], fontName=DEFAULT_FONT, fontSize=18, leading=22)
                    heading_style = ParagraphStyle('SectionHeading', parent=styles['Heading2'], fontName=DEFAULT_FONT, fontSize=14, leading=18, spaceBefore=10, spaceAfter=6)
                    cell_style = ParagraphStyle('TableCell', parent=styles['Normal'], fontName=DEFAULT_FONT, fontSize=9, leading=11)

                    elements.append(Paragraph(f"대리점 정산 리포트: {selected_agency}", title_style))
                    elements.append(Spacer(1, 12))

                    if not filtered_summary.empty:
                        elements.append(Paragraph("1. 정산 요약", heading_style))
                        summary_headers = [Paragraph(str(col), cell_style) for col in filtered_summary.columns]
                        summary_rows = [[Paragraph(str(val), cell_style) for val in row] for row in filtered_summary.astype(str).values.tolist()]
                        
                        t = Table([summary_headers] + summary_rows)
                        t.setStyle(TableStyle([
                            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#4A90E2")),
                            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                            ('BOTTOMPADDING', (0,0), (-1,0), 6),
                            ('GRID', (0,0), (-1,-1), 0.5, colors.grey)
                        ]))
                        elements.append(t)
                        elements.append(Spacer(1, 15))

                    if not filtered_incentive.empty:
                        elements.append(Paragraph("2. 인센티브 세부 내역", heading_style))
                        inc_headers = [Paragraph(str(col), cell_style) for col in filtered_incentive.columns]
                        inc_rows = [[Paragraph(str(val), cell_style) for val in row] for row in filtered_incentive.astype(str).values.tolist()]
                        
                        t_inc = Table([inc_headers] + inc_rows)
                        t_inc.setStyle(TableStyle([
                            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#50E3C2")),
                            ('TEXTCOLOR', (0,0), (-1,0), colors.black),
                            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                            ('BOTTOMPADDING', (0,0), (-1,0), 6),
                            ('GRID', (0,0), (-1,-1), 0.5, colors.grey)
                        ]))
                        elements.append(t_inc)

                    doc.build(elements)
                    buffer.seek(0)

                    st.markdown("---")
                    st.subheader("📥 리포트 다운로드")
                    st.download_button(
                        label="📄 PDF 리포트 다운로드",
                        data=buffer,
                        file_name=f"{selected_agency}_정산리포트.pdf",
                        mime="application/pdf"
                    )
        else:
            st.error("파일에서 '대리점' 관련 컬럼을 찾을 수 없습니다. 위의 '파일 구조 확인하기'를 열어 컬럼명을 확인해 주세요.")

    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")
