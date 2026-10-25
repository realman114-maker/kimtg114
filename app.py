import io
import os
import pandas as pd
import streamlit as st
from pyxlsb import open_workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

# --- 한글 폰트 등록 (Windows 기준 맑은 고딕, 환경에 따라 수정 가능) ---
try:
    font_path = "C:/Windows/Fonts/malgun.ttf"
    if os.path.exists(font_path):
        pdfmetrics.registerFont(TTFont('MalgunGothic', font_path))
        DEFAULT_FONT = 'MalgunGothic'
    else:
        DEFAULT_FONT = 'Helvetica'
except Exception:
    DEFAULT_FONT = 'Helvetica'

st.set_page_config(page_title="대리점 수수료 정산 자동화", layout="wide")

st.title("📊 대리점 수수료 정산 및 PDF 리포트 생성기")
st.markdown("본사 정산 파일(`.xlsb`)을 업로드하고 대리점명을 입력하면 요약과 인센티브 시트를 필터링하여 PDF로 생성합니다.")

# 사용자 입력 영역
uploaded_file = st.file_uploader("정산 파일(.xlsb) 업로드", type=["xlsb"])
target_agency = st.text_input("조회할 대리점명 입력", placeholder="예: 울산중앙대리점")

if uploaded_file is not None and target_agency:
    if st.button("정산 데이터 필터링 및 PDF 생성"):
        with st.spinner("데이터를 처리하고 있습니다..."):
            try:
                # 1. pyxlsb를 이용해 xlsb 파일 읽기
                with open_workbook(uploaded_file) as wb:
                    sheet_names = wb.sheets

                    # 첫 번째 시트 (요약) 읽기
                    data_summary = []
                    if len(sheet_names) > 0:
                        with wb.get_sheet(sheet_names[0]) as sheet:
                            for row in sheet.rows():
                                data_summary.append([item.v for item in row])

                    # 두 번째 시트 (인센티브) 읽기
                    data_incentive = []
                    if len(sheet_names) > 1:
                        with wb.get_sheet(sheet_names[1]) as sheet:
                            for row in sheet.rows():
                                data_incentive.append([item.v for item in row])

                # DataFrame 변환
                df_summary = pd.DataFrame(data_summary[1:], columns=data_summary[0]) if len(data_summary) > 1 else pd.DataFrame()
                df_incentive = pd.DataFrame(data_incentive[1:], columns=data_incentive[0]) if len(data_incentive) > 1 else pd.DataFrame()

                st.success("파일 로드 및 시트 추출 완료!")

                # 대리점명 컬럼 자동 탐색
                agency_col = None
                for col in df_summary.columns:
                    if col and "대리점" in str(col):
                        agency_col = col
                        break

                if agency_col:
                    # 입력한 대리점명과 일치하는 데이터 필터링 (공백 제거 후 비교)
                    filtered_summary = df_summary[df_summary[agency_col].astype(str).str.strip() == target_agency.strip()]

                    st.subheader(f"📌 [{target_agency}] 정산 요약 결과")
                    if not filtered_summary.empty:
                        st.dataframe(filtered_summary)
                    else:
                        st.warning(f"'{target_agency}'에 해당하는 정산 요약 데이터를 찾지 못했습니다.")

                    # 인센티브 시트 필터링 (존재할 경우)
                    filtered_incentive = pd.DataFrame()
                    if not df_incentive.empty:
                        inc_agency_col = None
                        for col in df_incentive.columns:
                            if col and "대리점" in str(col):
                                inc_agency_col = col
                                break
                        if inc_agency_col:
                            filtered_incentive = df_incentive[df_incentive[inc_agency_col].astype(str).str.strip() == target_agency.strip()]
                            if not filtered_incentive.empty:
                                st.subheader(f"🎁 [{target_agency}] 인센티브 세부 내역")
                                st.dataframe(filtered_incentive)

                    # --- PDF 생성 및 다운로드 버튼 로직 ---
                    if not filtered_summary.empty or not filtered_incentive.empty:
                        buffer = io.BytesIO()
                        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
                        elements = []

                        styles = getSampleStyleSheet()
                        title_style = ParagraphStyle('ReportTitle', parent=styles['Heading1'], fontName=DEFAULT_FONT, fontSize=18, leading=22)
                        heading_style = ParagraphStyle('SectionHeading', parent=styles['Heading2'], fontName=DEFAULT_FONT, fontSize=13, leading=16, spaceBefore=10, spaceAfter=6)
                        cell_style = ParagraphStyle('TableCell', parent=styles['Normal'], fontName=DEFAULT_FONT, fontSize=8, leading=10)

                        # 문서 제목
                        elements.append(Paragraph(f"대리점 수수료 정산 리포트: {target_agency}", title_style))
                        elements.append(Spacer(1, 10))

                        # 1. 요약 테이블 추가
                        if not filtered_summary.empty:
                            elements.append(Paragraph("1. 정산 요약", heading_style))
                            sum_headers = [Paragraph(str(col), cell_style) for col in filtered_summary.columns]
                            sum_rows = [[Paragraph(str(val), cell_style) for val in row] for row in filtered_summary.astype(str).values.tolist()]
                            
                            t_sum = Table([sum_headers] + sum_rows)
                            t_sum.setStyle(TableStyle([
                                ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#4A90E2")),
                                ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                                ('BOTTOMPADDING', (0,0), (-1,0), 5),
                                ('GRID', (0,0), (-1,-1), 0.5, colors.grey)
                            ]))
                            elements.append(t_sum)
                            elements.append(Spacer(1, 10))

                        # 2. 인센티브 테이블 추가
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
                                ('BOTTOMPADDING', (0,0), (-1,0), 5),
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
                            file_name=f"{target_agency}_정산리포트.pdf",
                            mime="application/pdf"
                        )
                else:
                    st.error("요약 시트에서 '대리점' 관련 컬럼을 찾을 수 없습니다.")

            except Exception as e:
                st.error(f"오류가 발생했습니다: {e}")
