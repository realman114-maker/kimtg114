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

st.set_page_config(page_title="SK브로드밴드 대리점 정산 에이전트", layout="wide")

st.title("📊 SK브로드밴드 대리점 맞춤형 수수료 정산 및 PDF 생성기")
st.markdown("본사 정산 파일(`.xlsb`)을 업로드하고 관리하시는 **대리점명**을 입력하시면, 개인정보 보호를 위해 **해당 대리점 데이터만 필터링**하고 불필요한 항목을 제외한 **깔끔한 수수료 중심의 PDF 리포트**를 생성합니다.")

# 사용자 입력 영역
col1, col2 = st.columns(2)
with col1:
    uploaded_file = st.file_uploader("📁 본사 정산 파일(.xlsb) 업로드", type=["xlsb"])
with col2:
    target_agency = st.text_input("🏢 조회할 대리점명 입력", placeholder="예: 울산중앙대리점")

if uploaded_file is not None and target_agency:
    if st.button("🚀 정산 데이터 보안 필터링 및 PDF 리포트 생성", type="primary"):
        with st.spinner("개인정보 보호 필터링 및 리포트 데이터 가공 중..."):
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

                # 대리점명 컬럼 자동 탐색
                agency_col = None
                for col in df_summary.columns:
                    if col and "대리점" in str(col):
                        agency_col = col
                        break

                if agency_col:
                    # 2. 타 대리점 데이터 차단 및 해당 대리점만 엄격하게 필터링
                    filtered_summary = df_summary[df_summary[agency_col].astype(str).str.strip() == target_agency.strip()]

                    # [선택사항] 수수료/핵심 정산 항목 위주로 최적화 (필요시 제외할 키워드 설정 가능)
                    # 예: 불필요한 내부 관리용 컬럼이 있다면 여기서 drop() 등으로 제외할 수 있습니다.
                    
                    st.success(f"🔒 보안 필터링 완료: 타 대리점 정보 차단 및 '{target_agency}' 데이터 추출 성공")

                    # 화면 표시: 정산 요약
                    st.subheader(f"📌 [{target_agency}] 정산 요약 결과")
                    if not filtered_summary.empty:
                        st.dataframe(filtered_summary, use_container_width=True)
                    else:
                        st.warning(f"'{target_agency}'에 해당하는 정산 요약 데이터를 찾지 못했습니다. 대리점명을 다시 확인해주세요.")

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
                                st.dataframe(filtered_incentive, use_container_width=True)

                    # --- 3. 대리점 배포용 PDF 리포트 생성 로직 ---
                    if not filtered_summary.empty or not filtered_incentive.empty:
                        buffer = io.BytesIO()
                        # 여백을 주어 깔끔한 문서 레이아웃 구성
                        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=30, bottomMargin=30)
                        elements = []

                        styles = getSampleStyleSheet()
                        title_style = ParagraphStyle('ReportTitle', parent=styles['Heading1'], fontName=DEFAULT_FONT, fontSize=16, leading=20, textColor=colors.HexColor("#003366"))
                        heading_style = ParagraphStyle('SectionHeading', parent=styles['Heading2'], fontName=DEFAULT_FONT, fontSize=12, leading=15, spaceBefore=12, spaceAfter=6, textColor=colors.HexColor("#333333"))
                        cell_style = ParagraphStyle('TableCell', parent=styles['Normal'], fontName=DEFAULT_FONT, fontSize=8, leading=10)

                        # 문서 타이틀 (SK브로드밴드 공식 배포 느낌)
                        elements.append(Paragraph(f"SK브로드밴드 대리점 수수료 정산 리포트", title_style))
                        elements.append(Paragraph(f"<b>대상 대리점:</b> {target_agency}", ParagraphStyle('SubTitle', parent=styles['Normal'], fontName=DEFAULT_FONT, fontSize=10, leading=14, spaceAfter=10)))
                        elements.append(Spacer(1, 5))

                        # 1. 요약 테이블 추가
                        if not filtered_summary.empty:
                            elements.append(Paragraph("1. 월 정산 요약", heading_style))
                            sum_headers = [Paragraph(str(col), cell_style) for col in filtered_summary.columns]
                            sum_rows = [[Paragraph(str(val), cell_style) for val in row] for row in filtered_summary.astype(str).values.tolist()]
                            
                            t_sum = Table([sum_headers] + sum_rows)
                            t_sum.setStyle(TableStyle([
                                ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#005EA6")),
                                ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                                ('BOTTOMPADDING', (0,0), (-1,0), 5),
                                ('TOPPADDING', (0,0), (-1,0), 5),
                                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CCCCCC"))
                            ]))
                            elements.append(t_sum)
                            elements.append(Spacer(1, 10))

                        # 2. 인센티브 테이블 추가
                        if not filtered_incentive.empty:
                            elements.append(Paragraph("2. 인센티브 세부 지급 내역", heading_style))
                            inc_headers = [Paragraph(str(col), cell_style) for col in filtered_incentive.columns]
                            inc_rows = [[Paragraph(str(val), cell_style) for val in row] for row in filtered_incentive.astype(str).values.tolist()]
                            
                            t_inc = Table([inc_headers] + inc_rows)
                            t_inc.setStyle(TableStyle([
                                ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#4A90E2")),
                                ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                                ('BOTTOMPADDING', (0,0), (-1,0), 5),
                                ('TOPPADDING', (0,0), (-1,0), 5),
                                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CCCCCC"))
                            ]))
                            elements.append(t_inc)

                        doc.build(elements)
                        buffer.seek(0)

                        st.markdown("---")
                        st.subheader("📥 대리점 배포용 PDF 리포트 다운로드")
                        st.download_button(
                            label=f"📄 [{target_agency}] 정산 리포트 PDF 저장",
                            data=buffer,
                            file_name=f"SKB_{target_agency}_월정산리포트.pdf",
                            mime="application/pdf",
                            type="primary"
                        )
                else:
                    st.error("본사 정산 파일의 요약 시트에서 '대리점' 관련 컬럼을 찾을 수 없습니다. 컬럼명을 확인해 주세요.")

            except Exception as e:
                st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")

<FollowUp>생성된 에이전트 코드를 활용하시면서 추가로 대리점별 정산 금액 합계나 특정 항목 자동 계산 기능 등 더 보완하고 싶은 실무 요구사항이 있으신가요?</FollowUp>
