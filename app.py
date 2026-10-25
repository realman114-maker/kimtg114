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

st.title("📊 SK브로드밴드 대리점 정산 및 PDF 리포트 에이전트")
st.markdown("본사 정산 파일(`.xlsb`)을 업로드하고 관리하시는 **대리점명**을 입력하시면, 개인정보 보호를 위해 타 대리점 정보를 완벽히 차단하고 지정된 대리점의 **'요약' 및 '인센티브' 데이터**만 추출하여 PDF 리포트를 생성합니다.")

# 사용자 입력 영역
col1, col2 = st.columns(2)
with col1:
    uploaded_file = st.file_uploader("📁 본사 정산 파일(.xlsb) 업로드", type=["xlsb"])
with col2:
    target_agency = st.text_input("🏢 조회할 대리점명 입력", placeholder="예: 울산중앙대리점")

if uploaded_file is not None and target_agency:
    if st.button("🚀 보안 필터링 및 PDF 리포트 생성", type="primary"):
        with st.spinner("개인정보 보호 필터링 및 '요약'/'인센티브' 시트 정밀 분석 중..."):
            try:
                # 1. pyxlsb를 이용해 xlsb 파일 읽기
                with open_workbook(uploaded_file) as wb:
                    sheet_names = wb.sheets
                    st.info(f"📋 파일 내 감지된 시트 목록: {sheet_names}")

                    # 지정된 시트명("요약", "인센티브") 또는 순서에 맞춰 시트 찾기
                    summary_sheet_name = None
                    incentive_sheet_name = None

                    for name in sheet_names:
                        if "요약" in name:
                            summary_sheet_name = name
                        elif "인센티브" in name:
                            incentive_sheet_name = name

                    # 만약 이름 매칭이 안 되면 기본 첫 번째, 두 번째 시트 사용
                    if not summary_sheet_name and len(sheet_names) > 0:
                        summary_sheet_name = sheet_names[0]
                    if not incentive_sheet_name and len(sheet_names) > 1:
                        incentive_sheet_name = sheet_names[1]

                    # 첫 번째 시트 (요약) 읽기
                    data_summary = []
                    if summary_sheet_name:
                        with wb.get_sheet(summary_sheet_name) as sheet:
                            for row in sheet.rows():
                                data_summary.append([item.v for item in row])

                    # 두 번째 시트 (인센티브) 읽기
                    data_incentive = []
                    if incentive_sheet_name:
                        with wb.get_sheet(incentive_sheet_name) as sheet:
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
                    # 2. 타 대리점 데이터 완벽 차단 및 해당 대리점만 필터링
                    filtered_summary = df_summary[df_summary[agency_col].astype(str).str.strip() == target_agency.strip()]

                    # [선택 최적화] 불필요한 내부 컬럼 제외하고 '수수료 정산 금액' 및 주요 지표 중심 컬럼만 남기기 (필요시 키워드 필터링)
                    # 예시: 특정 키워드가 포함된 컬럼만 추리고 싶다면 아래 주석을 해제하여 활용 가능
                    # target_keywords = ["대리점", "수수료", "금액", "정산", "요금", "건수", "인센티브"]
                    # keep_cols = [c for c in filtered_summary.columns if any(k in str(c) for k in target_keywords)]
                    # if keep_cols: filtered_summary = filtered_summary[keep_cols]

                    st.success(f"🔒 보안 필터링 완료: 타 대리점 정보 차단 및 '{target_agency}' 데이터 추출 성공")

                    # 화면 표시: 정산 요약
                    st.subheader(f"📌 [{target_agency}] 정산 요약 결과 (시트: {summary_sheet_name})")
                    if not filtered_summary.empty:
                        st.dataframe(filtered_summary, use_container_width=True)
                    else:
                        st.warning(f"'{target_agency}'에 해당하는 요약 데이터를 찾지 못했습니다. 대리점명을 다시 확인해주세요.")

                    # 인센티브 시트 필터링
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
                                st.subheader(f"🎁 [{target_agency}] 인센티브 세부 내역 (시트: {incentive_sheet_name})")
                                st.dataframe(filtered_incentive, use_container_width=True)

                    # --- 3. 대리점 배포용 PDF 리포트 생성 로직 ---
                    if not filtered_summary.empty or not filtered_incentive.empty:
                        buffer = io.BytesIO()
                        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=30, bottomMargin=30)
                        elements = []

                        styles = getSampleStyleSheet()
                        title_style = ParagraphStyle('ReportTitle', parent=styles['Heading1'], fontName=DEFAULT_FONT, fontSize=15, leading=19, textColor=colors.HexColor("#003366"))
                        heading_style = ParagraphStyle('SectionHeading', parent=styles['Heading2'], fontName=DEFAULT_FONT, fontSize=11, leading=14, spaceBefore=10, spaceAfter=4, textColor=colors.HexColor("#333333"))
                        cell_style = ParagraphStyle('TableCell', parent=styles['Normal'], fontName=DEFAULT_FONT, fontSize=7.5, leading=9.5)

                        # 문서 타이틀
                        elements.append(Paragraph(f"SK브로드밴드 대리점 수수료 정산 리포트", title_style))
                        elements.append(Paragraph(f"<b>대상 대리점:</b> {target_agency}", ParagraphStyle('SubTitle', parent=styles['Normal'], fontName=DEFAULT_FONT, fontSize=9.5, leading=13, spaceAfter=8)))
                        elements.append(Spacer(1, 4))

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
                                ('BOTTOMPADDING', (0,0), (-1,0), 4),
                                ('TOPPADDING', (0,0), (-1,0), 4),
                                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CCCCCC"))
                            ]))
                            elements.append(t_sum)
                            elements.append(Spacer(1, 8))

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
                                ('BOTTOMPADDING', (0,0), (-1,0), 4),
                                ('TOPPADDING', (0,0), (-1,0), 4),
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
                    st.error("요약 시트에서 '대리점' 관련 컬럼을 찾을 수 없습니다. 컬럼명을 확인해 주세요.")

            except Exception as e:
                st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")
