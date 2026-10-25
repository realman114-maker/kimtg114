import streamlit as st
import pandas as pd
from pyxlsb import open_workbook
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
import io

# 페이지 설정
st.set_page_config(page_title="대리점 수수료 정산 프로그램", layout="wide")

st.title("📊 대리점 수수료 정산 및 PDF 리포트 생성기")
st.write("본사 정산 파일(.xlsb)을 업로드하고 대리점명을 선택하여 요약 및 PDF 리포트를 생성하세요.")

# 1. 파일 업로드
uploaded_file = st.file_uploader("본사 수수료 정산 파일(.xlsb)을 업로드하세요", type=["xlsb"])

if uploaded_file is not None:
    try:
        # xlsb 읽기 (첫 번째 시트: 요약, 두 번째 시트: 인센티브)
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

        # 대리점명 컬럼 유연하게 찾기 (대리점명, 대리점, 점포명, 업체명 등)
        agency_col = None
        possible_keywords = ["대리점명", "대리점", "점포명", "업체명", "거래처"]
       
        for col in df_summary.columns:
            col_str = str(col).strip()
            for kw in possible_keywords:
                if kw in col_str:
                    agency_col = col
                    break
            if agency_col:
                break

        if agency_col:
            # 대리점 목록 추출
            agency_list = df_summary[agency_col].dropna().astype(str).str.strip().unique().tolist()
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
                        for kw in possible_keywords:
                            if kw in col_str:
                                inc_agency_col = col
                                break
                        if inc_agency_col:
                            break
                    if inc_agency_col:
                        filtered_incentive = df_incentive[df_incentive[inc_agency_col].astype(str).str.strip() == selected_agency]
                        st.subheader(f"🎁 [{selected_agency}] 인센티브 세부 내역")
                        st.dataframe(filtered_incentive)

                # PDF 다운로드 버튼 생성
                buffer = io.BytesIO()
                doc = SimpleDocTemplate(buffer, pagesize=A4)
                elements = []

                styles = getSampleStyleSheet()
                title_style = styles['Heading1']
               
                elements.append(Paragraph(f"Agency Settlement Report: {selected_agency}", title_style))
                elements.append(Spacer(1, 12))

                if not filtered_summary.empty:
                    table_data = [filtered_summary.columns.astype(str).tolist()] + filtered_summary.astype(str).values.tolist()
                    t = Table(table_data)
                    t.setStyle(TableStyle([
                        ('BACKGROUND', (0,0), (-1,0), colors.grey),
                        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                        ('BOTTOMPADDING', (0,0), (-1,0), 6),
                        ('GRID', (0,0), (-1,-1), 1, colors.black)
                    ]))
                    elements.append(t)

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
            st.error("파일에서 '대리점명' 또는 대리점 관련 컬럼을 찾을 수 없습니다.")
            st.write("엑셀 파일에 있는 전체 컬럼 목록:", list(df_summary.columns))

    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")
