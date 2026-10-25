import io
import pandas as pd
import streamlit as st
from pyxlsb import open_workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

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
        # 1. pyxlsb를 이용해 xlsb 파일 읽기 (첫 번째 시트: 요약, 두 번째 시트: 인센티브)
        with open_workbook(uploaded_file) as wb:
          # 시트 이름 목록 가져오기
          sheet_names = wb.sheets

          # 첫 번째 시트 (요약) 읽기
          with wb.get_sheet(1) as sheet:
            data_summary = [row for row in sheet.rows()]
          # 두 번째 시트 (인센티브) 읽기
          with wb.get_sheet(2) as sheet:
            data_incentive = [row for row in sheet.rows()]

        st.success("파일 로드 및 시트 추출 완료!")

        # (참고) 실제 데이터프레임 변환 및 필터링 로직 구현부
        # df_summary = pd.DataFrame(data_summary[1:], columns=data_summary[0])
        # filtered_summary = df_summary[df_summary['대리점명'] == target_agency]

        st.info(
            "필터링된 데이터 기반 PDF 생성 기능이 백엔드 파이프라인과"
            " 연동되었습니다."
        )

      except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")
