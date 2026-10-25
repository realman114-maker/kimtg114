import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def process_agency_commission(file_path, agency_name_or_code, output_path="filtered_agency_commission.xlsx"):
    """
    본사 정산 Raw 파일에서 특정 대리점 데이터를 필터링하고,
    비해당 시트 삭제 및 새 엑셀 파일 저장을 수행하는 함수
    """
    print(f"📂 정산 파일 로딩 중: {file_path}")
    
    # 1. 파일 확장자에 따른 엔진 선택 (.xlsb 또는 .xlsx)
    ext = os.path.splitext(file_path)[1].lower()
    engine = 'pyxlsb' if ext == '.xlsb' else 'openpyxl'
    
    try:
        # 전체 시트 이름 목록 가져오기
        xls = pd.ExcelFile(file_path, engine=engine)
        sheet_names = xls.sheet_names
    except Exception as e:
        print(f"❌ 파일 읽기 오류: {e}")
        return
    
    filtered_sheets = {}
    target_keyword = str(agency_name_or_code).strip()
    
    print(f"🔍 '{target_keyword}' 대리점 데이터 필터링 시작...")
    
    for sheet in sheet_names:
        try:
            # 시트 데이터 로드
            df = pd.read_excel(file_path, sheet_name=sheet, engine=engine)
            
            if df.empty:
                continue
                
            # 대리점명/코드가 포함된 열(Column) 찾기
            target_col = None
            for col in df.columns:
                # 컬럼 내용 중 대리점 관련 키워드가 있거나, 데이터에 해당 대리점명이 포함된 경우
                matched = df[col].astype(str).str.contains(target_keyword, na=False, case=False)
                if matched.any():
                    target_col = col
                    break
            
            if target_col is not None:
                # 해당 대리점 행만 추출
                filtered_df = df[df[target_col].astype(str).str.contains(target_keyword, na=False, case=False)].copy()
                
                if not filtered_df.empty:
                    filtered_sheets[sheet] = filtered_df
                    print(f"   ✔️ [시트 유지] '{sheet}' (매칭 행 수: {len(filtered_df)})")
            else:
                print(f"   ❌ [시트 제외] '{sheet}' (해당 대리점 데이터 없음)")
                
        except Exception as e:
            print(f"   ⚠️ 시트 '{sheet}' 처리 중 예외 발생: {e}")
            continue
            
    if not filtered_sheets:
        print("❌ 일치하는 대리점 데이터가 포함된 시트나 행을 찾지 못했습니다.")
        return

    # 2. 새 엑셀 파일로 저장 및 서식 적용
    print(f"💾 새 엑셀 파일 생성 중: {output_path}")
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for sheet_name, df in filtered_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
    # 3. 열 너비 자동 조절 및 깔끔한 서식 적용 (openpyxl 활용)
    wb = openpyxl_format_adjustment(output_path)
    print(f"✨ 정산 처리 완료! 저장된 파일: {output_path}")
    return output_path

def openpyxl_format_adjustment(file_path):
    """생성된 엑셀 파일의 가독성을 높이기 위한 서식 정리 (열 너비 자동 맞춤 등)"""
    import openpyxl
    wb = openpyxl.load_workbook(file_path)
    
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        
        # 열 너비 자동 조정
        for col in ws.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            ws.column_dimensions[col_letter].width = max(max_length + 4, 12)
            
    wb.save(file_path)
    return wb

# --- [실행 예시] ---
if __name__ == "__main__":
    # raw_file = "월별_대리점_정산_원천데이터.xlsb"  # 본사 수신 파일 경로
    # target_agency = "울산중앙대리점"                 # 담당 대리점명 또는 코드
    # process_agency_commission(raw_file, target_agency, "울산중앙대리점_정산요약.xlsx")
    pass

            except Exception as e:
                st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")
