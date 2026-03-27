import streamlit as st
import pandas as pd
import openpyxl
import re
from io import BytesIO

def extract_numbers_from_bracket(text):
    """문자열에서 괄호 안의 숫자를 찾아 합산합니다. 예: '수학(17), 과학(17)' -> 34"""
    if not text:
        return 0
    numbers = re.findall(r'\((\d+)\)', str(text))
    return sum(int(n) for n in numbers)

def evaluate_formula_string(text):
    """'70명 * 50,000원'과 같은 문자열에서 숫자와 기호만 추출하여 계산합니다."""
    if not text:
        return 0
    try:
        # 숫자와 기본 사칙연산 기호만 남기고 모두 제거
        clean_expr = re.sub(r'[^\d\.\*\+\-\/]', '', str(text))
        if clean_expr:
            return eval(clean_expr)
    except:
        return 0
    return 0

def process_file(file):
    results = []
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        filename = file.name
        
        # --- 1. 학교운영 현황 검토 ---
        ws1 = wb.get_sheet_by_name("1.학교운영 현황") if "1.학교운영 현황" in wb.sheetnames else None
        if ws1:
            theme_hours = ws1['D8'].value or 0
            career_hours = ws1['D9'].value or 0
            class_count = ws1['C11'].value or 0
            
            # 2-3-1 & 2-3-2 검토
            e8_text = ws1['E8'].value
            e9_text = ws1['E9'].value
            
            if extract_numbers_from_bracket(e8_text) != theme_hours:
                results.append(f"[1.학교운영 현황] D8(주제선택 시수: {theme_hours})와 E8 병합셀의 과목 시수 합이 불일치합니다.")
            if extract_numbers_from_bracket(e9_text) != career_hours:
                results.append(f"[1.학교운영 현황] D9(진로탐색 시수: {career_hours})와 E9 병합셀의 과목 시수 합이 불일치합니다.")
        else:
            results.append("시트 '1.학교운영 현황'을 찾을 수 없습니다.")
            return filename, results

        # --- 2. 자유학기 활동 검토 ---
        ws2 = wb.get_sheet_by_name("2. 자유학기 활동") if "2. 자유학기 활동" in wb.sheetnames else None
        has_outsourced = False # 개인위탁 여부 확인용
        
        if ws2:
            current_section = None
            total_theme_hours = 0
            total_career_hours = 0
            
            for row in range(5, ws2.max_row + 1):
                # A열 또는 B열에 섹션 제목이 있는지 확인 (병합 셀 고려)
                cell_val = str(ws2.cell(row=row, column=1).value or ws2.cell(row=row, column=2).value or "")
                
                if '주제선택 활동' in cell_val:
                    current_section = 'theme'
                    continue
                elif '진로 탐색 활동' in cell_val:
                    current_section = 'career'
                    continue
                    
                g_val = ws2.cell(row=row, column=7).value
                e_val = ws2.cell(row=row, column=5).value
                
                # 시수 합산
                if isinstance(g_val, (int, float)):
                    if current_section == 'theme':
                        total_theme_hours += g_val
                    elif current_section == 'career':
                        total_career_hours += g_val
                        
                # 개인위탁 확인
                if e_val and '개인위탁' in str(e_val):
                    has_outsourced = True

            # 3-7-1 & 3-7-2 검토
            required_theme = theme_hours * class_count
            required_career = career_hours * class_count
            
            if total_theme_hours < required_theme:
                results.append(f"[2. 자유학기 활동] 주제선택 총 시수({total_theme_hours})가 기준({required_theme})보다 부족합니다.")
            if total_career_hours < required_career:
                results.append(f"[2. 자유학기 활동] 진로탐색 총 시수({total_career_hours})가 기준({required_career})보다 부족합니다.")
        else:
            results.append("시트 '2. 자유학기 활동'을 찾을 수 없습니다.")

        # --- 3. 예산 계획서 검토 ---
        ws3 = wb.get_sheet_by_name("3. 예산 계획서") if "3. 예산 계획서" in wb.sheetnames else None
        if ws3:
            total_budget = ws3['E3'].value or 0
            
            # 6행부터 30행까지 검토
            for row in range(6, 31):
                b_val = ws3.cell(row=row, column=2).value
                c_val = ws3.cell(row=row, column=3).value
                d_val = ws3.cell(row=row, column=4).value
                
                # 17행 특별 검토
                if row == 17:
                    if str(b_val).strip() != '프로그램 개인위탁 운영비':
                        results.append("[3. 예산 계획서] B17 셀의 내용이 '프로그램 개인위탁 운영비'가 아닙니다.")
                    if has_outsourced:
                        if not c_val or not d_val:
                            results.append("[3. 예산 계획서] 개인위탁 교사가 있으나 17행의 산출근거 또는 소요예산이 누락되었습니다.")
                    continue
                
                # 일반 내역 검토 (B열에 내용이 있는 경우)
                if b_val:
                    if not c_val:
                        results.append(f"[3. 예산 계획서] {row}행: 내용이 있으나 산출근거가 없습니다.")
                    else:
                        calculated_budget = evaluate_formula_string(c_val)
                        if calculated_budget > 0 and d_val:
                            # 약간의 오차 허용 (예: 소수점)
                            if abs(calculated_budget - d_val) > 10: 
                                results.append(f"[3. 예산 계획서] {row}행: 산출근거 계산값과 소요예산이 일치하지 않습니다. (입력값: {d_val})")
            
            # 4-6-5 검토 (업무추진비 상한선)
            biz_expense = ws3['D31'].value or 0
            if biz_expense > (total_budget * 0.03):
                results.append(f"[3. 예산 계획서] D31 업무추진비({biz_expense})가 총 예산의 3%를 초과합니다.")
                
        else:
            results.append("시트 '3. 예산 계획서'를 찾을 수 없습니다.")

        if not results:
            results.append("모든 검토 항목을 통과했습니다. (특이사항 없음)")

    except Exception as e:
        results.append(f"파일을 읽는 중 오류가 발생했습니다: {e}")
        
    return filename, results

# --- Streamlit UI ---
st.set_page_config(page_title="자유학기 운영계획서 검토기", layout="wide")
st.title("📄 자유학기 운영계획서 자동 검토 웹앱")
st.write("여러 개의 엑셀 파일(.xlsx)을 업로드하면 운영계획서 작성 지침에 맞는지 자동으로 검토합니다.")

uploaded_files = st.file_uploader("검토할 엑셀 파일들을 업로드하세요", type=['xlsx'], accept_multiple_files=True)

if uploaded_files:
    if st.button("검토 시작"):
        report_data = []
        
        with st.spinner("파일을 검토하는 중입니다..."):
            for file in uploaded_files:
                filename, issues = process_file(file)
                
                # 1. 파일별로 나온 여러 개의 검토 결과를 줄바꿈(\n)으로 연결하여 한 줄로 만듭니다.
                combined_issues = "\n".join(issues)
                
                # 2. '특이사항 없음' 문구가 있는지 확인하여 이상 유무를 판별합니다.
                is_success = "특이사항 없음" in combined_issues
                
                report_data.append({
                    "파일명": filename, 
                    "검토 결과": combined_issues,
                    "is_success": is_success  # 정렬을 위한 임시 키
                })
        
        # 3. 이상이 없는 파일(is_success == True)이 맨 위에 오도록 내림차순 정렬합니다.
        report_data.sort(key=lambda x: x["is_success"], reverse=True)
        
        # 데이터프레임 생성 및 정렬용 임시 키 제거
        df_results = pd.DataFrame(report_data)
        display_df = df_results[['파일명', '검토 결과']]
        
        # 4. 파일명에 초록색 배경과 글자색을 입히는 스타일 함수
        def highlight_success(row):
            if "특이사항 없음" in row['검토 결과']:
                # 파일명(첫 번째 열)에는 초록색 적용, 검토 결과(두 번째 열)는 기본값
                return ['background-color: #e6ffe6; color: #006600; font-weight: bold', '']
            else:
                return [''] * len(row)
        
        # 데이터프레임에 스타일 적용
        styled_df = display_df.style.apply(highlight_success, axis=1)
        
        st.subheader("📊 검토 결과")
        # 적용된 스타일을 Streamlit 화면에 출력
        st.dataframe(styled_df, use_container_width=True)
        
        # 결과 다운로드 기능 (CSV 형식)
        csv = display_df.to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            label="📥 검토 결과 다운로드 (CSV)",
            data=csv,
            file_name="운영계획서_검토결과.csv",
            mime="text/csv"
        )
