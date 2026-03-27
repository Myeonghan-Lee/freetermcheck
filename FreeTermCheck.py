import streamlit as st
import pandas as pd
import openpyxl
import re
import io

# --- 헬퍼 함수 ---
def extract_and_sum_numbers(text):
    """문자열에서 괄호 안의 숫자를 찾아 모두 더함 (예: '도덕(17), 수학(17)' -> 34)"""
    if not text: return 0
    numbers = re.findall(r'\((\d+)\)', str(text))
    return sum(int(n) for n in numbers)

def calculate_basis(text):
    """산출근거 문자열에서 수식만 추출하여 계산 (예: '15,000원*225명' -> 3375000)"""
    if not text: return 0
    text = str(text).replace(',', '').replace(' ', '')
    # 숫자와 사칙연산 기호만 남김
    cleaned = re.sub(r'[^\d\+\-\*\/]', '', text)
    try:
        return eval(cleaned) if cleaned else 0
    except:
        return None

# --- 파일 검토 메인 함수 ---
def review_excel_file(file):
    results = []
    wb = openpyxl.load_workbook(file, data_only=True) # 수식 결과값만 읽기
    
    # ---------------------------------------------------------
    # 1. 학교운영 현황 시트 검토
    # ---------------------------------------------------------
    if "1.학교운영 현황" in wb.sheetnames:
        ws1 = wb["1.학교운영 현황"]
        subject_hours_target = ws1['D8'].value or 0
        career_hours_target = ws1['D9'].value or 0
        grade1_classes = ws1['C11'].value or 0
        
        # 2-3-1 검토: E8 셀(병합됨) 숫자 합산 비교
        e8_val = ws1['E8'].value
        if extract_and_sum_numbers(e8_val) != subject_hours_target:
            results.append("⚠️ [1.학교운영 현황] E8 셀의 과목 시수 합이 D8(주제선택 시수)과 일치하지 않습니다.")
            
        # 2-3-2 검토: E9 셀(병합됨) 숫자 합산 비교
        e9_val = ws1['E9'].value
        if extract_and_sum_numbers(e9_val) != career_hours_target:
            results.append("⚠️ [1.학교운영 현황] E9 셀의 과목 시수 합이 D9(진로탐색 시수)과 일치하지 않습니다.")
    else:
        results.append("❌ '1.학교운영 현황' 시트가 없습니다.")
        return results

    # ---------------------------------------------------------
    # 2. 자유학기 활동 시트 검토
    # ---------------------------------------------------------
    has_outsourced = False
    if "2. 자유학기 활동" in wb.sheetnames:
        ws2 = wb["2. 자유학기 활동"]
        subject_total_g = 0
        career_total_g = 0
        
        current_mode = None
        for row in range(1, ws2.max_row + 1):
            cell_a = ws2.cell(row=row, column=1).value
            
            if cell_a == '주제선택 활동':
                current_mode = 'subject'
                continue
            elif cell_a == '진로 탐색 활동':
                current_mode = 'career'
                continue
                
            prog_name = ws2.cell(row=row, column=2).value
            if prog_name and current_mode:
                hours = ws2.cell(row=row, column=7).value
                teacher = ws2.cell(row=row, column=5).value
                
                if isinstance(hours, (int, float)):
                    if current_mode == 'subject': subject_total_g += hours
                    elif current_mode == 'career': career_total_g += hours
                
                if teacher == '개인위탁':
                    has_outsourced = True
                    
        # 3-7-1-2 검토
        if subject_total_g < (subject_hours_target * grade1_classes):
            results.append(f"⚠️ [2. 자유학기 활동] 주제선택 활동 총 시수({subject_total_g})가 기준({subject_hours_target * grade1_classes})보다 부족합니다.")
            
        # 3-7-2-2 검토
        if career_total_g < (career_hours_target * grade1_classes):
            results.append(f"⚠️ [2. 자유학기 활동] 진로 탐색 활동 총 시수({career_total_g})가 기준({career_hours_target * grade1_classes})보다 부족합니다.")
    
    # ---------------------------------------------------------
    # 3. 예산 계획서 시트 검토
    # ---------------------------------------------------------
    if "3. 예산 계획서" in wb.sheetnames:
        ws3 = wb["3. 예산 계획서"]
        
        # 4-6-1, 4-6-2, 4-6-4 검토: 6행~30행 순회
        for row in range(6, 31):
            content = ws3.cell(row=row, column=2).value
            basis = ws3.cell(row=row, column=3).value
            budget = ws3.cell(row=row, column=4).value
            
            # 4-6-3 개인위탁 검토
            if row == 17:
                if content != '프로그램 개인위탁 운영비':
                    results.append("⚠️ [3. 예산 계획서] B17 셀은 '프로그램 개인위탁 운영비'여야 합니다.")
                if has_outsourced:
                    if not basis or not budget:
                        results.append("⚠️ [3. 예산 계획서] 개인위탁 교사가 있으나 B17 산출근거/소요예산이 누락되었습니다.")
            
            if content:
                if not basis:
                    results.append(f"⚠️ [3. 예산 계획서] {row}행: '내 용'이 있으나 '산출근거'가 없습니다.")
                else:
                    calc_val = calculate_basis(basis)
                    if calc_val is not None and budget is not None:
                        if calc_val != budget:
                            results.append(f"⚠️ [3. 예산 계획서] {row}행: 산출근거 계산값({calc_val})과 소요예산({budget})이 불일치합니다.")
                            
        # 4-6-5 업무추진비 검토
        total_budget = ws3['E3'].value or 0
        biz_expense = ws3['D31'].value or 0
        if biz_expense > (total_budget * 0.03):
             results.append(f"⚠️ [3. 예산 계획서] D31 업무추진비({biz_expense})가 총 예산의 3%({total_budget * 0.03})를 초과합니다.")
             
    if not results:
        results.append("✅ 모든 검토 항목 이상 없음")
        
    return results

# --- Streamlit UI 구성 ---
st.set_page_config(page_title="자유학기 운영계획서 자동 검토기", layout="wide")
st.title("📄 자유학기 운영계획서 자동 검토 웹앱")

uploaded_files = st.file_uploader("검토할 엑셀 파일(.xlsx)을 여러 개 올려주세요.", type=['xlsx'], accept_multiple_files=True)

if uploaded_files:
    if st.button("검토 시작"):
        all_results = []
        
        for file in uploaded_files:
            with st.spinner(f"'{file.name}' 검토 중..."):
                file_issues = review_excel_file(file)
                
                # 화면에 아코디언 형태로 결과 표시
                with st.expander(f"📁 {file.name} 검토 결과", expanded=True):
                    for issue in file_issues:
                        st.write(issue)
                
                # 다운로드용 데이터 수집
                for issue in file_issues:
                    all_results.append({"파일명": file.name, "검토 결과": issue})
                    
        # 5-4. 결과를 데이터프레임으로 변환하여 다운로드 버튼 제공
        df_results = pd.DataFrame(all_results)
        
        # 엑셀 파일로 메모리에 저장
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_results.to_excel(writer, index=False, sheet_name='검토결과')
        processed_data = output.getvalue()
        
        st.success("모든 파일의 검토가 완료되었습니다.")
        st.download_button(
            label="📥 전체 검토 결과 엑셀 다운로드",
            data=processed_data,
            file_name="자유학기_운영계획서_검토결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
