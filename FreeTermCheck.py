import streamlit as st
import pandas as pd
import openpyxl
import re
from io import BytesIO
from openpyxl.styles import Font, Color, Alignment, PatternFill

def extract_numbers_from_bracket(text):
    """문자열에서 괄호 안의 숫자를 찾아 합산합니다."""
    if not text:
        return 0
    numbers = re.findall(r'\((\d+)\)', str(text))
    return sum(int(n) for n in numbers)

def evaluate_formula_string(text):
    """문자열에서 숫자와 기호만 추출하여 계산합니다."""
    if not text:
        return 0
    try:
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
        ws1 = wb["1.학교운영 현황"] if "1.학교운영 현황" in wb.sheetnames else None
        theme_hours, career_hours, class_count = 0, 0, 0
        
        if ws1:
            theme_hours = ws1['D8'].value or 0
            career_hours = ws1['D9'].value or 0
            class_count = ws1['C11'].value or 0
            
            e8_text = ws1['E8'].value
            e9_text = ws1['E9'].value
            
            if extract_numbers_from_bracket(e8_text) != theme_hours:
                results.append(f"[1.학교운영 현황] D8(주제선택 시수: {theme_hours})와 E8 병합셀의 과목 시수 합이 불일치합니다.")
            if extract_numbers_from_bracket(e9_text) != career_hours:
                results.append(f"[1.학교운영 현황] D9(진로탐색 시수: {career_hours})와 E9 병합셀의 과목 시수 합이 불일치합니다.")
        else:
            results.append("[1.학교운영 현황] 시트를 찾을 수 없습니다.")

        # --- 2. 자유학기 활동 검토 ---
        ws2 = wb["2. 자유학기 활동"] if "2. 자유학기 활동" in wb.sheetnames else None
        has_outsourced = False
        
        if ws2:
            current_section = None
            total_theme_hours = 0
            total_career_hours = 0
            
            for row in range(5, ws2.max_row + 1):
                cell_val = str(ws2.cell(row=row, column=1).value or ws2.cell(row=row, column=2).value or "")
                
                if '주제선택 활동' in cell_val:
                    current_section = 'theme'
                    continue
                elif '진로 탐색 활동' in cell_val:
                    current_section = 'career'
                    continue
                    
                g_val = ws2.cell(row=row, column=7).value
                e_val = ws2.cell(row=row, column=5).value
                
                if isinstance(g_val, (int, float)):
                    if current_section == 'theme':
                        total_theme_hours += g_val
                    elif current_section == 'career':
                        total_career_hours += g_val
                        
                if e_val and '개인위탁' in str(e_val):
                    has_outsourced = True

            required_theme = theme_hours * class_count
            required_career = career_hours * class_count
            
            if total_theme_hours < required_theme:
                results.append(f"[2. 자유학기 활동] 주제선택 총 시수({total_theme_hours})가 기준({required_theme})보다 부족합니다.")
            if total_career_hours < required_career:
                results.append(f"[2. 자유학기 활동] 진로탐색 총 시수({total_career_hours})가 기준({required_career})보다 부족합니다.")
        else:
            results.append("[2. 자유학기 활동] 시트를 찾을 수 없습니다.")

        # --- 3. 예산 계획서 검토 ---
        ws3 = wb["3. 예산 계획서"] if "3. 예산 계획서" in wb.sheetnames else None
        if ws3:
            total_budget = ws3['E3'].value or 0
            
            for row in range(6, 31):
                b_val = ws3.cell(row=row, column=2).value
                c_val = ws3.cell(row=row, column=3).value
                d_val = ws3.cell(row=row, column=4).value
                
                if row == 17:
                    if str(b_val).strip() != '프로그램 개인위탁 운영비':
                        results.append("[3. 예산 계획서] B17 셀의 내용이 '프로그램 개인위탁 운영비'가 아닙니다.")
                    if has_outsourced:
                        if not c_val or not d_val:
                            results.append("[3. 예산 계획서] 개인위탁 교사가 있으나 17행의 산출근거 또는 소요예산이 누락되었습니다.")
                    continue
                
                if b_val:
                    if not c_val:
                        results.append(f"[3. 예산 계획서] {row}행: 내용이 있으나 산출근거가 없습니다.")
                    else:
                        calculated_budget = evaluate_formula_string(c_val)
                        if calculated_budget > 0 and d_val:
                            if abs(calculated_budget - d_val) > 10: 
                                results.append(f"[3. 예산 계획서] {row}행: 산출근거 계산값과 소요예산이 일치하지 않습니다. (입력값: {d_val})")
            
            biz_expense = ws3['D31'].value or 0
            if biz_expense > (total_budget * 0.03):
                results.append(f"[3. 예산 계획서] D31 업무추진비({biz_expense})가 총 예산의 3%를 초과합니다.")
                
        else:
            results.append("[3. 예산 계획서] 시트를 찾을 수 없습니다.")

        if not results:
            results.append("특이사항 없음 (모든 검토 항목을 통과했습니다.)")

    except Exception as e:
        results.append(f"파일을 읽는 중 오류가 발생했습니다: {e}")
        
    return filename, results

# --- 스타일 적용 함수 (화면 출력용 HTML) ---
def format_issue_for_html(issue):
    if "[1.학교운영 현황]" in issue:
        issue = issue.replace("[1.학교운영 현황]", "<strong style='color:#0052cc;'>[1.학교운영 현황]</strong>")
    elif "[2. 자유학기 활동]" in issue:
        issue = issue.replace("[2. 자유학기 활동]", "<strong style='color:#00875a;'>[2. 자유학기 활동]</strong>")
    elif "[3. 예산 계획서]" in issue:
        issue = issue.replace("[3. 예산 계획서]", "<strong style='color:#de350b;'>[3. 예산 계획서]</strong>")
    return f"• {issue}"

# --- Streamlit UI ---
st.set_page_config(page_title="자유학기 운영계획서 검토기", layout="wide")

st.markdown("""
<style>
table { width: 100%; border-collapse: collapse; }
th, td { text-align: left; padding: 12px; border: 1px solid #ddd; line-height: 1.6; }
</style>
""", unsafe_allow_html=True)

st.title("📄 자유학기 운영계획서 자동 검토 웹앱")
st.write("여러 개의 엑셀 파일(.xlsx)을 업로드하면 운영계획서 작성 지침에 맞는지 자동으로 검토합니다.")

uploaded_files = st.file_uploader("검토할 엑셀 파일들을 업로드하세요", type=['xlsx'], accept_multiple_files=True)

if uploaded_files:
    if st.button("검토 시작"):
        report_data = []
        
        with st.spinner("파일을 검토하는 중입니다..."):
            for file in uploaded_files:
                filename, issues = process_file(file)
                html_issues = "<br>".join([format_issue_for_html(issue) for issue in issues])
                excel_issues = "\n".join([f"• {issue}" for issue in issues])
                is_success = "특이사항 없음" in "".join(issues)
                
                report_data.append({
                    "파일명": filename, 
                    "검토 결과 (화면용)": html_issues,
                    "검토 결과 (Excel용)": excel_issues,
                    "is_success": is_success
                })
        
        report_data.sort(key=lambda x: x["is_success"], reverse=True)
        df_results = pd.DataFrame(report_data)
        
        display_df = df_results[['파일명', '검토 결과 (화면용)']].copy()
        display_df.rename(columns={'검토 결과 (화면용)': '검토 결과'}, inplace=True)
        
        def highlight_success(row):
            if "특이사항 없음" in row['검토 결과']:
                return ['background-color: #e6ffe6; color: #006600; font-weight: bold', '']
            else:
                return [''] * len(row)
        
        styled_df = display_df.style.apply(highlight_success, axis=1)
        html_table = styled_df.hide(axis="index").to_html(escape=False)
        
        st.subheader("📊 검토 결과")
        st.markdown(html_table, unsafe_allow_html=True)
        st.write("---")
        
        # --- 엑셀 파일 생성 및 스타일링 ---
        excel_buffer = BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            # 기본 데이터 저장
            df_for_excel = df_results[['파일명', '검토 결과 (Excel용)']].copy()
            df_for_excel.rename(columns={'검토 결과 (Excel용)': '검토 결과'}, inplace=True)
            df_for_excel.to_excel(writer, index=False, sheet_name='검토결과')
            
            workbook = writer.book
            worksheet = writer.sheets['검토결과']
            
            # 스타일 정의
            blue_font = Font(color="0052CC", bold=True) # 시트1 색상
            green_font = Font(color="00875A", bold=True) # 시트2 색상
            red_font = Font(color="DE350B", bold=True)   # 시트3 색상
            success_fill = PatternFill(start_color="E6FFE6", end_color="E6FFE6", fill_type="solid")
            success_font = Font(color="006600", bold=True)

            # 데이터 행 반복 처리
            for idx, row in enumerate(worksheet.iter_rows(min_row=2, max_col=2, max_row=len(report_data)+1)):
                file_cell, issue_cell = row
                
                # 셀 기본 설정 (줄바꿈, 상단 정렬)
                issue_cell.alignment = Alignment(wrapText=True, vertical='top')
                file_cell.alignment = Alignment(vertical='top')

                # '특이사항 없음'인 경우 행 전체 배경색 적용
                if "특이사항 없음" in str(issue_cell.value):
                    file_cell.fill = success_fill
                    file_cell.font = success_font
                    issue_cell.fill = success_fill
                    issue_cell.font = success_font
                else:
                    # 텍스트 내용에 따라 셀 전체 글자 색상 우선 지정 (주요 에러 색상 기준)
                    if "[1.학교운영 현황]" in str(issue_cell.value):
                        issue_cell.font = Font(color="0052CC")
                    elif "[2. 자유학기 활동]" in str(issue_cell.value):
                        issue_cell.font = Font(color="00875A")
                    elif "[3. 예산 계획서]" in str(issue_cell.value):
                        issue_cell.font = Font(color="DE350B")

            # 열 너비 조절
            worksheet.column_dimensions['A'].width = 30
            worksheet.column_dimensions['B'].width = 110
            
        st.download_button(
            label="📥 검토 결과 다운로드 (Excel)",
            data=excel_buffer.getvalue(),
            file_name="운영계획서_검토결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
