import streamlit as st
import pandas as pd
import openpyxl
import re
from io import BytesIO
from openpyxl.cell.rich_text import TextBlock, CellRichText
from openpyxl.cell.text import InlineFont

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
        ws1 = wb.get_sheet_by_name("1.학교운영 현황") if "1.학교운영 현황" in wb.sheetnames else None
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
        ws2 = wb.get_sheet_by_name("2. 자유학기 활동") if "2. 자유학기 활동" in wb.sheetnames else None
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
        ws3 = wb.get_sheet_by_name("3. 예산 계획서") if "3. 예산 계획서" in wb.sheetnames else None
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


# --- HTML 스타일 적용 함수 ---
def format_issue_for_html(issue):
    """시트 이름별로 다른 글자 색상을 적용하고 HTML 태그로 감쌉니다."""
    if "[1.학교운영 현황]" in issue:
        issue = issue.replace("[1.학교운영 현황]", "<strong style='color:#0052cc;'>[1.학교운영 현황]</strong>")
    elif "[2. 자유학기 활동]" in issue:
        issue = issue.replace("[2. 자유학기 활동]", "<strong style='color:#00875a;'>[2. 자유학기 활동]</strong>")
    elif "[3. 예산 계획서]" in issue:
        issue = issue.replace("[3. 예산 계획서]", "<strong style='color:#de350b;'>[3. 예산 계획서]</strong>")
    return f"• {issue}"

# --- 엑셀 Rich Text 적용 함수 ---
def build_rich_text_for_excel(issues):
    """openpyxl CellRichText를 활용해 엑셀 셀 내 특정 텍스트의 색상을 변경하고 줄바꿈을 처리합니다."""
    elements = []
    
    # 폰트 색상 및 굵기 설정 (HTML 코드와 동일한 계열)
    font_blue = InlineFont(color="0052CC", b=True)
    font_green = InlineFont(color="00875A", b=True)
    font_red = InlineFont(color="DE350B", b=True)
    
    for i, issue in enumerate(issues):
        text = f"• {issue}"
        
        if "[1.학교운영 현황]" in text:
            parts = text.split("[1.학교운영 현황]")
            if parts[0]: elements.append(parts[0])
            elements.append(TextBlock(font_blue, "[1.학교운영 현황]"))
            if parts[1]: elements.append(parts[1])
        elif "[2. 자유학기 활동]" in text:
            parts = text.split("[2. 자유학기 활동]")
            if parts[0]: elements.append(parts[0])
            elements.append(TextBlock(font_green, "[2. 자유학기 활동]"))
            if parts[1]: elements.append(parts[1])
        elif "[3. 예산 계획서]" in text:
            parts = text.split("[3. 예산 계획서]")
            if parts[0]: elements.append(parts[0])
            elements.append(TextBlock(font_red, "[3. 예산 계획서]"))
            if parts[1]: elements.append(parts[1])
        else:
            elements.append(text)
            
        # 마지막 항목이 아니면 줄바꿈 추가
        if i < len(issues) - 1:
            elements.append("\n")
            
    # 빈 문자열 방지 및 리치 텍스트 객체 반환
    return CellRichText([e for e in elements if e])


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
                
                # 1. 화면 출력용 (HTML, 줄바꿈 <br> 적용, 시트별 색상 적용)
                html_issues = "<br>".join([format_issue_for_html(issue) for issue in issues])
                
                # 2. 이상 유무 판별 (성공 여부 확인)
                is_success = "특이사항 없음" in "".join(issues)
                
                report_data.append({
                    "파일명": filename, 
                    "검토 결과 (화면용)": html_issues,
                    "raw_issues": issues, # 엑셀 Rich Text 변환을 위해 원본 리스트 보존
                    "is_success": is_success
                })
        
        # 이상이 없는 파일이 맨 위에 오도록 내림차순 정렬
        report_data.sort(key=lambda x: x["is_success"], reverse=True)
        
        # DataFrame 생성
        df_results = pd.DataFrame(report_data)
        
        # --- 화면 출력 처리 ---
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
        
        # --- 엑셀 파일 다운로드 처리 ---
        excel_buffer = BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            download_df = df_results[['파일명']].copy()
            download_df['검토 결과'] = "" # 텍스트가 들어갈 자리 임시 생성
            download_df.to_excel(writer, index=False, sheet_name='검토결과')
            
            worksheet = writer.sheets['검토결과']
            
            # 정렬된 df_results를 순회하며 Rich Text 삽입
            for row_idx, (_, row) in enumerate(df_results.iterrows(), start=2):
                cell = worksheet.cell(row=row_idx, column=2)
                
                # 위에서 정의한 엑셀 서식 적용 함수 호출
                cell.value = build_rich_text_for_excel(row['raw_issues'])
                
                # 엑셀 파일 내 줄바꿈(\n)이 정상적으로 보이도록 서식(Wrap Text) 적용
                cell.alignment = openpyxl.styles.Alignment(wrapText=True, vertical='top')
            
            # 열 너비 조절 (가독성 향상)
            worksheet.column_dimensions['A'].width = 30
            worksheet.column_dimensions['B'].width = 100
            
        # 다운로드 버튼
        st.download_button(
            label="📥 검토 결과 다운로드 (Excel)",
            data=excel_buffer.getvalue(),
            file_name="운영계획서_검토결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
