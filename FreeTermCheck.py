import streamlit as st
import pandas as pd
import openpyxl
import re
from io import BytesIO
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.cell.text import InlineFont
from openpyxl.cell.rich_text import TextBlock, CellRichText

def extract_numbers_from_bracket(text):
    if not text:
        return 0
    numbers = re.findall(r'\((\d+)\)', str(text))
    return sum(int(n) for n in numbers)

def evaluate_formula_string(text):
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


# --- HTML 화면 출력용 서식 ---
def format_issue_for_html(issue):
    if "[1.학교운영 현황]" in issue:
        issue = issue.replace("[1.학교운영 현황]", "<strong style='color:#0052cc;'>[1.학교운영 현황]</strong>")
    elif "[2. 자유학기 활동]" in issue:
        issue = issue.replace("[2. 자유학기 활동]", "<strong style='color:#00875a;'>[2. 자유학기 활동]</strong>")
    elif "[3. 예산 계획서]" in issue:
        issue = issue.replace("[3. 예산 계획서]", "<strong style='color:#de350b;'>[3. 예산 계획서]</strong>")
    return f"• {issue}"

# --- 엑셀용 Rich Text 변환 함수 ---
def create_excel_rich_text(issues):
    font_sheet1 = InlineFont(color="FF0052CC", b=True)
    font_sheet2 = InlineFont(color="FF00875A", b=True)
    font_sheet3 = InlineFont(color="FFDE350B", b=True)
    font_success = InlineFont(color="FF006600", b=True)
    
    elements = []
    for i, issue in enumerate(issues):
        prefix = "• "
        suffix = "\n" if i < len(issues) - 1 else ""
        
        if "[1.학교운영 현황]" in issue:
            elements.extend([prefix, TextBlock(font_sheet1, "[1.학교운영 현황]"), issue.replace("[1.학교운영 현황]", "") + suffix])
        elif "[2. 자유학기 활동]" in issue:
            elements.extend([prefix, TextBlock(font_sheet2, "[2. 자유학기 활동]"), issue.replace("[2. 자유학기 활동]", "") + suffix])
        elif "[3. 예산 계획서]" in issue:
            elements.extend([prefix, TextBlock(font_sheet3, "[3. 예산 계획서]"), issue.replace("[3. 예산 계획서]", "") + suffix])
        elif "특이사항 없음" in issue:
            elements.append(TextBlock(font_success, prefix + issue + suffix))
        else:
            elements.append(prefix + issue + suffix)
            
    # 빈 문자열 제거 후 CellRichText 반환
    elements = [e for e in elements if e]
    return CellRichText(*elements) if elements else ""


# --- Streamlit UI ---
st.set_page_config(page_title="자유학기 운영계획서 검토기", layout="wide")

st.markdown("""
<style>
table { width: 100%; border-collapse: collapse; }
th, td { text-align: left; padding: 12px; border: 1px solid #ddd; line-height: 1.6; }
</style>
""", unsafe_allow_html=True)

st.title("📄 자유학기 운영계획서 자동 검토 웹앱")
st.write("여러 개의 엑셀 파일(.xlsx)을 아래 영역에 마우스로 드래그 앤 드롭하거나 클릭하여 업로드하세요.")

# 드래그 앤 드롭을 지원하는 Streamlit 내장 파일 업로더
uploaded_files = st.file_uploader("검토할 엑셀 파일 업로드", type=['xlsx'], accept_multiple_files=True)

if uploaded_files:
    if st.button("검토 시작"):
        report_data = []
        raw_issues_dict = {} # 엑셀 다운로드를 위한 원본 데이터 저장
        
        with st.spinner("파일을 검토하는 중입니다..."):
            for file in uploaded_files:
                filename, issues = process_file(file)
                
                # 화면 출력용 HTML 구성
                html_issues = "<br>".join([format_issue_for_html(issue) for issue in issues])
                is_success = "특이사항 없음" in "".join(issues)
                
                report_data.append({
                    "파일명": filename, 
                    "검토 결과": html_issues,
                    "is_success": is_success
                })
                raw_issues_dict[filename] = issues
        
        # 이상 없는 파일이 맨 위에 오도록 내림차순 정렬
        report_data.sort(key=lambda x: x["is_success"], reverse=True)
        
        # DataFrame 생성 및 출력
        df_results = pd.DataFrame(report_data)
        display_df = df_results[['파일명', '검토 결과']].copy()
        
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
        
        # --- 엑셀 파일 다운로드 생성 (openpyxl 활용하여 색상 및 줄바꿈 적용) ---
        excel_buffer = BytesIO()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "검토결과"
        
        # 헤더 설정
        headers = ["파일명", "검토 결과"]
        ws.append(headers)
        header_font = Font(bold=True)
        header_fill = PatternFill(start_color="FFF4F5F7", end_color="FFF4F5F7", fill_type="solid")
        for col in range(1, 3):
            cell = ws.cell(row=1, column=col)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        # 데이터 채우기
        for idx, row_data in enumerate(report_data, start=2):
            filename = row_data["파일명"]
            is_success = row_data["is_success"]
            issues = raw_issues_dict[filename]
            
            # 파일명 셀 설정
            cell_filename = ws.cell(row=idx, column=1, value=filename)
            cell_filename.alignment = Alignment(vertical='top', wrapText=True)
            if is_success:
                cell_filename.fill = PatternFill(start_color="FFE6FFE6", end_color="FFE6FFE6", fill_type="solid")
                cell_filename.font = Font(color="FF006600", bold=True)
                
            # 검토 결과 셀 (Rich Text 및 줄바꿈 적용)
            cell_result = ws.cell(row=idx, column=2)
            cell_result.value = create_excel_rich_text(issues)
            cell_result.alignment = Alignment(vertical='top', wrapText=True)
            
        # 열 너비 설정
        ws.column_dimensions['A'].width = 40
        ws.column_dimensions['B'].width = 100
        
        wb.save(excel_buffer)
        
        # 다운로드 버튼
        st.download_button(
            label="📥 검토 결과 다운로드 (Excel)",
            data=excel_buffer.getvalue(),
            file_name="운영계획서_검토결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
