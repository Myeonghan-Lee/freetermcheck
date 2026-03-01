
import streamlit as st
import pandas as pd
import openpyxl
import re
import io
from collections import defaultdict

st.set_page_config(page_title="자유학기 운영계획서 점검", page_icon="📋", layout="wide")

# ─────────────────────────────────────────────
# 유틸 함수
# ─────────────────────────────────────────────

def safe_float(v):
    """셀 값을 float로 변환, 실패 시 0"""
    if v is None:
        return 0.0
    try:
        return float(v)
    except (ValueError, TypeError):
        # 콤마 제거 후 재시도
        try:
            return float(str(v).replace(",", "").replace(" ", ""))
        except:
            return 0.0

def safe_str(v):
    """셀 값을 문자열로 변환"""
    if v is None:
        return ""
    return str(v).strip()

def get_row_all_text(ws, row_num):
    """해당 행의 모든 셀 텍스트를 합쳐서 반환 (병합셀 대비)"""
    texts = []
    for c in range(1, ws.max_column + 1):
        val = safe_str(ws.cell(row=row_num, column=c).value)
        if val:
            texts.append(val)
    return " ".join(texts)

def parse_calc(expr_str):
    """
    산출근거 문자열에서 숫자 계산을 시도.
    예: '15,000원*225명' → 3,375,000
        '34시시*40,000원*2명' → 2,720,000
        '12개프로그램*100,000원*2기*월/수' → 2,400,000 (숫자만 추출)
        '12명*20,000원' → 240,000
    패턴: '+' 로 분리 → 각각에서 숫자 추출 후 곱산 → 합산
    """
    if not expr_str or not expr_str.strip():
        return None

    s = expr_str.replace(",", "").replace("，", "").replace(" ", "")

    # '+' 로 분리 → 각각 곱산 후 합산
    parts = re.split(r'\+', s)
    total = 0
    has_any_number = False
    for part in parts:
        # 곱셈 기호로 분리: *, ×, x, X
        factors = re.split(r'[*×xX]', part)
        product = 1
        found_num = False
        for factor in factors:
            nums = re.findall(r'[\d]+(?:\.[\d]+)?', factor)
            if nums:
                # 각 factor에서 첫 번째 숫자만 사용
                product *= float(nums[0])
                found_num = True
        if found_num:
            total += product
            has_any_number = True

    return total if has_any_number else None


# ─────────────────────────────────────────────
# 시트3 섹션 탐지 함수
# ─────────────────────────────────────────────

def find_section_rows(ws):
    """
    시트3에서 '주제선택 활동' 행과 '진로 탐색 활동' 행의 위치를 찾는다.
    병합셀이 있을 수 있으므로 모든 열을 확인한다.
    """
    max_row = ws.max_row
    topic_row = None      # 주제선택 활동 헤더 행
    career_row = None     # 진로 탐색 활동 헤더 행

    for r in range(1, max_row + 1):
        row_text = get_row_all_text(ws, r)
        # 공백/특수문자 제거하여 비교
        normalized = row_text.replace(" ", "").replace("\n", "")

        if "주제선택활동" in normalized and topic_row is None:
            topic_row = r
        elif "진로탐색활동" in normalized and topic_row is not None and career_row is None:
            career_row = r

    return topic_row, career_row


# ─────────────────────────────────────────────
# 시트4 총예산 탐지 함수
# ─────────────────────────────────────────────

def find_total_budget(ws):
    """
    시트4에서 자유학기 총 예산을 찾는다.
    1순위: E3 셀
    2순위: '자유학기 총 예산' 텍스트가 있는 행에서 숫자가 있는 셀
    3순위: '합계' 행의 소요예산 열
    """
    # 1순위: E3
    val = safe_float(ws.cell(row=3, column=5).value)  # E3
    if val > 0:
        return val

    # 2순위: '자유학기' + '총' + '예산' 텍스트 검색
    for r in range(1, min(ws.max_row + 1, 10)):
        row_text = get_row_all_text(ws, r)
        if "총" in row_text and "예산" in row_text:
            # 해당 행에서 가장 큰 숫자 찾기
            max_val = 0
            for c in range(1, ws.max_column + 1):
                v = safe_float(ws.cell(row=r, column=c).value)
                if v > max_val:
                    max_val = v
            if max_val > 0:
                return max_val

    # 3순위: '합계' 행
    for r in range(ws.max_row, 0, -1):
        row_text = get_row_all_text(ws, r)
        if "합계" in row_text:
            for c in range(1, ws.max_column + 1):
                v = safe_float(ws.cell(row=r, column=c).value)
                if v > 0:
                    return v

    return 0


# ─────────────────────────────────────────────
# 시트4 예산 행 파싱
# ─────────────────────────────────────────────

def find_budget_header_row(ws):
    """시트4에서 '내 용' 또는 '내용' 헤더 행을 찾아 데이터 시작행 반환"""
    for r in range(1, min(ws.max_row + 1, 15)):
        row_text = get_row_all_text(ws, r)
        normalized = row_text.replace(" ", "")
        if "내용" in normalized and ("산출근거" in normalized or "소요예산" in normalized):
            return r  # 헤더행 → 데이터는 다음행부터
    return 5  # 기본값


# ─────────────────────────────────────────────
# 핵심 점검 함수
# ─────────────────────────────────────────────

def inspect_file(uploaded_file):
    """하나의 엑셀 파일을 점검하고 결과 리스트 반환"""
    results = []  # (카테고리, 점검항목, 결과, 상세)

    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    sheets = wb.sheetnames

    if len(sheets) < 4:
        return [("구조", "시트 개수", "❌ 실패",
                 f"시트가 4개 미만입니다 ({len(sheets)}개: {sheets})")], {}

    ws2 = wb[sheets[1]]  # 1.학교운영 현황
    ws3 = wb[sheets[2]]  # 2.자유학기 활동
    ws4 = wb[sheets[3]]  # 3.예산 계획서

    debug_info = {}  # 디버그용

    # ═══════════════════════════════════════════
    # 1. 시트2 시수 점검
    # ═══════════════════════════════════════════
    d8 = safe_float(ws2["D8"].value)   # 주제선택 시수
    d9 = safe_float(ws2["D9"].value)   # 진로탐색 시수
    f6 = safe_float(ws2["F6"].value)   # 자유학기 운영 시수

    sum_d = d8 + d9
    check1 = (sum_d == f6)
    results.append((
        "1. 시트2 – 시수 합산",
        "주제선택 시수(D8) + 진로탐색 시수(D9) = 자유학기 운영 시수(F6)",
        "✅ 통과" if check1 else "❌ 실패",
        f"주제선택({d8:.0f}) + 진로탐색({d9:.0f}) = {sum_d:.0f}  |  자유학기 운영 시수 = {f6:.0f}"
    ))

    # ═══════════════════════════════════════════
    # 2. 시트3 자유학기 활동 점검
    # ═══════════════════════════════════════════
    topic_row, career_row = find_section_rows(ws3)
    debug_info["주제선택 활동 행"] = topic_row
    debug_info["진로 탐색 활동 행"] = career_row

    if topic_row is None or career_row is None:
        results.append((
            "2. 시트3 – 구조",
            "'주제선택 활동' / '진로 탐색 활동' 섹션 탐지",
            "❌ 실패",
            f"섹션을 찾을 수 없습니다. (주제선택 행={topic_row}, 진로탐색 행={career_row})"
        ))
        topic_programs = []
        career_programs = []
    else:
        results.append((
            "2. 시트3 – 구조",
            "'주제선택 활동' / '진로 탐색 활동' 섹션 탐지",
            "✅ 통과",
            f"주제선택 활동: {topic_row}행 | 진로 탐색 활동: {career_row}행"
        ))

        # ── 주제선택 활동 파싱: topic_row+1 ~ career_row-1 ──
        topic_programs = []
        for r in range(topic_row + 1, career_row):
            prog = safe_str(ws3.cell(row=r, column=2).value)    # B열: 프로그램명
            subj = safe_str(ws3.cell(row=r, column=3).value)    # C열: 연계교과
            teacher = safe_str(ws3.cell(row=r, column=5).value) # E열: 지도교사
            sessions = safe_float(ws3.cell(row=r, column=6).value)  # F열: 학기당 운영 회기
            hours = safe_float(ws3.cell(row=r, column=7).value)     # G열: 학기당 총 운영 시수
            if prog:  # 프로그램명이 있는 행만
                topic_programs.append({
                    "row": r, "프로그램명": prog, "연계교과": subj,
                    "지도교사": teacher, "운영회기": sessions, "총운영시수": hours
                })

        # ── 진로 탐색 활동 파싱: career_row+1 ~ 끝 ──
        career_programs = []
        for r in range(career_row + 1, ws3.max_row + 1):
            prog = safe_str(ws3.cell(row=r, column=2).value)
            subj = safe_str(ws3.cell(row=r, column=3).value)
            teacher = safe_str(ws3.cell(row=r, column=5).value)
            sessions = safe_float(ws3.cell(row=r, column=6).value)
            hours = safe_float(ws3.cell(row=r, column=7).value)
            if prog:
                career_programs.append({
                    "row": r, "프로그램명": prog, "연계교과": subj,
                    "지도교사": teacher, "운영회기": sessions, "총운영시수": hours
                })

    debug_info["주제선택 프로그램 수"] = len(topic_programs)
    debug_info["진로탐색 프로그램 수"] = len(career_programs)

    # ── 2-라. 주제선택: 프로그램별 학기당 운영 회기 ≥ 2 ──
    for p in topic_programs:
        ok = (p["운영회기"] >= 2)
        results.append((
            "2-라. 주제선택 운영 회기",
            f"'{p['프로그램명']}' 운영 회기 ≥ 2회",
            "✅ 통과" if ok else "❌ 실패",
            f"운영 회기 = {p['운영회기']:.0f}회 (행 {p['row']})"
        ))

    # ── 2-나. 주제선택: 총 운영 시수 합 ≥ D8 ──
    topic_hours_sum = sum(p["총운영시수"] for p in topic_programs)
    check2b = (topic_hours_sum >= d8) if d8 > 0 else False
    results.append((
        "2-나. 주제선택 시수 합산",
        "주제선택 프로그램 총 운영 시수 합 ≥ 주제선택 시수(D8)",
        "✅ 통과" if check2b else "❌ 실패",
        f"프로그램별 시수 합 = {topic_hours_sum:.0f}  |  D8(주제선택 시수) = {d8:.0f}"
    ))

    # ── 2-다. 진로탐색: 총 운영 시수 합 == D9 ──
    career_hours_sum = sum(p["총운영시수"] for p in career_programs)
    check2c = (career_hours_sum == d9) if d9 > 0 else False
    results.append((
        "2-다. 진로탐색 시수 합산",
        "진로탐색 프로그램 총 운영 시수 합 = 진로탐색 시수(D9)",
        "✅ 통과" if check2c else "❌ 실패",
        f"프로그램별 시수 합 = {career_hours_sum:.0f}  |  D9(진로탐색 시수) = {d9:.0f}"
    ))

    # ═══════════════════════════════════════════
    # 3. 시트4 예산 점검
    # ═══════════════════════════════════════════

    # ── 총 예산 찾기 ──
    total_budget = find_total_budget(ws4)
    debug_info["자유학기 총 예산"] = total_budget

    # ── 데이터 행 파싱 ──
    header_row = find_budget_header_row(ws4)
    data_start = header_row + 1

    budget_rows = []
    current_category = ""

    for r in range(data_start, ws4.max_row + 1):
        a_val = safe_str(ws4.cell(row=r, column=1).value)
        b_val = safe_str(ws4.cell(row=r, column=2).value)
        c_val = safe_str(ws4.cell(row=r, column=3).value)
        d_val_raw = ws4.cell(row=r, column=4).value
        d_val = safe_float(d_val_raw)
        e_val = safe_str(ws4.cell(row=r, column=5).value)

        # 카테고리 업데이트 (A열에 값이 있으면)
        if a_val:
            # '소계'나 '합계'는 카테고리가 아님
            if "소계" not in a_val and "합계" not in a_val:
                current_category = a_val

        # '소계'/'합계' 행은 데이터에서 제외
        row_text = get_row_all_text(ws4, r)
        if "소계" in row_text or "합계" in row_text:
            continue

        # 내용(B열)이 있는 행만 데이터로 수집
        if b_val:
            budget_rows.append({
                "row": r, "카테고리": current_category,
                "내용": b_val, "산출근거": c_val,
                "소요예산": d_val, "비고": e_val
            })

    debug_info["예산 데이터 행 수"] = len(budget_rows)

    # ── 3-라. 개인위탁 존재 시 예산 입력 확인 ──
    all_teachers = [p["지도교사"] for p in topic_programs + career_programs]
    has_personal_consign = any("개인위탁" in t for t in all_teachers)

    consign_programs = [p["프로그램명"] for p in topic_programs + career_programs if "개인위탁" in p["지도교사"]]

    if has_personal_consign:
        consign_budget_exists = any(
            "개인위탁" in br["내용"] for br in budget_rows
        )
        results.append((
            "3-라. 개인위탁 예산",
            "시트3에 '개인위탁' 지도교사 존재 → 시트4에 개인위탁 운영비 입력 확인",
            "✅ 통과" if consign_budget_exists else "❌ 실패",
            f"개인위탁 프로그램: {consign_programs} → 예산 {'입력됨' if consign_budget_exists else '미입력'}"
        ))
    else:
        results.append((
            "3-라. 개인위탁 예산",
            "시트3에 '개인위탁' 지도교사 존재 여부",
            "ℹ️ 해당없음",
            "개인위탁 지도교사가 없으므로 점검 생략"
        ))

    # ── 3-마. 내용별 산출근거 입력 여부 ──
    missing_calc = []
    for br in budget_rows:
        if br["내용"] and not br["산출근거"]:
            missing_calc.append(f"행{br['row']}: '{br['내용']}'")
    check3b = len(missing_calc) == 0
    results.append((
        "3-마. 산출근거 입력",
        "모든 내용(B열)에 산출근거(C열) 입력 여부",
        "✅ 통과" if check3b else "❌ 실패",
        "모두 입력됨" if check3b else "미입력 항목: " + " / ".join(missing_calc)
    ))

    # ── 3-바. 산출근거 계산 vs 소요예산 일치 ──
    calc_mismatches = []
    calc_details = []
    for br in budget_rows:
        if br["산출근거"] and br["소요예산"] > 0:
            calc_result = parse_calc(br["산출근거"])
            if calc_result is not None:
                diff = abs(calc_result - br["소요예산"])
                if diff > 1:  # 1원 오차 허용
                    calc_mismatches.append(
                        f"행{br['row']}: '{br['내용']}' → 산출근거 계산={calc_result:,.0f}원, 소요예산={br['소요예산']:,.0f}원 (차이: {diff:,.0f}원)"
                    )
                calc_details.append(
                    f"행{br['row']}: '{br['내용']}' | 산출근거='{br['산출근거']}' → 계산={calc_result:,.0f} vs 예산={br['소요예산']:,.0f}"
                )
    check3c = len(calc_mismatches) == 0
    detail_text = "모두 일치" if check3c else " / ".join(calc_mismatches)
    results.append((
        "3-바. 산출근거 일치",
        "산출근거 계산값과 소요예산(D열) 일치 여부",
        "✅ 통과" if check3c else "❌ 실패",
        detail_text
    ))
    debug_info["산출근거 계산 상세"] = calc_details

    # ── 3-사. 업무추진비 < 총 예산의 3% ──
    admin_budget = sum(
        br["소요예산"] for br in budget_rows
        if "업무추진비" in br["카테고리"]
    )
    debug_info["업무추진비 합계"] = admin_budget

    if total_budget > 0:
        admin_ratio = admin_budget / total_budget * 100
        check3d = (admin_ratio < 3)
        results.append((
            "3-사. 업무추진비 비율",
            "업무추진비 < 총 예산의 3%",
            "✅ 통과" if check3d else "❌ 실패",
            f"업무추진비 = {admin_budget:,.0f}원 ({admin_ratio:.2f}%)  |  총 예산 = {total_budget:,.0f}원  |  3% 기준 = {total_budget * 0.03:,.0f}원"
        ))
    else:
        results.append((
            "3-사. 업무추진비 비율",
            "업무추진비 < 총 예산의 3%",
            "⚠️ 경고",
            f"총 예산을 찾을 수 없습니다 (E3={ws4['E3'].value}). 시트4 구조를 확인해주세요."
        ))

    # ── 3-아. 개인위탁 운영비 ≤ 총 예산의 40% ──
    consign_budget = sum(
        br["소요예산"] for br in budget_rows
        if "개인위탁" in br["내용"]
    )
    debug_info["개인위탁 운영비"] = consign_budget

    if total_budget > 0:
        if consign_budget > 0:
            consign_ratio = consign_budget / total_budget * 100
            check3e = (consign_ratio <= 40)
            results.append((
                "3-아. 개인위탁 비율",
                "프로그램 개인위탁 운영비 ≤ 총 예산의 40%",
                "✅ 통과" if check3e else "❌ 실패",
                f"개인위탁 운영비 = {consign_budget:,.0f}원 ({consign_ratio:.2f}%)  |  총 예산 = {total_budget:,.0f}원  |  40% 기준 = {total_budget * 0.4:,.0f}원"
            ))
        else:
            results.append((
                "3-아. 개인위탁 비율",
                "프로그램 개인위탁 운영비 ≤ 총 예산의 40%",
                "ℹ️ 해당없음",
                "개인위탁 운영비 항목이 없습니다."
            ))
    else:
        results.append((
            "3-아. 개인위탁 비율",
            "프로그램 개인위탁 운영비 ≤ 총 예산의 40%",
            "⚠️ 경고",
            "총 예산을 찾을 수 없습니다."
        ))

    wb.close()

    # 부가 정보 (프로그램 목록)
    extra = {
        "debug": debug_info,
        "topic_programs": topic_programs,
        "career_programs": career_programs,
        "budget_rows": budget_rows,
        "d8": d8, "d9": d9, "f6": f6,
        "total_budget": total_budget
    }

    return results, extra


# ─────────────────────────────────────────────
# Streamlit UI
# ─────────────────────────────────────────────

st.title("📋 자유학기 운영계획서 점검 시스템")
st.markdown("""
> **2026학년도 중학교 자유학기 운영계획서** 엑셀 파일을 업로드하면  
> 시수 · 프로그램 · 예산 항목을 자동으로 점검합니다.  
> 여러 학교 파일을 **한 번에** 업로드할 수 있습니다.
""")

st.divider()

uploaded_files = st.file_uploader(
    "운영계획서 엑셀(.xlsx) 파일 업로드",
    type=["xlsx"],
    accept_multiple_files=True,
    help="여러 파일을 동시에 선택하거나 드래그하여 업로드하세요."
)

if uploaded_files:
    # 전체 요약 카운터
    total_pass = 0
    total_fail = 0
    total_warn = 0
    all_school_results = {}

    for uf in uploaded_files:
        st.markdown(f"---")
        st.markdown(f"## 🏫 {uf.name}")

        try:
            file_copy = io.BytesIO(uf.read())
            uf.seek(0)
            res, extra = inspect_file(file_copy)

            pass_cnt = sum(1 for r in res if "통과" in r[2])
            fail_cnt = sum(1 for r in res if "실패" in r[2])
            warn_cnt = sum(1 for r in res if "경고" in r[2] or "해당없음" in r[2])

            total_pass += pass_cnt
            total_fail += fail_cnt
            total_warn += warn_cnt

            # 학교별 요약 메트릭
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("✅ 통과", f"{pass_cnt}건")
            col2.metric("❌ 실패", f"{fail_cnt}건")
            col3.metric("⚠️ 기타", f"{warn_cnt}건")
            col4.metric("📊 총 예산", f"{extra.get('total_budget',0):,.0f}원")

            # 기본 정보 표시
            with st.expander("📌 파일 기본 정보", expanded=False):
                info_col1, info_col2 = st.columns(2)
                with info_col1:
                    st.markdown(f"""
                    | 항목 | 값 |
                    |:---|:---|
                    | 주제선택 시수 (D8) | **{extra.get('d8',0):.0f}** |
                    | 진로탐색 시수 (D9) | **{extra.get('d9',0):.0f}** |
                    | 자유학기 운영 시수 (F6) | **{extra.get('f6',0):.0f}** |
                    | 자유학기 총 예산 (E3) | **{extra.get('total_budget',0):,.0f}원** |
                    """)
                with info_col2:
                    st.markdown("**주제선택 프로그램:**")
                    for p in extra.get("topic_programs", []):
                        st.markdown(f"- {p['프로그램명']} ({p['연계교과']}) - {p['지도교사']} | 회기:{p['운영회기']:.0f} | 시수:{p['총운영시수']:.0f}")
                    st.markdown("**진로탐색 프로그램:**")
                    for p in extra.get("career_programs", []):
                        st.markdown(f"- {p['프로그램명']} ({p['연계교과']}) - {p['지도교사']} | 회기:{p['운영회기']:.0f} | 시수:{p['총운영시수']:.0f}")

            # 결과 테이블
            df = pd.DataFrame(res, columns=["카테고리", "점검 항목", "결과", "상세 내용"])

            def highlight_result(row):
                if "실패" in str(row["결과"]):
                    return ["background-color: #ffcccc"] * len(row)
                elif "통과" in str(row["결과"]):
                    return ["background-color: #ccffcc"] * len(row)
                elif "경고" in str(row["결과"]):
                    return ["background-color: #fff3cd"] * len(row)
                else:
                    return ["background-color: #e8f4f8"] * len(row)

            st.dataframe(
                df.style.apply(highlight_result, axis=1),
                use_container_width=True,
                hide_index=True,
                height=min(len(df) * 38 + 50, 800)
            )

            # 실패 항목만 별도 표시
            fail_df = df[df["결과"].str.contains("실패")]
            if not fail_df.empty:
                with st.expander(f"🔴 실패 항목만 보기 ({fail_cnt}건)", expanded=True):
                    st.dataframe(fail_df, use_container_width=True, hide_index=True)
            else:
                st.success("🎉 모든 점검 항목을 통과했습니다!")

            # 디버그 정보
            with st.expander("🔧 디버그 정보 (셀 위치 확인용)", expanded=False):
                st.json(extra.get("debug", {}))
                if extra.get("budget_rows"):
                    st.markdown("**예산 데이터 파싱 결과:**")
                    budget_df = pd.DataFrame(extra["budget_rows"])
                    st.dataframe(budget_df, use_container_width=True, hide_index=True)

            all_school_results[uf.name] = df

        except Exception as e:
            st.error(f"❌ 파일 처리 중 오류: {e}")
            import traceback
            st.code(traceback.format_exc())

    # ═══════ 전체 요약 ═══════
    st.divider()
    st.markdown("## 📊 전체 점검 요약")
    sc1, sc2, sc3, sc4 = st.columns(4)
    sc1.metric("📁 점검 파일 수", f"{len(uploaded_files)}개")
    sc2.metric("✅ 총 통과", f"{total_pass}건")
    sc3.metric("❌ 총 실패", f"{total_fail}건")
    sc4.metric("⚠️ 총 기타", f"{total_warn}건")

    # CSV 다운로드
    if all_school_results:
        all_dfs = []
        for school_name, df in all_school_results.items():
            df_copy = df.copy()
            df_copy.insert(0, "학교(파일명)", school_name)
            all_dfs.append(df_copy)
        combined = pd.concat(all_dfs, ignore_index=True)
        csv_data = combined.to_csv(index=False, encoding="utf-8-sig")
        st.download_button(
            label="📥 전체 점검 결과 CSV 다운로드",
            data=csv_data,
            file_name="자유학기_점검결과_전체.csv",
            mime="text/csv"
        )

else:
    st.info("👆 위에서 엑셀 파일을 업로드하면 자동 점검이 시작됩니다.")

    with st.expander("📌 점검 항목 안내", expanded=True):
        st.markdown("""
| 번호 | 카테고리 | 점검 내용 |
|:---:|:---:|:---|
| **1** | 시트2 시수 | 주제선택 시수(D8) + 진로탐색 시수(D9) = 자유학기 운영 시수(F6) |
| **2-라** | 시트3 주제선택 | 프로그램별 학기당 운영 회기 ≥ 2회 |
| **2-나** | 시트3 주제선택 | 프로그램별 총 운영 시수 합 ≥ 주제선택 시수(D8) |
| **2-다** | 시트3 진로탐색 | 프로그램별 총 운영 시수 합 = 진로탐색 시수(D9) |
| **3-라** | 시트4 개인위탁 | 개인위탁 지도교사 존재 시 예산 입력 여부 |
| **3-마** | 시트4 산출근거 | 모든 내용에 산출근거 입력 여부 |
| **3-바** | 시트4 계산검증 | 산출근거 계산값 = 소요예산 |
| **3-사** | 시트4 업무추진비 | 업무추진비 < 총 예산의 3% |
| **3-아** | 시트4 개인위탁 비율 | 개인위탁 운영비 ≤ 총 예산의 40% |
        """)

    with st.expander("📂 엑셀 파일 구조 안내", expanded=False):
        st.markdown("""
        **시트 구성:**
        - 시트1: 작성안내
        - 시트2: 1.학교운영 현황 (`D8`=주제선택시수, `D9`=진로탐색시수, `F6`=자유학기운영시수)
        - 시트3: 2.자유학기 활동 (`B`=프로그램명, `C`=연계교과, `E`=지도교사, `F`=운영회기, `G`=총운영시수)
        - 시트4: 3.예산 계획서 (`E3`=총예산, `A`=카테고리, `B`=내용, `C`=산출근거, `D`=소요예산)
        """)

st.markdown("---")
st.caption("© 2026 자유학기 운영계획서 점검 시스템 | Streamlit")
