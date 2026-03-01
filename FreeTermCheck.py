
# ============================================================
# 📁 프로젝트 구조
# ============================================================
# my-streamlit-app/
# ├── app.py                ← 아래 코드 전체
# ├── requirements.txt      ← streamlit, openpyxl, pandas
# └── README.md
# ============================================================

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
        return 0.0

def safe_str(v):
    """셀 값을 문자열로 변환"""
    if v is None:
        return ""
    return str(v).strip()

def parse_calc(expr_str):
    """
    산출근거 문자열에서 숫자 계산을 시도.
    예: '50,000원 × 3회 × 2명', '50000*3*2' → 300000
    패턴: 숫자(콤마 포함)를 추출하고 모두 곱함
    합산이 필요한 경우 '+' 기준 분리 후 각각 곱산
    """
    if not expr_str or not expr_str.strip():
        return None

    s = expr_str.replace(",", "").replace("，", "")

    # '+' 로 분리 → 각각 곱산 후 합산
    parts = re.split(r'\+', s)
    total = 0
    for part in parts:
        nums = re.findall(r'[\d]+(?:\.[\d]+)?', part)
        if not nums:
            continue
        product = 1
        for n in nums:
            product *= float(n)
        total += product

    return total if total != 0 else None


# ─────────────────────────────────────────────
# 핵심 점검 함수
# ─────────────────────────────────────────────

def inspect_file(uploaded_file):
    """하나의 엑셀 파일을 점검하고 결과 dict 반환"""
    results = []  # (카테고리, 점검항목, 결과, 상세)
    warnings = []

    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    sheets = wb.sheetnames

    if len(sheets) < 4:
        return [("구조", "시트 개수", "❌ 실패", f"시트가 4개 미만입니다 ({len(sheets)}개)")], warnings

    ws2 = wb[sheets[1]]  # 1.학교운영 현황
    ws3 = wb[sheets[2]]  # 2.자유학기 활동
    ws4 = wb[sheets[3]]  # 3.예산 계획서

    # ───── 1. 시트2 시수 점검 ─────
    d8 = safe_float(ws2["D8"].value)   # 주제선택 시수
    d9 = safe_float(ws2["D9"].value)   # 진로탐색 시수
    f6 = safe_float(ws2["F6"].value)   # 자유학기 운영 시수

    sum_d = d8 + d9
    check1 = (sum_d == f6)
    results.append((
        "시트2 – 시수",
        "주제선택 시수(D8) + 진로탐색 시수(D9) = 자유학기 운영 시수(F6)",
        "✅ 통과" if check1 else "❌ 실패",
        f"주제선택({d8}) + 진로탐색({d9}) = {sum_d}  /  자유학기 운영 시수 = {f6}"
    ))

    # ───── 2. 시트3 파싱 ─────
    # B열 스캔하여 '주제선택 활동' 행, '진로 탐색 활동' 행 찾기
    max_row3 = ws3.max_row
    topic_start_row = None
    career_start_row = None

    for r in range(1, max_row3 + 1):
        val = safe_str(ws3.cell(row=r, column=2).value)  # B열
        # A열도 확인 (병합셀 대비)
        val_a = safe_str(ws3.cell(row=r, column=1).value)
        combined = val + " " + val_a
        # 여러 열을 확인
        row_texts = []
        for c in range(1, ws3.max_column + 1):
            row_texts.append(safe_str(ws3.cell(row=r, column=c).value))
        full_row = " ".join(row_texts)

        if "주제선택" in full_row and "활동" in full_row and topic_start_row is None:
            topic_start_row = r
        if "진로" in full_row and "탐색" in full_row and "활동" in full_row and career_start_row is None and topic_start_row is not None:
            if r > (topic_start_row if topic_start_row else 0):
                career_start_row = r

    if topic_start_row is None or career_start_row is None:
        results.append((
            "시트3 – 구조",
            "'주제선택 활동' / '진로 탐색 활동' 행 탐지",
            "⚠️ 경고",
            f"주제선택 행={topic_start_row}, 진로탐색 행={career_start_row}. 구조를 확인해주세요."
        ))
        # 찾지 못한 경우에도 계속 진행하되 빈 리스트 처리
        topic_programs = []
        career_programs = []
    else:
        # 주제선택 활동: topic_start_row+1 ~ career_start_row-1
        topic_programs = []
        for r in range(topic_start_row + 1, career_start_row):
            prog = safe_str(ws3.cell(row=r, column=2).value)   # B열: 프로그램명
            subj = safe_str(ws3.cell(row=r, column=3).value)   # C열: 연계교과
            teacher = safe_str(ws3.cell(row=r, column=5).value) # E열: 지도교사
            sessions = safe_float(ws3.cell(row=r, column=6).value)  # F열: 학기당 운영 회기
            hours = safe_float(ws3.cell(row=r, column=7).value)    # G열: 학기당 총 운영 시수
            if prog:
                topic_programs.append({
                    "row": r, "프로그램명": prog, "연계교과": subj,
                    "지도교사": teacher, "운영회기": sessions, "총운영시수": hours
                })

        # 진로 탐색 활동: career_start_row+1 ~ 끝
        career_programs = []
        for r in range(career_start_row + 1, max_row3 + 1):
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

    # ── 2-라. 주제선택: 프로그램별 학기당 운영 회기 ≥ 2 ──
    for p in topic_programs:
        ok = (p["운영회기"] >= 2)
        results.append((
            "시트3 – 주제선택 회기",
            f"'{p['프로그램명']}' 운영 회기 ≥ 2회",
            "✅ 통과" if ok else "❌ 실패",
            f"운영 회기 = {p['운영회기']}회 (행 {p['row']})"
        ))

    # ── 2-나. 주제선택: 총 운영 시수 합 ≥ D8 ──
    topic_hours_sum = sum(p["총운영시수"] for p in topic_programs)
    check2b = (topic_hours_sum >= d8)
    results.append((
        "시트3 – 주제선택 시수",
        "주제선택 총 운영 시수 합 ≥ 주제선택 시수(D8)",
        "✅ 통과" if check2b else "❌ 실패",
        f"시수 합 = {topic_hours_sum}  /  D8 = {d8}"
    ))

    # ── 2-다. 진로탐색: 총 운영 시수 합 == D9 ──
    career_hours_sum = sum(p["총운영시수"] for p in career_programs)
    check2c = (career_hours_sum == d9)
    results.append((
        "시트3 – 진로탐색 시수",
        "진로탐색 총 운영 시수 합 = 진로탐색 시수(D9)",
        "✅ 통과" if check2c else "❌ 실패",
        f"시수 합 = {career_hours_sum}  /  D9 = {d9}"
    ))

    # ───── 3. 시트4 예산 점검 ─────
    max_row4 = ws4.max_row

    # 데이터 파싱 (6행부터)
    budget_rows = []
    current_category = ""
    for r in range(6, max_row4 + 1):
        a_val = safe_str(ws4.cell(row=r, column=1).value)
        b_val = safe_str(ws4.cell(row=r, column=2).value)
        c_val = safe_str(ws4.cell(row=r, column=3).value)
        d_val_raw = ws4.cell(row=r, column=4).value
        d_val = safe_float(d_val_raw)
        e_val = safe_str(ws4.cell(row=r, column=5).value)

        if a_val:
            current_category = a_val

        if b_val or c_val or d_val_raw is not None:
            budget_rows.append({
                "row": r, "카테고리": current_category,
                "내용": b_val, "산출근거": c_val,
                "소요예산": d_val, "비고": e_val
            })

    # ── 3-라. 개인위탁 존재 시 예산 입력 확인 ──
    all_teachers = [p["지도교사"] for p in topic_programs + career_programs]
    has_personal_consign = any("개인위탁" in t for t in all_teachers)

    if has_personal_consign:
        consign_budget_exists = any(
            "개인위탁" in br["내용"] for br in budget_rows
        )
        results.append((
            "시트4 – 개인위탁",
            "시트3에 '개인위탁' 지도교사 존재 → 예산 입력 확인",
            "✅ 통과" if consign_budget_exists else "❌ 실패",
            "시트4에 프로그램 개인위탁 운영비 " + ("있음" if consign_budget_exists else "없음")
        ))
    else:
        results.append((
            "시트4 – 개인위탁",
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
        "시트4 – 산출근거 입력",
        "모든 내용(B열)에 산출근거(C열) 입력 여부",
        "✅ 통과" if check3b else "❌ 실패",
        "모두 입력됨" if check3b else "미입력: " + " / ".join(missing_calc)
    ))

    # ── 3-바. 산출근거 계산 vs 소요예산 일치 ──
    calc_mismatches = []
    for br in budget_rows:
        if br["산출근거"]:
            calc_result = parse_calc(br["산출근거"])
            if calc_result is not None:
                if abs(calc_result - br["소요예산"]) > 1:  # 1원 오차 허용
                    calc_mismatches.append(
                        f"행{br['row']}: '{br['내용']}' → 계산={calc_result:,.0f}, 예산={br['소요예산']:,.0f}"
                    )
    check3c = len(calc_mismatches) == 0
    results.append((
        "시트4 – 산출근거 일치",
        "산출근거 계산값과 소요예산(D열) 일치 여부",
        "✅ 통과" if check3c else "❌ 실패",
        "모두 일치" if check3c else " / ".join(calc_mismatches)
    ))

    # ── 3-사. 업무추진비 < 총 예산의 3% ──
    total_budget = sum(br["소요예산"] for br in budget_rows)
    admin_budget = sum(br["소요예산"] for br in budget_rows if "업무추진비" in br["카테고리"])

    if total_budget > 0:
        admin_ratio = admin_budget / total_budget * 100
        check3d = (admin_ratio < 3)
        results.append((
            "시트4 – 업무추진비",
            "업무추진비 < 총 예산의 3%",
            "✅ 통과" if check3d else "❌ 실패",
            f"업무추진비 = {admin_budget:,.0f}원 ({admin_ratio:.2f}%)  /  총 예산 = {total_budget:,.0f}원"
        ))
    else:
        results.append((
            "시트4 – 업무추진비",
            "업무추진비 < 총 예산의 3%",
            "⚠️ 경고",
            "총 예산이 0원입니다. 예산 데이터를 확인해주세요."
        ))

    # ── 3-아. 개인위탁 운영비 ≤ 총 예산의 40% ──
    consign_budget = sum(
        br["소요예산"] for br in budget_rows
        if "개인위탁" in br["내용"]
    )
    if total_budget > 0:
        consign_ratio = consign_budget / total_budget * 100
        check3e = (consign_ratio <= 40)
        results.append((
            "시트4 – 개인위탁 비율",
            "프로그램 개인위탁 운영비 ≤ 총 예산의 40%",
            "✅ 통과" if check3e else "❌ 실패",
            f"개인위탁 운영비 = {consign_budget:,.0f}원 ({consign_ratio:.2f}%)  /  총 예산 = {total_budget:,.0f}원"
        ))
    else:
        results.append((
            "시트4 – 개인위탁 비율",
            "프로그램 개인위탁 운영비 ≤ 총 예산의 40%",
            "⚠️ 경고",
            "총 예산이 0원입니다."
        ))

    wb.close()
    return results, warnings


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
        st.markdown(f"## 🏫 {uf.name}")

        try:
            file_copy = io.BytesIO(uf.read())
            uf.seek(0)
            res, warns = inspect_file(file_copy)

            pass_cnt = sum(1 for r in res if "통과" in r[2])
            fail_cnt = sum(1 for r in res if "실패" in r[2])
            warn_cnt = sum(1 for r in res if "경고" in r[2] or "해당없음" in r[2])

            total_pass += pass_cnt
            total_fail += fail_cnt
            total_warn += warn_cnt

            # 학교별 요약 메트릭
            col1, col2, col3 = st.columns(3)
            col1.metric("✅ 통과", f"{pass_cnt}건")
            col2.metric("❌ 실패", f"{fail_cnt}건")
            col3.metric("⚠️ 기타", f"{warn_cnt}건")

            # 결과 테이블
            df = pd.DataFrame(res, columns=["카테고리", "점검 항목", "결과", "상세 내용"])

            # 실패 항목 하이라이트
            def highlight_result(row):
                if "실패" in row["결과"]:
                    return ["background-color: #ffcccc"] * len(row)
                elif "통과" in row["결과"]:
                    return ["background-color: #ccffcc"] * len(row)
                else:
                    return ["background-color: #fff3cd"] * len(row)

            st.dataframe(
                df.style.apply(highlight_result, axis=1),
                use_container_width=True,
                hide_index=True,
                height=min(len(df) * 40 + 40, 600)
            )

            # 실패 항목만 별도 표시
            fail_df = df[df["결과"].str.contains("실패")]
            if not fail_df.empty:
                with st.expander(f"🔴 실패 항목만 보기 ({fail_cnt}건)", expanded=False):
                    st.dataframe(fail_df, use_container_width=True, hide_index=True)

            all_school_results[uf.name] = df

        except Exception as e:
            st.error(f"❌ 파일 처리 중 오류: {e}")
            import traceback
            st.code(traceback.format_exc())

        st.divider()

    # ───── 전체 요약 ─────
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

    # 점검 항목 안내
    with st.expander("📌 점검 항목 안내", expanded=True):
        st.markdown("""
| 번호 | 카테고리 | 점검 내용 |
|:---:|:---:|:---|
| 1 | 시트2-시수 | 주제선택 시수(D8) + 진로탐색 시수(D9) = 자유학기 운영 시수(F6) |
| 2-라 | 시트3-주제선택 | 프로그램별 학기당 운영 회기 ≥ 2회 |
| 2-나 | 시트3-주제선택 | 총 운영 시수 합 ≥ 주제선택 시수(D8) |
| 2-다 | 시트3-진로탐색 | 총 운영 시수 합 = 진로탐색 시수(D9) |
| 3-라 | 시트4-개인위탁 | 개인위탁 지도교사 존재 시 예산 입력 여부 |
| 3-마 | 시트4-산출근거 | 모든 내용에 산출근거 입력 여부 |
| 3-바 | 시트4-계산검증 | 산출근거 계산값 = 소요예산 |
| 3-사 | 시트4-업무추진비 | 업무추진비 < 총 예산의 3% |
| 3-아 | 시트4-개인위탁 비율 | 개인위탁 운영비 ≤ 총 예산의 40% |
        """)

st.markdown("---")
st.caption("© 2026 자유학기 운영계획서 점검 시스템 | Streamlit")
