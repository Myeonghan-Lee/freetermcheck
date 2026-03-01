# 📋 자유학기 운영계획서 점검 시스템

2026학년도 중학교 자유학기 운영계획서 엑셀 파일을 업로드하면 자동으로 점검하는 웹앱입니다.

## 🚀 배포 방법

### 1. GitHub 저장소 생성
```bash
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO.git
git push -u origin main
```

### 2. Streamlit Cloud 배포
1. [share.streamlit.io](https://share.streamlit.io) 접속
2. GitHub 계정 연동
3. Repository / Branch / Main file path (`app.py`) 선택
4. **Deploy** 클릭

## 📌 점검 항목
| 번호 | 점검 내용 |
|:---:|:---|
| 1 | 주제선택 시수 + 진로탐색 시수 = 자유학기 운영 시수 |
| 2-라 | 주제선택 프로그램별 운영 회기 ≥ 2회 |
| 2-나 | 주제선택 총 운영 시수 합 ≥ 주제선택 시수 |
| 2-다 | 진로탐색 총 운영 시수 합 = 진로탐색 시수 |
| 3-라 | 개인위탁 지도교사 존재 시 예산 입력 여부 |
| 3-마 | 모든 내용에 산출근거 입력 여부 |
| 3-바 | 산출근거 계산값 = 소요예산 |
| 3-사 | 업무추진비 < 총 예산의 3% |
| 3-아 | 개인위탁 운영비 ≤ 총 예산의 40% |

## 📁 엑셀 파일 구조
- 시트1: 작성안내
- 시트2: 1.학교운영 현황
- 시트3: 2.자유학기 활동
- 시트4: 3.예산 계획서
