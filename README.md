# 📋 발주서 자동화

발주서 사진을 업로드하면 **Gemini AI**가 바코드와 수량을 자동 인식하고, 마트별 기준 파일과 대조하여 **전산 업로드용 엑셀**을 생성해주는 프로그램입니다.

## 주요 기능

- 🏪 **마트 선택** — 와마트 / 킹마트 / 팜마트
- 📷 **이미지 분석** — Gemini 2.5 Flash로 바코드(13자리) + 수량 자동 추출
- 📊 **기준 파일 대조** — 자재코드 · 납품단가 자동 매칭
- 📥 **엑셀 다운로드** — 전산 업로드 양식에 맞춘 결과 파일 생성

## 실행 방법

### 1. 패키지 설치

```bash
pip install -r requirements.txt
```

### 2. API 키 설정

`app.py` 상단의 `GEMINI_API_KEY`를 본인의 키로 변경하세요.

> API 키는 [Google AI Studio](https://aistudio.google.com/apikey)에서 무료 발급 가능합니다.

### 3. 기준 파일 준비

아래 3개 파일을 프로젝트 폴더에 넣어주세요:

| 파일명 | 설명 |
|--------|------|
| `기준_와.xlsx` | 와마트 품목 기준표 |
| `기준_킹.xlsx` | 킹마트 품목 기준표 |
| `기준_팜.xlsx` | 팜마트 품목 기준표 |

기준 파일 컬럼: `바코드`, `자재코드`, `제품명`, `단가`

### 4. 웹 앱 실행

```bash
streamlit run app.py
```

브라우저에서 `http://localhost:8501`이 자동으로 열립니다.

### 5. CLI 실행 (선택)

```bash
python extract_order.py
```

## 파일 구조

```
발주서 자동화/
├── app.py                  # Streamlit 웹 앱 (메인)
├── extract_order.py        # CLI 버전
├── requirements.txt        # 패키지 목록
├── .gitignore
├── 기준_와.xlsx            # 마트별 기준 파일
├── 기준_킹.xlsx
├── 기준_팜.xlsx
└── 주문등록 엑셀업로드.xlsx  # 출력 양식 참조
```

## 기술 스택

- **Python 3.12+**
- **Streamlit** — 웹 UI
- **Google Gemini AI** — 이미지 OCR
- **pandas / openpyxl** — 엑셀 처리
