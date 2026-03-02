# 📋 발주서 자동화 v4.0

발주서 사진을 업로드하면 **Gemini AI**가 바코드·수량을 자동 인식하고,  
마스터 파일과 대조하여 **전산 업로드용 엑셀**을 생성하는 Streamlit 웹 앱입니다.

## ✨ 주요 기능

| 기능 | 설명 |
|---|---|
| **마트 선택** | 와 / 킹 / 팜 셀렉트박스 |
| **사진 업로드** | JPG·PNG 여러 장 동시 업로드 |
| **AI 바코드 인식** | 880·489·693 접두사 13자리 우선, 8·12자리 허용 |
| **검수 편집** | `st.data_editor`로 바코드·수량 직접 수정 |
| **실시간 대조** | 바코드 수정 시 마스터 파일 기준 제품명·단가 자동 반영 |
| **엑셀 다운로드** | 등록 항목만 전산업로드용 엑셀로 즉시 다운로드 |

## 🚀 실행 방법

```bash
# 1. 의존성 설치
pip install -r requirements.txt

# 2. API 키 설정 (.streamlit/secrets.toml)
# GEMINI_API_KEY = "your_api_key_here"

# 3. 실행
streamlit run app.py
```

## 📁 프로젝트 구조

```
발주서 자동화/
├── app.py                  # Streamlit 웹 앱 (메인)
├── requirements.txt        # Python 의존성
├── .streamlit/
│   └── secrets.toml        # API 키 (Git 제외)
├── 기준_와.xlsx             # 와마트 마스터 파일
├── 기준_킹.xlsx             # 킹마트 마스터 파일
├── 기준_팜.xlsx             # 팜마트 마스터 파일
└── .gitignore
```

## 🔧 기술 스택

- **프론트엔드**: Streamlit
- **AI 모델**: Gemini 2.0 Flash Lite (기본) / Gemini 1.5 Flash (폴백)
- **데이터**: pandas, openpyxl
- **API**: Google GenAI

## ⚙️ API 키 설정

### 로컬 실행
`.streamlit/secrets.toml` 파일에 추가:
```toml
GEMINI_API_KEY = "your_api_key_here"
```

### Streamlit Cloud 배포
Settings → Secrets에 동일하게 추가합니다.

> ⚠️ `.streamlit/secrets.toml`과 `.env` 파일은 `.gitignore`에 포함되어 있어 GitHub에 업로드되지 않습니다.
