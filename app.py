# -*- coding: utf-8 -*-
"""
발주서 자동화 - Streamlit 웹 앱

실행: streamlit run app.py
"""

import os
import json
import io
import time
from pathlib import Path

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # Streamlit Cloud에서는 dotenv 불필요 (secrets 사용)

import streamlit as st
import pandas as pd
import openpyxl
from google import genai
from PIL import Image

# ──────────────────────────────────────────────
# ★ 설정 ★
# ──────────────────────────────────────────────
# API 키: Streamlit secrets → .env / 환경변수 순서로 탐색
def _get_api_key():
    try:
        return st.secrets["GEMINI_API_KEY"]
    except Exception:
        key = os.environ.get("GEMINI_API_KEY")
        if not key:
            st.error("❌ GEMINI_API_KEY가 설정되지 않았습니다.\n\n"
                     "**로컬**: `.env` 파일에 `GEMINI_API_KEY=your_key` 추가\n\n"
                     "**Streamlit Cloud**: Settings → Secrets에 `GEMINI_API_KEY = \"your_key\"` 추가")
            st.stop()
        return key

GEMINI_API_KEY = _get_api_key()
BASE_DIR = Path(__file__).resolve().parent

MART_OPTIONS = {
    "와마트": "기준_와.xlsx",
    "킹마트": "기준_킹.xlsx",
    "팜마트": "기준_팜.xlsx",
}


# ──────────────────────────────────────────────
# 핵심 함수들
# ──────────────────────────────────────────────

@st.cache_resource
def get_genai_client():
    """Gemini 클라이언트 (앱 전체에서 1회만 생성)"""
    return genai.Client(api_key=GEMINI_API_KEY)


@st.cache_data
def load_reference(ref_filename):
    """기준 파일을 로드하여 바코드→{자재코드, 단가} 딕셔너리 반환"""
    ref_path = BASE_DIR / ref_filename
    if not ref_path.exists():
        st.error(f"❌ 기준 파일을 찾을 수 없습니다: `{ref_filename}`\n\n"
                 f"프로젝트 폴더에 해당 파일이 있는지 확인하세요.")
        st.stop()
    df = pd.read_excel(ref_path)
    # 컬럼 순서: [0]바코드, [1]자재코드, [2]제품명, [3]단가
    ref_dict = {}
    for idx in range(len(df)):
        barcode_str = str(int(df.iloc[idx, 0]))
        ref_dict[barcode_str] = {
            "자재코드": int(df.iloc[idx, 1]),
            "단가": int(df.iloc[idx, 3]),
        }
    return ref_dict


def analyze_image(uploaded_file, max_retries=3):
    """Gemini API로 이미지에서 바코드+수량 추출"""
    client = get_genai_client()
    img = Image.open(uploaded_file)

    prompt = ("이 발주서 이미지에서 바코드(상품코드)와 수량만 추출해줘. "
              "바코드는 반드시 13자리 숫자여야 해. 수량은 정수로 추출해. "
              "제품명, 단가, 합계 등은 무시해. "
              '반드시 JSON 배열로만 응답해: [{"barcode": "1234567890123", "qty": 3}] '
              "JSON만 응답하고 다른 설명은 하지 마.")

    for attempt in range(1, max_retries + 1):
        try:
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=[prompt, img],
            )
            text = response.text.strip()

            if "```json" in text:
                text = text.split("```json")[1].split("```")[0].strip()
            elif "```" in text:
                text = text.split("```")[1].split("```")[0].strip()

            items = json.loads(text)

            results = []
            warnings = []
            for item in items:
                barcode = str(item.get("barcode", "")).strip()
                qty = item.get("qty", 0)

                if not barcode.isdigit() or len(barcode) != 13:
                    warnings.append(f"'{barcode}' ({len(barcode)}자리)")
                    continue

                try:
                    qty = int(qty)
                except (ValueError, TypeError):
                    qty = 0

                results.append((barcode, qty))

            return results, warnings

        except Exception as e:
            error_msg = str(e)
            if "429" in error_msg and attempt < max_retries:
                wait = 60 * attempt
                st.warning(f"⏳ API 할당량 초과. {wait}초 후 재시도... ({attempt}/{max_retries})")
                time.sleep(wait)
            else:
                st.error(f"❌ API 오류: {e}")
                return [], []

    return [], []


def match_with_reference(order_data, ref_dict):
    """바코드를 기준 파일과 대조"""
    matched = []
    unmatched = []

    for barcode, qty in order_data:
        if barcode in ref_dict:
            info = ref_dict[barcode]
            matched.append({
                "자재코드": info["자재코드"],
                "BOX": qty,
                "EA": "",
                "단가": info["단가"],
            })
        else:
            unmatched.append({"바코드": barcode, "수량": qty, "상태": "❌ 확인 필요"})

    return matched, unmatched


def create_excel_bytes(matched_data):
    """전산업로드용 엑셀을 메모리에 생성하여 bytes 반환"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # 헤더 1행
    ws.cell(row=1, column=1, value="자재코드")
    ws.cell(row=1, column=2, value="수량")
    ws.cell(row=1, column=4, value="단가")

    # 헤더 2행
    ws.cell(row=2, column=2, value="BOX")
    ws.cell(row=2, column=3, value="EA")

    # 데이터
    for i, item in enumerate(matched_data, start=3):
        ws.cell(row=i, column=1, value=item["자재코드"])
        ws.cell(row=i, column=2, value=item["BOX"])
        ws.cell(row=i, column=4, value=item["단가"])

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ──────────────────────────────────────────────
# Streamlit UI
# ──────────────────────────────────────────────

st.set_page_config(page_title="발주서 자동화", page_icon="📋", layout="wide")

# 커스텀 CSS
st.markdown("""
<style>
    .main-header {
        text-align: center;
        padding: 1rem 0;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 10px;
        color: white;
        margin-bottom: 2rem;
    }
    .main-header h1 { color: white; margin: 0; }
    .main-header p { color: rgba(255,255,255,0.8); margin: 0.5rem 0 0 0; }
    .stRadio > div { display: flex; gap: 1rem; }
    .metric-card {
        background: #f8f9fa;
        border-radius: 8px;
        padding: 1rem;
        text-align: center;
        border: 1px solid #e9ecef;
    }
</style>
""", unsafe_allow_html=True)

# 헤더
st.markdown("""
<div class="main-header">
    <h1>📋 발주서 자동화</h1>
    <p>발주서 사진 → 전산 업로드용 엑셀 자동 생성</p>
</div>
""", unsafe_allow_html=True)

# ── Step 1: 마트 선택 ──
st.subheader("1️⃣ 마트 선택")
selected_mart = st.radio(
    "발주서를 보낸 마트를 선택하세요",
    list(MART_OPTIONS.keys()),
    horizontal=True,
)

st.divider()

# ── Step 2: 이미지 업로드 ──
st.subheader("2️⃣ 발주서 사진 업로드")
uploaded_files = st.file_uploader(
    "발주서 이미지를 선택하세요 (여러 장 가능)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

# 업로드된 이미지 미리보기
if uploaded_files:
    cols = st.columns(min(len(uploaded_files), 4))
    for i, f in enumerate(uploaded_files):
        with cols[i % 4]:
            st.image(f, caption=f.name, use_container_width=True)

st.divider()

# ── Step 3: 분석 실행 ──
st.subheader("3️⃣ 분석 실행")

if uploaded_files:
    if st.button("🚀 분석 시작", type="primary", use_container_width=True):
        ref_file = MART_OPTIONS[selected_mart]
        ref_dict = load_reference(ref_file)

        all_orders = []
        all_warnings = []

        # 이미지별 분석
        progress = st.progress(0, text="분석 준비 중...")
        for i, f in enumerate(uploaded_files):
            progress.progress(
                (i) / len(uploaded_files),
                text=f"🔄 '{f.name}' 분석 중... ({i+1}/{len(uploaded_files)})"
            )
            f.seek(0)
            results, warnings = analyze_image(f)
            all_orders.extend(results)
            all_warnings.extend(warnings)

        progress.progress(1.0, text="✅ 분석 완료!")

        if not all_orders:
            st.error("❌ 추출된 데이터가 없습니다. 이미지를 확인해 주세요.")
        else:
            # 매칭
            matched, unmatched = match_with_reference(all_orders, ref_dict)

            # 결과 요약 카드
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("📷 분석 이미지", f"{len(uploaded_files)}장")
            col2.metric("🔍 추출 바코드", f"{len(all_orders)}개")
            col3.metric("✅ 매칭 성공", f"{len(matched)}개")
            col4.metric("⚠️ 확인 필요", f"{len(unmatched)}개")

            st.divider()

            # ── 매칭 결과 표 ──
            if matched:
                st.subheader("📊 매칭 결과")
                df_matched = pd.DataFrame(matched)
                st.dataframe(
                    df_matched,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "자재코드": st.column_config.NumberColumn("자재코드", format="%d"),
                        "BOX": st.column_config.NumberColumn("BOX (수량)"),
                        "EA": st.column_config.TextColumn("EA"),
                        "단가": st.column_config.NumberColumn("단가", format="%d원"),
                    },
                )

            # ── 확인 필요 항목 ──
            if unmatched:
                st.subheader("⚠️ 확인 필요 (기준 파일에 없는 바코드)")
                df_unmatched = pd.DataFrame(unmatched)
                st.dataframe(df_unmatched, use_container_width=True, hide_index=True)

            # ── 바코드 자릿수 경고 ──
            if all_warnings:
                with st.expander(f"🔔 바코드 자릿수 오류 ({len(all_warnings)}건)"):
                    for w in all_warnings:
                        st.write(f"  • {w}")

            st.divider()

            # ── 엑셀 다운로드 ──
            if matched:
                st.subheader("4️⃣ 전산 파일 다운로드")
                excel_bytes = create_excel_bytes(matched)

                st.download_button(
                    label="📥 전산업로드용_결과.xlsx 다운로드",
                    data=excel_bytes,
                    file_name="전산업로드용_결과.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )
else:
    st.info("👆 위에서 발주서 이미지를 업로드하면 분석을 시작할 수 있습니다.")

# 푸터
st.markdown("---")
st.caption("발주서 자동화 v2.0 | Gemini AI 기반 바코드 인식")
