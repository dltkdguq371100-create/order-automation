# -*- coding: utf-8 -*-
"""
발주서 자동화 - Streamlit 웹 앱 v5.0 (Groq API)

기능:
  1. 마트 선택 (와 / 킹 / 팜) 라디오 버튼
  2. 발주서 사진 업로드 (st.file_uploader)
  3. Groq API 분석 (llama-3.2-11b-vision-preview)
  4. st.data_editor 기반 검수 편집 [바코드, 수량, 제품명, 단가, 상태]
  5. 바코드/수량 수정 → 마스터 파일 실시간 대조
  6. 최종 확정 및 엑셀 다운로드 버튼

실행: streamlit run app.py
"""

import io
import json
import time
import base64
from pathlib import Path

import streamlit as st
import pandas as pd
import openpyxl
from groq import Groq
from PIL import Image

# ──────────────────────────────────────────────
# ★ 설정 ★
# ──────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent

MART_OPTIONS = {
    "와": "기준_와.xlsx",
    "킹": "기준_킹.xlsx",
    "팜": "기준_팜.xlsx",
}

# Groq 모델: meta-llama/llama-4-scout-17b-16e-instruct (이미지 분석 + JSON 모드 지원)
GROQ_MODEL = "meta-llama/llama-4-scout-17b-16e-instruct"


# ──────────────────────────────────────────────
# 핵심 함수들
# ──────────────────────────────────────────────

@st.cache_resource
def get_groq_client():
    """Groq 클라이언트 (앱 전체에서 1회만 생성)"""
    return Groq(api_key=st.secrets["GROQ_API_KEY"])


@st.cache_data
def load_reference(ref_filename: str) -> dict:
    """기준 파일을 로드하여 바코드→{자재코드, 제품명, 단가} 딕셔너리 반환"""
    ref_path = BASE_DIR / ref_filename
    if not ref_path.exists():
        st.error(f"❌ 기준 파일을 찾을 수 없습니다: `{ref_filename}`")
        st.stop()

    df = pd.read_excel(ref_path)
    ref_dict = {}
    for idx in range(len(df)):
        try:
            raw_val = df.iloc[idx, 0]
            # 소수점 제거: 8.8012345678e+12 또는 '8801234567890.0' 등 처리
            barcode_str = str(raw_val).strip()
            if '.' in barcode_str:
                barcode_str = barcode_str.split('.')[0]
            # 순수 숫자가 아니면 int 변환 후 다시 문자열
            if not barcode_str.isdigit():
                barcode_str = str(int(float(raw_val)))
            barcode_str = barcode_str.strip()
            ref_dict[barcode_str] = {
                "자재코드": int(df.iloc[idx, 1]),
                "제품명": str(df.iloc[idx, 2]).strip(),
                "단가": int(df.iloc[idx, 3]),
            }
        except (ValueError, TypeError):
            continue
    print(f"[DEBUG] 기준파일 '{ref_filename}' 로드 완료: {len(ref_dict)}개 바코드")
    if ref_dict:
        sample_keys = list(ref_dict.keys())[:5]
        print(f"[DEBUG] 바코드 샘플: {sample_keys}")
    return ref_dict


def fix_14digit_barcode(barcode: str) -> str:
    """14자리 바코드에서 중복 숫자 패턴을 제거하여 13자리로 변환"""
    if len(barcode) != 14 or not barcode.isdigit():
        return barcode

    # 연속 3개 동일 숫자 패턴 (999, 000, 111 등) 중 하나를 제거
    for i in range(len(barcode) - 2):
        if barcode[i] == barcode[i + 1] == barcode[i + 2]:
            # 중복 숫자 하나 제거하여 13자리로
            candidate = barcode[:i] + barcode[i + 1:]
            if len(candidate) == 13:
                return candidate

    # 연속 2개 동일 숫자가 있는 위치를 찾아 하나 제거 시도
    for i in range(len(barcode) - 1):
        if barcode[i] == barcode[i + 1]:
            candidate = barcode[:i] + barcode[i + 1:]
            if len(candidate) == 13:
                return candidate

    return barcode


def _barcode_similarity(bc1: str, bc2: str) -> int:
    """두 바코드의 앞자리부터 일치하는 문자 수 반환 (유사도 점수)"""
    score = 0
    for a, b in zip(bc1, bc2):
        if a == b:
            score += 1
        else:
            break
    return score


def _find_by_suffix(barcode: str, ref_dict: dict) -> list:
    """바코드 뒷자리(6~7자리) 기준으로 마스터 파일에서 후보 검색"""
    candidates = []
    for suffix_len in (7, 6):  # 7자리 우선, 그 다음 6자리
        if len(barcode) < suffix_len:
            continue
        suffix = barcode[-suffix_len:]
        for ref_bc in ref_dict:
            if ref_bc.endswith(suffix):
                candidates.append(ref_bc)
        if candidates:
            return candidates
    return candidates


def lookup_barcode(barcode: str, ref_dict: dict) -> dict:
    """바코드 하나를 마스터에서 조회 (스마트 매칭: 정확매칭 → 14자리보정 → 뒷자리 검색)"""
    # 철저한 전처리: str 변환 + strip + 소수점 제거
    barcode = str(barcode).strip()
    if '.' in barcode:
        barcode = barcode.split('.')[0].strip()

    # 1) 정확 일치
    if barcode in ref_dict:
        info = ref_dict[barcode]
        return {"바코드": barcode, "제품명": info["제품명"], "단가": info["단가"], "상태": "✅ 등록"}

    # 2) 14자리 → 13자리 자동 보정
    if len(barcode) == 14 and barcode.isdigit():
        fixed = fix_14digit_barcode(barcode)
        if fixed != barcode and fixed in ref_dict:
            info = ref_dict[fixed]
            return {"바코드": fixed, "제품명": info["제품명"], "단가": info["단가"], "상태": "✅ 자동교정"}

    # 3) 뒷자리(6~7자리) 기준 유사 매칭
    if barcode.isdigit() and len(barcode) >= 8:
        candidates = _find_by_suffix(barcode, ref_dict)

        if len(candidates) == 1:
            matched_bc = candidates[0]
            info = ref_dict[matched_bc]
            return {"바코드": matched_bc, "제품명": info["제품명"], "단가": info["단가"], "상태": "✅ 자동교정"}

        elif len(candidates) > 1:
            best_bc = max(candidates, key=lambda c: _barcode_similarity(barcode, c))
            info = ref_dict[best_bc]
            return {"바코드": best_bc, "제품명": info["제품명"], "단가": info["단가"], "상태": "✅ 자동교정"}

    # 4) 유효 바코드이지만 미등록 — 디버깅 로그 출력
    if barcode.isdigit() and len(barcode) in (8, 12, 13, 14):
        # 유사한 바코드가 마스터에 있는지 앞/뒤 5자리 비교 로그
        similar = [k for k in ref_dict if k[-5:] == barcode[-5:]] if len(barcode) >= 5 else []
        print(f"[DEBUG 미등록] AI값='{barcode}' (길이:{len(barcode)}) | 유사 마스터={similar[:5]}")
        return {"바코드": barcode, "제품명": "", "단가": 0, "상태": "⚠️ 미등록"}

    print(f"[DEBUG 확인필요] AI값='{barcode}' (길이:{len(barcode)}, 숫자여부:{barcode.isdigit()})")
    return {"바코드": barcode, "제품명": "", "단가": 0, "상태": "❓ 바코드 확인"}


def apply_lookup(df: pd.DataFrame, ref_dict: dict) -> pd.DataFrame:
    """데이터프레임 전체에 대해 바코드 기반 제품명/단가/상태를 갱신 (자동교정 시 바코드도 업데이트)"""
    barcodes, names, prices, statuses = [], [], [], []
    for _, row in df.iterrows():
        result = lookup_barcode(row.get("바코드", ""), ref_dict)
        barcodes.append(result["바코드"])
        names.append(result["제품명"])
        prices.append(result["단가"])
        statuses.append(result["상태"])
    df["바코드"] = barcodes
    df["제품명"] = names
    df["단가"] = prices
    df["상태"] = statuses
    return df


def encode_image_to_base64(uploaded_file) -> str:
    """업로드된 이미지를 base64 문자열로 변환"""
    uploaded_file.seek(0)
    return base64.b64encode(uploaded_file.read()).decode("utf-8")


def analyze_image(uploaded_file, max_retries=3):
    """Groq API로 이미지에서 바코드+수량+제품명 추출 (llama-3.2-11b-vision-preview)"""
    client = get_groq_client()

    # 이미지를 base64로 인코딩
    img_base64 = encode_image_to_base64(uploaded_file)

    # 파일 확장자로 MIME 타입 결정
    fname = uploaded_file.name.lower()
    if fname.endswith(".png"):
        mime_type = "image/png"
    else:
        mime_type = "image/jpeg"

    prompt = """이 발주서 이미지를 정밀하게 분석하여 모든 주문 항목을 추출하세요.

[표 구조 분석 규칙]
- 발주서 표는 일반적으로 [번호, 제품명/규격, 바코드, 수량, 단가, 금액] 열로 구성됩니다.
- 각 행에서 제품명, 바코드, 수량을 정확히 한 줄로 매칭하세요.
- 절대로 단가(원), 금액(원), 합계 등 다른 숫자 열을 바코드로 혼동하지 마세요.

[바코드 식별 규칙 - 매우 중요]
- 880으로 시작하면 한국 바코드 (가장 흔함)
- 489로 시작하면 홍콩 바코드
- 693으로 시작하면 중국 바코드
- 위 접두사로 시작하는 13자리 숫자를 최우선으로 바코드로 인식하세요.
- 13자리가 아니더라도, 8자리 또는 12자리 숫자 코드도 바코드로 허용합니다.
- 4~5자리 숫자는 단가, 5~7자리 숫자는 금액일 가능성이 높으니 바코드와 구별하세요.

[수량 규칙]
- 수량은 보통 1~999 범위의 작은 정수입니다.
- 단가나 금액과 혼동하지 마세요.

[응답 형식]
반드시 아래 형식의 JSON 객체로만 응답하세요. 마크다운이나 다른 설명은 절대 포함하지 마세요:

{"items": [{"product_name": "제품명/규격", "barcode": "8801234567890", "qty": 3}, {"product_name": "제품명/규격", "barcode": "8809876543210", "qty": 1}]}"""

    # ── API 호출 전 2초 대기 (할당량 초과 방지) ──
    time.sleep(2)

    last_error = None
    for attempt in range(1, max_retries + 1):
        try:
            response = client.chat.completions.create(
                model=GROQ_MODEL,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "text",
                                "text": prompt,
                            },
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:{mime_type};base64,{img_base64}",
                                },
                            },
                        ],
                    }
                ],
                temperature=0.1,
                max_tokens=4096,
                response_format={"type": "json_object"},
            )
            text = response.choices[0].message.content.strip()

            # JSON 추출 (JSON 모드이므로 순수 JSON 반환, 폴백으로 코드블록 처리)
            if "```json" in text:
                text = text.split("```json")[1].split("```")[0].strip()
            elif "```" in text:
                text = text.split("```")[1].split("```")[0].strip()

            parsed = json.loads(text)
            # JSON 모드: {"items": [...]} 형태 또는 직접 배열 [...] 둘 다 지원
            if isinstance(parsed, dict):
                items = parsed.get("items", [])
            else:
                items = parsed

            results = []
            warnings = []
            for item in items:
                product_name = str(item.get("product_name", "")).strip()
                barcode = str(item.get("barcode", "")).strip()
                qty = item.get("qty", 0)

                # 바코드 유효성: 8, 12, 13, 14자리 숫자 허용 (14자리는 후처리에서 자동보정)
                if not barcode.isdigit() or len(barcode) not in (8, 12, 13, 14):
                    warnings.append(f"'{barcode}' ({len(barcode)}자리) - {product_name}")
                    continue

                try:
                    qty = int(qty)
                except (ValueError, TypeError):
                    qty = 0

                results.append({
                    "바코드": barcode,
                    "수량": qty,
                    "제품명": product_name,
                })

            return results, warnings

        except Exception as e:
            last_error = e
            error_msg = str(e).lower()

            if attempt < max_retries:
                # 429 할당량 초과 → 길게 대기
                if "429" in error_msg or "rate_limit" in error_msg:
                    wait = 30 * attempt
                    st.warning(f"⏳ API 할당량 초과. {wait}초 후 재시도... ({attempt}/{max_retries})")
                    time.sleep(wait)
                # 기타 에러 → 10초 대기 후 재시도
                else:
                    st.warning(f"⚠️ 오류 발생: {e}. 10초 후 재시도... ({attempt}/{max_retries})")
                    time.sleep(10)
            else:
                st.error(f"❌ 최대 재시도 초과. 분석 실패: {last_error}")

    return [], []


def create_excel_bytes(final_df: pd.DataFrame, ref_dict: dict):
    """확정된 데이터로 전산업로드용 엑셀 생성 (등록된 항목만 포함)"""
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

    row_num = 3
    for _, row in final_df.iterrows():
        barcode = str(row.get("바코드", "")).strip()
        if barcode in ref_dict:
            info = ref_dict[barcode]
            ws.cell(row=row_num, column=1, value=info["자재코드"])
            ws.cell(row=row_num, column=2, value=int(row["수량"]))
            ws.cell(row=row_num, column=4, value=info["단가"])
            row_num += 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue(), row_num - 3


# ──────────────────────────────────────────────
# Streamlit UI
# ──────────────────────────────────────────────

st.set_page_config(page_title="발주서 자동화", page_icon="📋", layout="wide")

st.markdown("""
<style>
    .main-header {
        text-align: center;
        padding: 1.2rem 0;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 12px;
        color: white;
        margin-bottom: 2rem;
    }
    .main-header h1 { color: white; margin: 0; font-size: 2rem; }
    .main-header p  { color: rgba(255,255,255,0.85); margin: 0.4rem 0 0 0; }
    .stRadio > div { display: flex; gap: 1rem; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h1>📋 발주서 자동화 v4.0</h1>
    <p>사진 업로드 → AI 분석 → 검수 편집 → 전산 엑셀 다운로드</p>
</div>
""", unsafe_allow_html=True)

# ── 1. 마트 선택 ──
st.subheader("1️⃣ 마트 선택")
selected_mart = st.selectbox(
    "발주서를 보낸 마트를 선택하세요",
    list(MART_OPTIONS.keys()),
)

st.divider()

# ── 2. 발주서 사진 업로드 ──
st.subheader("2️⃣ 발주서 사진 업로드")
uploaded_files = st.file_uploader(
    "발주서 이미지를 선택하세요 (여러 장 가능)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

if uploaded_files:
    cols = st.columns(min(len(uploaded_files), 4))
    for i, f in enumerate(uploaded_files):
        with cols[i % len(cols)]:
            st.image(f, caption=f.name, use_container_width=True)

st.divider()

# ── 3. AI 분석 ──
st.subheader("3️⃣ AI 분석")

if uploaded_files:
    if st.button("🚀 분석 시작", type="primary", use_container_width=True):
        ref_file = MART_OPTIONS[selected_mart]
        ref_dict = load_reference(ref_file)

        all_results = []
        all_warnings = []

        progress = st.progress(0, text="분석 준비 중...")
        for i, f in enumerate(uploaded_files):
            progress.progress(
                i / len(uploaded_files),
                text=f"🔄 '{f.name}' 분석 중... ({i + 1}/{len(uploaded_files)})",
            )
            f.seek(0)
            results, warnings = analyze_image(f)
            all_results.extend(results)
            all_warnings.extend(warnings)

        progress.progress(1.0, text="✅ 분석 완료!")

        if not all_results:
            st.error("❌ 추출된 데이터가 없습니다. 이미지를 확인해 주세요.")
        else:
            df = pd.DataFrame(all_results, columns=["바코드", "수량", "제품명"])
            df["단가"] = 0
            df["상태"] = ""
            df = apply_lookup(df, ref_dict)

            st.session_state["ocr_df"] = df
            st.session_state["warnings"] = all_warnings
            st.session_state["selected_mart_for_review"] = selected_mart
            st.session_state["analysis_done"] = True
            st.success(f"✅ {len(all_results)}개 항목 추출 완료! 아래에서 검수하세요.")
            st.rerun()
else:
    st.info("👆 위에서 발주서 이미지를 업로드하면 분석을 시작할 수 있습니다.")

# ── 4. 검수용 편집 표 ──
if st.session_state.get("analysis_done"):
    st.divider()
    st.subheader("4️⃣ 검수 및 편집")

    # 경고 표시
    warnings = st.session_state.get("warnings", [])
    if warnings:
        with st.expander(f"🔔 바코드 인식 경고 ({len(warnings)}건)", expanded=False):
            for w in warnings:
                st.write(f"  • {w}")

    st.markdown(
        "> 💡 **사용법**: 셀을 **더블클릭**하여 바코드·수량을 수정하세요. "
        "수정 후 **'바코드 대조 갱신'** 버튼을 누르면 제품명·단가가 업데이트됩니다."
    )

    # 기준 파일 로드
    review_mart = st.session_state.get("selected_mart_for_review", selected_mart)
    ref_file = MART_OPTIONS[review_mart]
    ref_dict = load_reference(ref_file)

    df = st.session_state["ocr_df"].copy()

    # 컬럼 순서 보장: [바코드, 수량, 제품명, 단가, 상태]
    for col in ["바코드", "수량", "제품명", "단가", "상태"]:
        if col not in df.columns:
            df[col] = "" if col in ("바코드", "제품명", "상태") else 0
    df = df[["바코드", "수량", "제품명", "단가", "상태"]]

    edited_df = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=False,
        column_config={
            "바코드": st.column_config.TextColumn(
                "바코드",
                help="13자리(우선), 8자리, 12자리 숫자. 수정 가능",
                width="medium",
            ),
            "수량": st.column_config.NumberColumn(
                "수량",
                help="주문 수량. 수정 가능",
                min_value=0,
                max_value=9999,
                step=1,
                width="small",
            ),
            "제품명": st.column_config.TextColumn(
                "제품명",
                help="마스터 파일 기준 제품명 (자동 반영)",
                disabled=True,
                width="large",
            ),
            "단가": st.column_config.NumberColumn(
                "단가",
                help="마스터 파일 기준 단가 (자동 반영)",
                disabled=True,
                width="small",
            ),
            "상태": st.column_config.TextColumn(
                "상태",
                help="마스터 파일 대조 결과",
                disabled=True,
                width="small",
            ),
        },
        key="order_editor",
    )

    # 바코드 대조 갱신 버튼
    if st.button("🔄 바코드 대조 갱신", use_container_width=True):
        updated = edited_df[["바코드", "수량"]].copy()
        updated["제품명"] = ""
        updated["단가"] = 0
        updated["상태"] = ""
        updated = apply_lookup(updated, ref_dict)
        st.session_state["ocr_df"] = updated
        st.rerun()

    # 요약 통계
    total = len(edited_df)
    matched = len(edited_df[edited_df["상태"] == "✅ 등록"]) if "상태" in edited_df.columns else 0
    unmatched = len(edited_df[edited_df["상태"] == "⚠️ 미등록"]) if "상태" in edited_df.columns else 0
    unknown = total - matched - unmatched

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📦 전체", f"{total}개")
    c2.metric("✅ 등록", f"{matched}개")
    c3.metric("⚠️ 미등록", f"{unmatched}개")
    c4.metric("❓ 확인필요", f"{unknown}개")

    # ── 5. 최종 확정 및 엑셀 다운로드 ──
    st.divider()
    st.subheader("5️⃣ 최종 확정 및 엑셀 다운로드")

    if matched == 0:
        st.warning("⚠️ 등록된 항목이 없습니다. 바코드를 확인해 주세요.")

    st.info(
        f"✅ **{matched}개** 등록 항목이 엑셀에 포함됩니다."
        + (f" ⚠️ **{unmatched}개** 미등록 항목은 제외됩니다." if unmatched > 0 else "")
    )

    if st.button(
        "📥 최종 확정 및 엑셀 다운로드",
        type="primary",
        use_container_width=True,
        disabled=(matched == 0),
    ):
        excel_bytes, count = create_excel_bytes(edited_df, ref_dict)
        st.session_state["excel_bytes"] = excel_bytes
        st.session_state["excel_count"] = count
        st.success(f"✅ 전산 엑셀 생성 완료! ({count}개 품목)")

    if st.session_state.get("excel_bytes"):
        st.download_button(
            label=f"💾 전산업로드용_결과.xlsx 다운로드 ({st.session_state.get('excel_count', 0)}개 품목)",
            data=st.session_state["excel_bytes"],
            file_name="전산업로드용_결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # 초기화
    st.divider()
    if st.button("🗑️ 분석 결과 초기화 (새로 시작)", use_container_width=True):
        for key in ["ocr_df", "warnings", "analysis_done", "excel_bytes",
                     "excel_count", "selected_mart_for_review"]:
            st.session_state.pop(key, None)
        st.rerun()

# 푸터
st.markdown("---")
st.caption("발주서 자동화 v5.0 | Groq AI (Llama Vision) 바코드 인식 + 수동 검수 + 전산 엑셀 생성")
