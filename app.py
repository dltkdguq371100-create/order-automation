# -*- coding: utf-8 -*-
"""
발주서 자동화 - Streamlit 웹 앱 (Groq API 최적화)
- 만능 JSON 파서 및 한글 키 대응 방어 로직 추가
- 정규식을 통한 바코드/수량 불순물(글자) 완벽 제거
- 응답이 잘린 불완전 JSON의 강제 복구 로직 탑재
- 품목명 역방향 매칭 강화
"""

import io
import json
import re
import time
import base64
from pathlib import Path

import streamlit as st
import pandas as pd
import openpyxl
from groq import Groq

# ──────────────────────────────────────────────
# ★ 설정 ★
# ──────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent

MART_OPTIONS = {
    "와": "기준_와.xlsx",
    "킹": "기준_킹.xlsx",
    "팜": "기준_팜.xlsx",
}

# Groq 최신 비전 모델 (공식 지원 명칭)
GROQ_MODEL = "meta-llama/llama-4-scout-17b-16e-instruct"

# ──────────────────────────────────────────────
# 핵심 함수들
# ──────────────────────────────────────────────

@st.cache_resource
def get_groq_client():
    return Groq(api_key=st.secrets["GROQ_API_KEY"])

@st.cache_data
def load_reference(ref_filename: str) -> dict:
    ref_path = BASE_DIR / ref_filename
    if not ref_path.exists():
        st.error(f"❌ 기준 파일을 찾을 수 없습니다: `{ref_filename}`")
        st.stop()

    df = pd.read_excel(ref_path)
    ref_dict = {}
    for idx in range(len(df)):
        try:
            raw_val = df.iloc[idx, 0]
            barcode_str = str(raw_val).strip()
            if '.' in barcode_str:
                barcode_str = barcode_str.split('.')[0]
            if not barcode_str.isdigit():
                barcode_str = str(int(float(raw_val)))
            barcode_str = barcode_str.strip()
            
            # 제품명 공백 제거 및 소문자화
            product_name = str(df.iloc[idx, 2]).strip()
            
            ref_dict[barcode_str] = {
                "자재코드": int(df.iloc[idx, 1]),
                "제품명": product_name,
                "단가": int(df.iloc[idx, 3]),
            }
        except (ValueError, TypeError):
            continue
    return ref_dict

def fix_14digit_barcode(barcode: str) -> str:
    if len(barcode) != 14 or not barcode.isdigit():
        return barcode
    for i in range(len(barcode) - 2):
        if barcode[i] == barcode[i + 1] == barcode[i + 2]:
            candidate = barcode[:i] + barcode[i + 1:]
            if len(candidate) == 13: return candidate
    for i in range(len(barcode) - 1):
        if barcode[i] == barcode[i + 1]:
            candidate = barcode[:i] + barcode[i + 1:]
            if len(candidate) == 13: return candidate
    return barcode

def _barcode_similarity(bc1: str, bc2: str) -> int:
    score = 0
    for a, b in zip(bc1, bc2):
        if a == b: score += 1
        else: break
    return score

def _find_by_suffix(barcode: str, ref_dict: dict) -> list:
    candidates = []
    for suffix_len in (7, 6):
        if len(barcode) < suffix_len: continue
        suffix = barcode[-suffix_len:]
        for ref_bc in ref_dict:
            if ref_bc.endswith(suffix):
                candidates.append(ref_bc)
        if candidates: return candidates
    return candidates

def lookup_by_product_name(product_name: str, ref_dict: dict) -> dict | None:
    """품목명 키워드로 마스터 파일에서 바코드를 역으로 찾는 함수"""
    if not product_name or len(product_name.strip()) < 2:
        return None

    name_clean = product_name.strip().lower()
    best_match = None
    best_score = 0

    clean_words = re.sub(r'[^가-힣a-z0-9]', ' ', name_clean).split()
    keywords = [w for w in clean_words if len(w) >= 2]
    
    if not keywords:
        return None

    for ref_bc, info in ref_dict.items():
        ref_name = info["제품명"].lower()
        matched = sum(1 for kw in keywords if kw in ref_name)
        score = matched / len(keywords)

        # 35% 이상 일치 시 후보로 허용 (킹마트 환각 대응)
        if score > best_score and score >= 0.35:
            best_score = score
            best_match = {
                "바코드": ref_bc,
                "제품명": info["제품명"],
                "단가": info["단가"],
                "점수": best_score,
            }

    return best_match

def lookup_barcode(barcode: str, ref_dict: dict, product_name: str = "") -> dict:
    barcode = str(barcode).strip()
    if '.' in barcode: barcode = barcode.split('.')[0].strip()

    # 1) 정확 일치
    if barcode in ref_dict:
        info = ref_dict[barcode]
        return {"바코드": barcode, "제품명": info["제품명"], "단가": info["단가"], "상태": "✅ 등록"}

    # 2) 14자리 자동 보정
    if len(barcode) == 14 and barcode.isdigit():
        fixed = fix_14digit_barcode(barcode)
        if fixed != barcode and fixed in ref_dict:
            info = ref_dict[fixed]
            return {"바코드": fixed, "제품명": info["제품명"], "단가": info["단가"], "상태": "✅ 자동교정"}

    # 3) 뒷자리 매칭
    if barcode.isdigit() and len(barcode) >= 8:
        candidates = _find_by_suffix(barcode, ref_dict)
        if len(candidates) == 1:
            return {"바코드": candidates[0], "제품명": ref_dict[candidates[0]]["제품명"], "단가": ref_dict[candidates[0]]["단가"], "상태": "✅ 자동교정"}
        elif len(candidates) > 1:
            best_bc = max(candidates, key=lambda c: _barcode_similarity(barcode, c))
            return {"바코드": best_bc, "제품명": ref_dict[best_bc]["제품명"], "단가": ref_dict[best_bc]["단가"], "상태": "✅ 자동교정"}

    # 4) 역방향 매칭 (바코드가 틀렸거나 없을 때 제품명으로 강력하게 유추)
    if product_name:
        name_match = lookup_by_product_name(product_name, ref_dict)
        if name_match:
            return {"바코드": name_match["바코드"], "제품명": name_match["제품명"], "단가": name_match["단가"], "상태": "✅ 품명매칭"}

    return {"바코드": barcode, "제품명": "", "단가": 0, "상태": "⚠️ 미등록"}

def apply_lookup(df: pd.DataFrame, ref_dict: dict) -> pd.DataFrame:
    barcodes, names, prices, statuses = [], [], [], []
    for _, row in df.iterrows():
        product_name = str(row.get("제품명", "")).strip()
        result = lookup_barcode(row.get("바코드", ""), ref_dict, product_name=product_name)
        barcodes.append(result["바코드"])
        names.append(result["제품명"])
        prices.append(result["단가"])
        statuses.append(result["상태"])
    df["바코드"] = barcodes
    df["제품명"] = names
    df["단가"] = prices
    df["상태"] = statuses
    return df

def get_prompt_for_mart(mart_type: str) -> str:
    base = """[역할]
당신은 물류 센터의 데이터 입력 AI입니다.
이미지의 표 데이터를 빠짐없이 정확하게 읽어 JSON으로 변환하세요.
- 반드시 아래 형식의 JSON 객체로 반환하세요.
- 키 이름은 무조건 영어로 고정하세요: "name" (상품명), "barcode" (바코드), "qty" (수량).
- 표의 마지막 행까지 읽었으면 즉시 생성을 종료하세요. 같은 항목을 절대 반복하지 마세요.
{"items": [{"name": "상품명", "barcode": "8801234567890", "qty": 1}]}
"""
    if mart_type == "킹":
        return base + """
[킹마트 발주서 특별 지침]
1. 2열(상품명)과 3열(규격)의 텍스트를 하나로 합쳐서 "name"에 적으세요.
2. 4열(바코드)의 13자리 숫자를 "barcode"에 적으세요. 바코드가 지워졌거나 헷갈리면 지어내지 말고 비워두세요("").
3. 5열(발주량)의 숫자를 "qty"에 적으세요. 크게 그려진 볼펜 동그라미(O)는 0이 아니니 무시하고, 그 안의 실제 숫자만 적으세요.
"""
    elif mart_type == "팜":
        return base + """
[팜마트 발주서 특별 지침]
1. 2번째 열('코드')이 바코드입니다. "barcode"에 적으세요.
2. 5번째 열이 발주량입니다. "qty"에 적으세요.
3. 바로 옆의 품목명을 "name"에 적으세요. 선이 촘촘하니 행이 섞이지 않게 주의하세요.
"""
    else:
        return base + """
[와마트 발주서 분석 지침]
- 880, 489, 693으로 시작하는 13자리 숫자가 바코드야. 단가나 금액 열과 혼동하지 마.
"""

def force_parse_json(text: str):
    """(핵심!) AI 응답이 중간에 끊기거나 망가져도 강제로 괄호를 닫아 JSON을 살려내는 함수"""
    text = text.strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass # 파싱 실패 시 아래 복구 로직 실행
    
    # 불필요한 쉼표 등 찌꺼기 제거
    if text.endswith(','): 
        text = text[:-1]
    
    # 열린 괄호 개수 추적해서 강제로 닫아버리기
    open_braces = text.count('{') - text.count('}')
    open_brackets = text.count('[') - text.count(']')
    
    if open_braces > 0: text += '}' * open_braces
    if open_brackets > 0: text += ']' * open_brackets
    
    try:
        return json.loads(text)
    except Exception:
        return None # 도저히 복구 불가일 때만 None 반환

def extract_list_from_json(parsed_data):
    """JSON 안의 어떤 키에 배열이 숨어있든 샅샅이 뒤져서 찾아내는 함수"""
    if isinstance(parsed_data, list):
        return parsed_data
    if isinstance(parsed_data, dict):
        # AI가 멋대로 쓸 수 있는 모든 키값 대비
        for key in ["items", "orders", "data", "rows", "목록", "발주서", "result"]:
            if key in parsed_data and isinstance(parsed_data[key], list):
                return parsed_data[key]
        for key, value in parsed_data.items():
            if isinstance(value, list):
                return value
    return []

def analyze_image(uploaded_file, mart_type="와"):
    client = get_groq_client()
    uploaded_file.seek(0)
    img_base64 = base64.b64encode(uploaded_file.read()).decode("utf-8")
    mime_type = "image/png" if uploaded_file.name.lower().endswith(".png") else "image/jpeg"

    prompt = get_prompt_for_mart(mart_type)

    try:
        response = client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[
                {
                    "role": "user", 
                    "content": [
                        {"type": "text", "text": prompt}, 
                        {"type": "image_url", "image_url": {"url": f"data:{mime_type};base64,{img_base64}"}}
                    ]
                }
            ],
            temperature=0.1, 
            max_tokens=4096, 
            response_format={"type": "json_object"}
        )
        text = response.choices[0].message.content.strip()

        # 마크다운 찌꺼기 제거
        if "```json" in text: 
            text = text.split("```json")[1].split("```")[0].strip()
        elif "```" in text: 
            text = text.split("```")[1].split("```")[0].strip()
        
        # 제어문자 정화
        text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', text)

        # ✅ 망가진 JSON 강제 복구 파싱
        parsed = force_parse_json(text)
        if parsed is None:
            return [], ["AI 응답 형식이 완전히 망가져서 복구할 수 없습니다."]

        items = extract_list_from_json(parsed)

        results = []
        warnings = []
        prev_hash = "" 

        for item in items:
            if not isinstance(item, dict): continue

            # ✅ AI의 '한글 키' 변덕 완벽 방어
            product_name = str(item.get("name", item.get("product_name", item.get("상품명", item.get("품명", item.get("제품명", "")))))).strip()
            barcode_raw = str(item.get("barcode", item.get("code", item.get("바코드", item.get("상품코드", "")))))
            qty_raw = str(item.get("qty", item.get("quantity", item.get("amount", item.get("수량", item.get("발주량", 0))))))

            # ✅ 글자가 섞여 들어오는 현상 방어 (정규식으로 오직 숫자만 싹둑 추출)
            barcode = re.sub(r'[^0-9]', '', barcode_raw)
            qty_str = re.sub(r'[^0-9]', '', qty_raw)

            # 바코드와 제품명이 둘 다 없으면 쓸모없는 행이므로 패스
            if not barcode and not product_name: continue 

            try:
                qty = int(qty_str) if qty_str else 0
            except:
                qty = 0

            # 무한 루프 물리적 차단 로직
            current_hash = f"{barcode}_{product_name}_{qty}"
            if current_hash == prev_hash:
                continue 
            prev_hash = current_hash

            if barcode and (not barcode.isdigit() or len(barcode) not in (8, 12, 13, 14)):
                warnings.append(f"형식 의심: '{barcode}' - {product_name}")

            results.append({"바코드": barcode, "수량": qty, "제품명": product_name})

        return results, warnings

    except Exception as e:
        st.error(f"❌ 분석 중 에러 발생: {e}")
        return [], []

def create_excel_bytes(final_df: pd.DataFrame, ref_dict: dict):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.cell(row=1, column=1, value="자재코드")
    ws.cell(row=1, column=2, value="수량")
    ws.cell(row=1, column=4, value="단가")
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

st.markdown("""<div style='text-align: center; padding: 1rem; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border-radius: 10px; margin-bottom: 2rem;'>
    <h2 style='color: white; margin: 0;'>📋 발주서 자동화 앱 (Groq 최적화)</h2>
</div>""", unsafe_allow_html=True)

selected_mart = st.selectbox("1️⃣ 마트를 선택하세요", list(MART_OPTIONS.keys()))
uploaded_files = st.file_uploader("2️⃣ 발주서 사진을 업로드하세요", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

if uploaded_files and st.button("🚀 사진 분석 시작", type="primary", use_container_width=True):
    ref_dict = load_reference(MART_OPTIONS[selected_mart])
    all_results, all_warnings = [], []
    
    with st.spinner("AI가 사진을 분석하고 있습니다... (약 5~10초 소요)"):
        for f in uploaded_files:
            results, warnings = analyze_image(f, mart_type=selected_mart)
            all_results.extend(results)
            all_warnings.extend(warnings)

    if not all_results:
        st.error("❌ 추출된 데이터가 없습니다. 이미지를 다시 확인해주세요.")
    else:
        df = pd.DataFrame(all_results, columns=["바코드", "수량", "제품명"])
        df["단가"] = 0; df["상태"] = ""
        df = apply_lookup(df, ref_dict)
        
        st.session_state["ocr_df"] = df
        st.session_state["warnings"] = all_warnings
        st.session_state["selected_mart"] = selected_mart
        st.rerun()

if "ocr_df" in st.session_state:
    st.divider()
    st.subheader("3️⃣ 검수 및 수정 (더블클릭)")
    if st.session_state["warnings"]:
        with st.expander("🔔 AI 바코드 인식 경고", expanded=False):
            for w in st.session_state["warnings"]: st.write(f"• {w}")

    ref_dict = load_reference(MART_OPTIONS[st.session_state["selected_mart"]])
    
    edited_df = st.data_editor(
        st.session_state["ocr_df"], 
        num_rows="dynamic", 
        use_container_width=True, 
        hide_index=True
    )
    
    if st.button("🔄 수정사항 반영 (바코드/제품명 재대조)", use_container_width=True):
        updated = apply_lookup(edited_df.copy(), ref_dict)
        st.session_state["ocr_df"] = updated
        st.rerun()
        
    matched = len(edited_df[edited_df["상태"].str.contains("✅")]) if "상태" in edited_df.columns else 0
    st.success(f"현재 엑셀로 변환 가능한 정상 항목: **{matched}건**")

    st.divider()
    st.subheader("4️⃣ 최종 확정 및 엑셀 다운로드")
    if st.button("📥 최종 확정 및 엑셀 다운로드", type="primary", disabled=(matched==0), use_container_width=True):
        excel_bytes, count = create_excel_bytes(edited_df, ref_dict)
        st.download_button(
            "💾 전산업로드용 엑셀 다운로드", 
            data=excel_bytes, 
            file_name="전산업로드_결과.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )