import streamlit as st
import tempfile
import os
import json

from automation_core_cloud import (
    extract_pdf_text,
    parse_pdf_data,
    load_config,
    load_json,
    save_json,
    save_config,
    run_automation,
    CRO_DB_PATH,
)

import asyncio
import sys

if sys.platform.startswith("win"):
    try:
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
    except Exception:
        pass

st.set_page_config(page_title="업무의뢰서 자동입력", layout="wide", page_icon="🤖")
st.title("📄 업무의뢰서 자동입력 및 폼 채우기")

# ── State 초기화 ──────────────────────────────────────────
if "pdf_parsed" not in st.session_state:
    st.session_state.pdf_parsed = False
if "data" not in st.session_state:
    st.session_state.data = {}

# ── 상수 ─────────────────────────────────────────────────
CHECKBOXES = {
    "EDC": "item_22323_52866_0_0",
    "IWRS(F)": "item_22323_52865_0_0",
    "IWRS(S)": "item_22323_52864_0_0",
    "ePRO": "item_22323_291222_0_0",
    "FTP": "item_22323_4383463_0_0",
    "PSDV": "item_22323_4383464_0_0",
    "IE(F)": "item_22323_2028988_0_0",
    "IE(S)": "item_22323_2028990_0_0",
    "Central LAB": "item_22323_1249040_0_0",
    "DM": "item_22323_63688_0_0",
    "STAT": "item_22323_3129351_0_0",
    "SDTM": "item_22323_212029_0_0",
    "SAFETY": "item_22323_63689_0_0",
    "TMF": "item_22323_689306_0_0",
    "CTMS": "item_22323_63690_0_0",
    "CONSENT": "item_22323_530313_0_0",
    "DDC": "item_22323_2881324_0_0",
    "RBQM": "item_22323_3396617_0_0",
    "LMS": "item_22323_2028991_0_0",
}

CHECKBOX_RATIOS = {
    "item_22323_52866_0_0": 1,
    "item_22323_52865_0_0": 0.5,
    "item_22323_52864_0_0": 0.25,
    "item_22323_291222_0_0": 0.5,
    "item_22323_4383463_0_0": 0.5,
    "item_22323_4383464_0_0": 0.25,
    "item_22323_2028988_0_0": 0.5,
    "item_22323_2028990_0_0": 0.25,
    "item_22323_63689_0_0": 0.5,
    "item_22323_689306_0_0": 0.5,
    "item_22323_63690_0_0": 0.5,
    "item_22323_530313_0_0": 0.5,
    "item_22323_2881324_0_0": 0.5,
    "item_22323_3396617_0_0": 0.5,
    "item_22323_2028991_0_0": 0.5,
}

# ── PDF 업로드 ────────────────────────────────────────────
st.subheader("1. 업무의뢰서 (PDF) 업로드")
uploaded_file = st.file_uploader("PDF 파일을 드래그하거나 선택하세요", type="pdf")

if uploaded_file is not None and not st.session_state.pdf_parsed:
    with st.spinner("PDF 분석 중..."):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name
        text = extract_pdf_text(tmp_path)
        data = parse_pdf_data(text)
        st.session_state.data = data
        st.session_state.pdf_parsed = True
        st.rerun()
elif uploaded_file is None:
    st.session_state.pdf_parsed = False

# ── 폼 UI ────────────────────────────────────────────────
if st.session_state.pdf_parsed:
    data = st.session_state.data
    config = load_config()
    cro_db = load_json(CRO_DB_PATH, {})

    with st.expander("🛠️ 로그인 및 설정 정보", expanded=False):
        col1, col2, col3 = st.columns(3)
        site_url = col1.selectbox(
            "사이트 URL",
            ["https://cubecdms.com/login", "https://drn.cubecdms.com/login"],
            index=1 if "drn" in config.get("site_url", "") else 0,
        )
        user_id = col2.text_input("아이디 (ID)", value=config.get("id", ""))
        user_pw = col3.text_input("비밀번호 (PW)", value=config.get("pw", ""), type="password")

    col_proj, col_cro = st.columns(2)
    with col_proj:
        st.subheader("📋 과제 정보")
        protocol_title = st.text_input("임상연구제목", value=data.get("protocol_title", ""))
        protocol_no = st.text_input("Protocol No.", value=data.get("protocol_no", ""))
        phase = st.text_input("Phase", value=data.get("phase", ""))
        duration_month = st.number_input(
            "운용 기간(개월)",
            value=int(data.get("duration_month", 0)) if data.get("duration_month") else 0,
            step=1,
        )
        subject_num = st.number_input(
            "시험대상자수",
            value=int(data.get("subject_num", 0)) if data.get("subject_num") else 0,
            step=1,
        )
        monthly_fee = st.number_input(
            "월 비용 (원)",
            value=int(data.get("monthly_fee", 0)) if data.get("monthly_fee") else 0,
            step=1000,
        )
        billing_cycle = st.number_input(
            "청구 주기 (개월)",
            value=int(data.get("billing_cycle", 0)) if data.get("billing_cycle") else 0,
            step=1,
        )

    with col_cro:
        st.subheader("🤝 CRO 및 담당자 정보")
        cro_name = st.text_input("CRO 명칭", value=data.get("cro_name", ""))
        cro_info = cro_db.get(cro_name, {})
        person = st.text_input("담당자 성함", value=cro_info.get("person", ""))
        phone = st.text_input("전화번호", value=cro_info.get("phone", ""))
        email = st.text_input("이메일", value=cro_info.get("email", ""))

    st.subheader("☑️ 체크박스 선택 (PDF 자동인식)")
    cols = st.columns(4)
    selected_checkbox_ids = []
    for i, (label, cb_id) in enumerate(CHECKBOXES.items()):
        with cols[i % 4]:
            default_val = cb_id in data.get("selected_checkbox_ids", [])
            if st.checkbox(label, value=default_val, key=cb_id):
                selected_checkbox_ids.append(cb_id)

    st.markdown("---")

    col_btn1, col_btn2 = st.columns([1, 4])
    with col_btn1:
        if st.button("🔄 초기화"):
            st.session_state.pdf_parsed = False
            st.session_state.data = {}
            st.rerun()

    with col_btn2:
        if st.button("🚀 자동화 시작"):
            save_config({"site_url": site_url, "id": user_id, "pw": user_pw})
            cro_db[cro_name] = {"person": person, "phone": phone, "email": email}
            save_json(CRO_DB_PATH, cro_db)

            if not (site_url and user_id and user_pw):
                st.error("로그인 정보를 모두 입력해주세요.")
                st.stop()
            if not duration_month > 0 or not monthly_fee > 0 or not billing_cycle > 0:
                st.error("숫자 항목(운용기간, 월비용, 청구주기)은 0보다 커야 합니다.")
                st.stop()

            total_amount = monthly_fee * billing_cycle
            total_ratio = sum(CHECKBOX_RATIOS.get(cb, 0) for cb in selected_checkbox_ids)
            checkbox_amounts = {}
            allocated_sum = 0
            for i, cb_id in enumerate(selected_checkbox_ids):
                ratio = CHECKBOX_RATIOS.get(cb_id, 0)
                if i == len(selected_checkbox_ids) - 1:
                    amount = total_amount - allocated_sum
                else:
                    amount = int(total_amount * ratio / total_ratio) if total_ratio > 0 else 0
                allocated_sum += amount
                checkbox_amounts[cb_id] = amount

            result_payload = {
                "login_data": {"login_url": site_url, "user_id": user_id, "user_pw": user_pw},
                "form_data": {
                    "protocol_title": protocol_title,
                    "protocol_no": protocol_no,
                    "phase": phase,
                    "duration_month": duration_month,
                    "subject_num": str(subject_num),
                    "monthly_fee": monthly_fee,
                    "billing_cycle": billing_cycle,
                    "total_amount": total_amount,
                    "checkbox_amounts": checkbox_amounts,
                },
                "selected_checkbox_ids": selected_checkbox_ids,
                "cro_info": {"name": cro_name, "person": person, "phone": phone, "email": email},
            }

            log_container = st.container()
            log_messages = []

            def status_update(msg):
                log_messages.append(msg)
                with log_container:
                    for m in log_messages:
                        st.write(m)

            with st.spinner("웹 자동화 진행 중... 잠시 기다려주세요."):
                res = run_automation(result_payload, status_callback=status_update)

            if res["success"]:
                st.success("🎉 " + res["message"])
                st.balloons()
            else:
                st.error("⛔ 자동화 실패: " + res["message"])
                st.subheader("📋 실행 로그")
                for m in log_messages:
                    st.write(m)
