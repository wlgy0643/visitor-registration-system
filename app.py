# -*- coding: utf-8 -*-
"""
방문객 출입등록 웹 애플리케이션 (테스트 프로토타입)

- 종이 방문록을 대체하는 Streamlit 기반 웹 페이지
- 방문객이 정보를 입력하고 담당자를 검색/선택하면 visitor_log.csv 에 기록을 저장
- 담당자에게는 실제 메일을 보내지 않고, 화면과 서버 로그에 "테스트 발송 내역"만 표시
"""

import os
import re
import uuid
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
from PIL import Image
from streamlit_drawable_canvas import st_canvas

from utils import (
    BASE_DIR,
    EMPLOYEE_CSV_PATH,
    KST,
    SIGNATURE_DIR,
    VISITOR_LOG_COLUMNS,
    VISITOR_LOG_PATH,
    get_setting,
    logger,
)

# ---------------------------------------------------------------------------
# 기본 설정
# ---------------------------------------------------------------------------

REQUIRED_EMPLOYEE_COLUMNS = ["employee_id", "name", "department", "email", "active"]

MAX_SEARCH_RESULTS = 20

# 실제 메일 발송 기능을 위한 설정값은 환경변수 또는 Streamlit secrets 에서만 읽어온다.
# (코드에는 비밀번호/계정정보를 절대 직접 작성하지 않는다)

EMAIL_TEST_MODE = get_setting("EMAIL_TEST_MODE", "true").strip().lower() != "false"


# ---------------------------------------------------------------------------
# 직원 명단 로딩
# ---------------------------------------------------------------------------

def load_employees() -> pd.DataFrame:
    """employees_dummy.csv 를 읽어서 활성(active=='Y') 직원 명단만 반환한다."""
    if not os.path.exists(EMPLOYEE_CSV_PATH):
        st.error(
            f"직원 명단 파일({EMPLOYEE_CSV_PATH})을 찾을 수 없습니다. "
            "프로젝트 폴더에 파일이 있는지 확인해주세요."
        )
        st.stop()

    try:
        df = pd.read_csv(EMPLOYEE_CSV_PATH, dtype=str, encoding="utf-8-sig")
    except Exception as e:
        st.error(f"직원 명단 파일을 읽는 중 오류가 발생했습니다: {e}")
        st.stop()

    df.columns = [c.strip() for c in df.columns]
    missing_columns = [c for c in REQUIRED_EMPLOYEE_COLUMNS if c not in df.columns]
    if missing_columns:
        st.error(
            "직원 명단 파일에 다음 필수 열이 없습니다: "
            + ", ".join(missing_columns)
            + f"\n(필요한 열: {', '.join(REQUIRED_EMPLOYEE_COLUMNS)})"
        )
        st.stop()

    for col in REQUIRED_EMPLOYEE_COLUMNS:
        df[col] = df[col].fillna("").astype(str).str.strip()

    active_df = df[df["active"] == "Y"].copy()
    if active_df.empty:
        st.warning("현재 검색 가능한(active=Y) 담당자가 없습니다. 직원 명단을 확인해주세요.")

    return active_df


def search_employees(employees_df: pd.DataFrame, query: str) -> pd.DataFrame:
    """이름 일부 또는 부서명으로 활성 직원을 검색한다."""
    query = query.strip().lower()
    if not query:
        return employees_df.iloc[0:0]

    mask = (
        employees_df["name"].str.lower().str.contains(query, na=False)
        | employees_df["department"].str.lower().str.contains(query, na=False)
    )
    return employees_df[mask]


# ---------------------------------------------------------------------------
# 방문 기록 저장
# ---------------------------------------------------------------------------

def generate_visit_id() -> str:
    """타임스탬프 + 임의값으로 중복되지 않는 visit_id 를 생성한다."""
    now = datetime.now(KST)
    return f"V{now.strftime('%Y%m%d%H%M%S')}-{uuid.uuid4().hex[:6].upper()}"


def append_visitor_log(record: dict) -> None:
    """방문 기록 한 건을 visitor_log.csv 에 추가한다. 기존 파일에 새 열이 없으면 자동 보완한다."""
    row_df = pd.DataFrame([record], columns=VISITOR_LOG_COLUMNS)
    file_exists = os.path.exists(VISITOR_LOG_PATH) and os.path.getsize(VISITOR_LOG_PATH) > 0

    if file_exists:
        existing_df = pd.read_csv(VISITOR_LOG_PATH, dtype=str, encoding="utf-8-sig")
        schema_changed = False
        for column in VISITOR_LOG_COLUMNS:
            if column not in existing_df.columns:
                existing_df[column] = ""
                schema_changed = True

        if schema_changed or list(existing_df.columns) != VISITOR_LOG_COLUMNS:
            existing_df = existing_df.reindex(columns=VISITOR_LOG_COLUMNS)
            existing_df.to_csv(VISITOR_LOG_PATH, index=False, encoding="utf-8-sig")

    row_df.to_csv(
        VISITOR_LOG_PATH,
        mode="a",
        header=not file_exists,
        index=False,
        encoding="utf-8-sig",
    )


def has_signature(image_data) -> bool:
    """서명 캔버스에 실제 필기 흔적이 있는지 확인한다."""
    if image_data is None:
        return False

    image_array = np.asarray(image_data)
    if image_array.ndim != 3:
        return False

    rgb = image_array[:, :, :3]
    # 흰색 배경이 아닌 픽셀이 일정량 이상이면 서명한 것으로 판단한다.
    non_white_pixels = np.any(rgb < 245, axis=2)
    return int(non_white_pixels.sum()) >= 20


def save_signature(image_data, visit_id: str) -> str:
    """서명 이미지를 PNG로 저장하고 app.py 기준 상대경로를 반환한다."""
    os.makedirs(SIGNATURE_DIR, exist_ok=True)
    filename = f"{visit_id}_signature.png"
    absolute_path = os.path.join(SIGNATURE_DIR, filename)

    image_array = np.asarray(image_data).astype("uint8")
    Image.fromarray(image_array).save(absolute_path, format="PNG")

    return os.path.relpath(absolute_path, BASE_DIR).replace("\\", "/")


# ---------------------------------------------------------------------------
# 이메일 테스트 모드
# ---------------------------------------------------------------------------

def build_email_content(
    host_name: str,
    visitor_name: str,
    visitor_company: str,
    visitor_phone: str,
    vehicle_number: str,
    visit_purpose: str,
    visit_location: str,
    registered_at_str: str,
) -> tuple:
    """담당자에게 보낼 메일의 제목과 본문을 만든다."""
    subject = f"[방문객 도착 안내] {visitor_name} 님이 방문하였습니다."
    vehicle_display = vehicle_number if vehicle_number else "미등록"

    body = (
        f"{host_name} 님,\n\n"
        f"{visitor_name} 님이 방문하였습니다. 아래 방문 정보를 확인해주세요.\n\n"
        f"- 방문자 성명: {visitor_name}\n"
        f"- 소속 회사명: {visitor_company}\n"
        f"- 연락처: {visitor_phone}\n"
        f"- 차량번호: {vehicle_display}\n"
        f"- 방문 목적: {visit_purpose}\n"
        f"- 방문 장소: {visit_location}\n"
        f"- 등록 시각: {registered_at_str}\n\n"
        "본 메일은 방문객 출입등록 시스템에서 자동 발송되었습니다."
    )
    return subject, body


def send_test_email(host_name: str, host_email: str, subject: str, body: str) -> dict:
    """
    실제 메일을 발송하지 않는 테스트 모드 함수.
    화면 표시 및 서버 로그 기록용 정보를 반환한다.

    실제 발송으로 전환할 때는 이 함수 대신 SMTP_HOST/SMTP_USERNAME/SMTP_PASSWORD 등을
    환경변수(.env) 또는 Streamlit secrets 에서 읽어와 smtplib 로 발송하도록 구현한다.
    """
    scheduled_at = datetime.now(KST)
    scheduled_at_str = scheduled_at.strftime("%Y-%m-%d %H:%M:%S")

    result = {
        "recipient_name": host_name,
        "recipient_email": host_email,
        "subject": subject,
        "body": body,
        "scheduled_at": scheduled_at_str,
        "test_mode": EMAIL_TEST_MODE,
    }

    logger.info(
        "[TEST EMAIL] 수신자=%s <%s> 제목=%s 발송예정시각=%s",
        host_name,
        host_email,
        subject,
        scheduled_at_str,
    )
    logger.info("[TEST EMAIL] 본문:\n%s", body)

    return result


# ---------------------------------------------------------------------------
# 입력값 검증
# ---------------------------------------------------------------------------

PHONE_PATTERN = re.compile(r"^[0-9\-\s()+]+$")


def validate_inputs(
    visitor_name: str,
    visitor_company: str,
    visitor_phone: str,
    visit_purpose: str,
    visit_location: str,
    selected_employee_id,
    consent: bool,
    signature_entered: bool,
) -> list:
    """필수 입력값, 담당자 선택 여부, 개인정보 동의 여부를 검증하고 오류 목록을 반환한다."""
    errors = []

    if not visitor_name.strip():
        errors.append("방문자 성명을 입력해주세요.")
    if not visitor_company.strip():
        errors.append("소속 회사명을 입력해주세요.")

    phone = visitor_phone.strip()
    if not phone:
        errors.append("연락처를 입력해주세요.")
    elif not PHONE_PATTERN.match(phone):
        errors.append("연락처는 숫자, 하이픈(-), 공백, 괄호만 입력할 수 있습니다.")

    if not visit_purpose.strip():
        errors.append("방문 목적을 선택해주세요. 기타를 선택한 경우 내용을 입력해주세요.")
    if not visit_location.strip():
        errors.append("방문 장소를 선택해주세요. 기타를 선택한 경우 내용을 입력해주세요.")

    if selected_employee_id is None:
        errors.append(
            "사내 담당자를 검색 결과 목록에서 선택해주세요. (텍스트만 입력하고 후보를 "
            "선택하지 않으면 등록할 수 없습니다.)"
        )

    if not consent:
        errors.append("개인정보 수집·이용에 동의해야 등록할 수 있습니다.")

    if not signature_entered:
        errors.append("방문객 서명을 입력해주세요.")

    return errors


# ---------------------------------------------------------------------------
# 화면(UI)
# ---------------------------------------------------------------------------

st.set_page_config(page_title="방문객 출입등록", page_icon="🚧", layout="centered")

# 폼 초기화 요청이 있으면, 위젯을 새로 그리기 전에 세션 상태를 먼저 비운다.
if st.session_state.get("_do_reset"):
    for key in [
        "visitor_name",
        "visitor_company",
        "visitor_phone",
        "vehicle_number",
        "visit_purpose_option",
        "visit_purpose_other",
        "visit_location_option",
        "visit_location_other",
        "host_query",
        "consent",
    ]:
        st.session_state.pop(key, None)
    st.session_state["signature_canvas_version"] = st.session_state.get("signature_canvas_version", 0) + 1
    st.session_state["_do_reset"] = False

st.title("방문객 출입등록")

st.warning(
    "⚠️ 테스트 모드로 운영 중입니다. 이 화면에서 등록해도 **실제 이메일은 발송되지 않으며**, "
    "직원 명단(employees_dummy.csv)의 이메일도 example.com 테스트 주소입니다.",
    icon="⚠️",
)

st.markdown(
    "출입 관리대장을 대체하는 방문객 등록 페이지입니다. 아래 항목을 입력하고 "
    "사내 담당자를 검색하여 선택한 뒤 등록해주세요. **\\*** 표시는 필수 입력 항목입니다."
)

employees_df = load_employees()

st.divider()

# --- 방문객 기본 정보 -------------------------------------------------------
st.subheader("방문객 정보")

visitor_name = st.text_input("방문자 성명 *", key="visitor_name")
visitor_company = st.text_input("소속 회사명 *", key="visitor_company")
visitor_phone = st.text_input(
    "연락처 *", key="visitor_phone", placeholder="예: 010-1234-5678"
)
vehicle_number = st.text_input(
    "차량번호 (선택)", key="vehicle_number", placeholder="예: 12가 3456"
)
visit_purpose_option = st.selectbox(
    "방문 목적 *",
    options=["선택해주세요", "업무", "회의", "공사", "용역", "기타"],
    key="visit_purpose_option",
)

if visit_purpose_option == "기타":
    visit_purpose_other = st.text_input(
        "기타 방문 목적 *",
        key="visit_purpose_other",
        placeholder="방문 목적을 직접 입력해주세요",
    )
    visit_purpose = visit_purpose_other.strip()
else:
    visit_purpose = "" if visit_purpose_option == "선택해주세요" else visit_purpose_option

visit_location_option = st.selectbox(
    "방문 장소 *",
    options=["선택해주세요", "본관동", "발전동", "기타"],
    key="visit_location_option",
)

if visit_location_option == "기타":
    visit_location_other = st.text_input(
        "기타 방문 장소 *",
        key="visit_location_other",
        placeholder="방문 장소를 직접 입력해주세요",
    )
    visit_location = visit_location_other.strip()
else:
    visit_location = "" if visit_location_option == "선택해주세요" else visit_location_option

# --- 담당자 검색(자동완성) ---------------------------------------------------
st.subheader("사내 담당자 검색 *")
st.caption("이름 일부 또는 부서명을 입력하면 후보가 표시됩니다. 목록에서 반드시 한 명을 선택해야 합니다.")

host_query = st.text_input(
    "담당자 이름 또는 부서 검색",
    key="host_query",
    placeholder="예: 김 또는 경영지원부",
    label_visibility="collapsed",
)

selected_employee_id = None
selected_employee_label = None

if host_query.strip():
    matches = search_employees(employees_df, host_query)

    if matches.empty:
        st.info("검색 결과가 없습니다. 다른 이름 또는 부서명으로 검색해주세요.")
    else:
        truncated = len(matches) > MAX_SEARCH_RESULTS
        matches = matches.head(MAX_SEARCH_RESULTS)

        label_map = {
            row["employee_id"]: f"{row['name']} | {row['department']}"
            for _, row in matches.iterrows()
        }
        option_ids = list(label_map.keys())

        selected_employee_id = st.radio(
            "검색 결과에서 담당자를 선택하세요",
            options=option_ids,
            format_func=lambda eid: label_map[eid],
            index=None,
            key=f"host_radio_{host_query.strip().lower()}",
        )

        if truncated:
            st.caption(
                f"검색 결과가 많아 상위 {MAX_SEARCH_RESULTS}건만 표시합니다. "
                "검색어를 구체적으로 입력하면 더 쉽게 찾을 수 있습니다."
            )

        if selected_employee_id is not None:
            selected_employee_label = label_map[selected_employee_id]
            st.caption(f"✅ 선택된 담당자: {selected_employee_label}")

st.divider()

# --- 방문객 서명 --------------------------------------------------------------
st.subheader("방문객 서명 *")
st.caption("아래 서명란에 방문객 본인이 마우스 또는 손가락으로 직접 서명해주세요.")

canvas_result = st_canvas(
    fill_color="rgba(255, 255, 255, 0)",
    stroke_width=3,
    stroke_color="#000000",
    background_color="#FFFFFF",
    height=180,
    width=650,
    drawing_mode="freedraw",
    display_toolbar=True,
    update_streamlit=True,
    key=f"signature_canvas_{st.session_state.get('signature_canvas_version', 0)}",
)

signature_entered = has_signature(canvas_result.image_data)

st.divider()

# --- 개인정보 동의 -----------------------------------------------------------
st.subheader("개인정보 수집·이용 동의 *")
st.markdown(
    "- **수집 항목**: 방문자 성명, 소속 회사명, 연락처, 차량번호(선택), 방문 목적, 방문 장소\n"
    "- **수집 목적**: 방문객 출입 관리 및 사내 담당자 방문 안내\n"
    "- **보유 기간**: 방문일로부터 90일간 보관 후 파기 "
    "(테스트 환경에서는 로컬 visitor_log.csv 파일에 저장됩니다)\n"
    "- 동의를 거부할 권리가 있으며, 동의하지 않을 경우 출입등록 및 담당자 안내가 제한됩니다."
)
consent = st.checkbox("위 개인정보 수집·이용에 동의합니다. (필수)", key="consent")

st.divider()

# --- 등록 처리 ---------------------------------------------------------------
submitted = st.button("등록하기", type="primary", width="stretch")

if submitted:
    errors = validate_inputs(
        visitor_name,
        visitor_company,
        visitor_phone,
        visit_purpose,
        visit_location,
        selected_employee_id,
        consent,
        signature_entered,
    )

    if errors:
        st.error("등록할 수 없습니다. 다음 항목을 확인해주세요:\n\n" + "\n".join(f"- {e}" for e in errors))
    else:
        host_row = employees_df[employees_df["employee_id"] == selected_employee_id].iloc[0]
        host_name = host_row["name"]
        host_department = host_row["department"]
        host_email = host_row["email"]

        now_kst = datetime.now(KST)
        registered_at_str = now_kst.strftime("%Y-%m-%d %H:%M:%S")
        visit_id = generate_visit_id()
        signature_path = save_signature(canvas_result.image_data, visit_id)

        record = {
            "visit_id": visit_id,
            "registered_at": registered_at_str,
            "checkout_at": "",  # 방문객 등록 시점에는 알 수 없으며, 관리자 페이지에서 나중에 입력한다.
            "report_sent_at": "",  # 아직 이메일 보고에 포함되지 않은 상태를 의미한다.
            "visitor_name": visitor_name.strip(),
            "visitor_company": visitor_company.strip(),
            "visitor_phone": visitor_phone.strip(),
            "vehicle_number": vehicle_number.strip(),
            "visit_purpose": visit_purpose.strip(),
            "host_employee_id": selected_employee_id,
            "host_name": host_name,
            "host_department": host_department,
            "host_email": host_email,
            "visit_location": visit_location.strip(),
            "privacy_consent": "Y",
            "signature_path": signature_path,
        }

        append_visitor_log(record)

        subject, body = build_email_content(
            host_name=host_name,
            visitor_name=record["visitor_name"],
            visitor_company=record["visitor_company"],
            visitor_phone=record["visitor_phone"],
            vehicle_number=record["vehicle_number"],
            visit_purpose=record["visit_purpose"],
            visit_location=record["visit_location"],
            registered_at_str=registered_at_str,
        )
        email_result = send_test_email(host_name, host_email, subject, body)

        st.session_state["_last_result"] = {
            "visit_id": visit_id,
            "email": email_result,
        }
        st.session_state["_do_reset"] = True
        st.rerun()

# --- 등록 완료 결과 표시 ------------------------------------------------------
last_result = st.session_state.get("_last_result")
if last_result:
    st.success(
        f"등록이 완료되었습니다. (방문 ID: {last_result['visit_id']}) "
        "담당자에게 방문 안내가 전달되었습니다.\n\n"
        "⚠️ 단, 현재는 테스트 모드이므로 실제 이메일은 발송되지 않았습니다."
    )

    with st.expander("📧 테스트 이메일 발송 내역 보기", expanded=True):
        email = last_result["email"]
        st.write(f"**수신자 이름**: {email['recipient_name']}")
        st.write(f"**수신자 이메일**: {email['recipient_email']}")
        st.write(f"**메일 제목**: {email['subject']}")
        st.write(f"**발송 예정 시각**: {email['scheduled_at']}")
        st.write("**메일 본문**:")
        st.code(email["body"], language=None)

