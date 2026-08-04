# -*- coding: utf-8 -*-
"""
관리자 페이지: visitor_log.csv 의 방문 기록을 조회/검색/퇴실일시 수정하고,
엑셀로 다운로드하거나 미발송 기록을 이메일 보고(테스트 모드)로 확인한다.

- 방문객이 쓰는 app.py 와 달리, 비밀번호(ADMIN_PASSWORD)로 접근을 제한한다.
- 방문 기록 저장 방식(CSV 스키마, 경로)은 app.py 와 utils.py 를 그대로 따른다.
"""

import os
import sys
from datetime import datetime

import pandas as pd
import streamlit as st

# pages/ 폴더의 스크립트는 실행 방식에 따라 프로젝트 루트가 sys.path 에 없을 수 있으므로,
# utils 모듈을 확실히 찾을 수 있도록 루트 폴더를 직접 추가한다.
_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

from utils import (  # noqa: E402
    KST,
    build_excel_bytes,
    get_setting,
    get_unsent_records,
    load_report_recipients,
    load_visitor_log_df,
    logger,
    save_checkout_times,
    simulate_report_email,
)

st.set_page_config(page_title="관리자 페이지 - 방문 기록 조회", page_icon="🔐", layout="wide")

st.title("관리자 페이지 - 방문 기록 조회")

# ---------------------------------------------------------------------------
# 관리자 인증
# ---------------------------------------------------------------------------

ADMIN_PASSWORD = get_setting("ADMIN_PASSWORD", "")

if "admin_authenticated" not in st.session_state:
    st.session_state["admin_authenticated"] = False


def render_login() -> None:
    """관리자 비밀번호 입력 화면을 그린다. 비밀번호가 맞지 않으면 방문 기록을 절대 보여주지 않는다."""
    st.info("방문 기록을 조회하려면 관리자 비밀번호를 입력해주세요.")

    if not ADMIN_PASSWORD:
        st.error(
            "관리자 비밀번호(ADMIN_PASSWORD)가 설정되어 있지 않습니다. "
            "'.env' 파일 또는 서버 환경변수에 ADMIN_PASSWORD 값을 설정한 뒤 다시 접속해주세요. "
            "(설정 방법은 README.md 를 참고하세요)"
        )
        return

    with st.form("admin_login_form"):
        password_input = st.text_input("관리자 비밀번호", type="password")
        login_submitted = st.form_submit_button("로그인")

    if login_submitted:
        if password_input == ADMIN_PASSWORD:
            st.session_state["admin_authenticated"] = True
            st.rerun()
        else:
            st.error("비밀번호가 일치하지 않습니다. 다시 확인해주세요.")


if not st.session_state["admin_authenticated"]:
    render_login()
    st.stop()

header_col, logout_col = st.columns([5, 1])
with header_col:
    st.caption("로그인되었습니다. 아래에서 방문 기록을 검색/조회할 수 있습니다.")
with logout_col:
    if st.button("로그아웃", width="stretch"):
        st.session_state["admin_authenticated"] = False
        st.rerun()

st.divider()

# ---------------------------------------------------------------------------
# 방문 기록 로딩 (visitor_log.csv 가 없거나 비어 있거나 손상되어도 앱이 죽지 않게 처리)
# ---------------------------------------------------------------------------

try:
    log_df = load_visitor_log_df()
except RuntimeError as e:
    st.error(str(e))
    if st.button("다시 시도"):
        st.rerun()
    st.stop()

if log_df.empty:
    st.info("아직 등록된 방문 기록이 없습니다. 방문객 등록 페이지에서 등록이 이루어지면 여기에 표시됩니다.")

# ---------------------------------------------------------------------------
# 상단 요약 지표(KPI)
# ---------------------------------------------------------------------------

today_kst = datetime.now(KST).date()

total_count = len(log_df)
today_count = int((log_df["registered_at_dt"].dt.date == today_kst).sum())
this_month_count = int(
    (
        (log_df["registered_at_dt"].dt.year == today_kst.year)
        & (log_df["registered_at_dt"].dt.month == today_kst.month)
    ).sum()
)
vehicle_count = int((log_df["vehicle_number"] != "").sum())

kpi1, kpi2, kpi3, kpi4 = st.columns(4)
kpi1.metric("전체 방문 건수", f"{total_count}건")
kpi2.metric("오늘 방문 건수", f"{today_count}건")
kpi3.metric("이번 달 방문 건수", f"{this_month_count}건")
kpi4.metric("차량번호 등록 건수", f"{vehicle_count}건")

st.divider()

# ---------------------------------------------------------------------------
# 검색 및 필터
# ---------------------------------------------------------------------------

st.subheader("검색 및 필터")


def unique_sorted_values(series: pd.Series) -> list:
    """실제 데이터에 존재하는 값만 정렬해서 선택지로 반환한다. (빈 값 제외)"""
    return sorted({v for v in series if v})


purpose_options = unique_sorted_values(log_df["visit_purpose"])
department_options = unique_sorted_values(log_df["host_department"])
location_options = unique_sorted_values(log_df["visit_location"])

date_col1, date_col2 = st.columns(2)
with date_col1:
    date_start = st.date_input("방문일자 시작", value=None, key="filter_date_start")
with date_col2:
    date_end = st.date_input("방문일자 종료", value=None, key="filter_date_end")

text_col1, text_col2, text_col3, text_col4 = st.columns(4)
with text_col1:
    name_query = st.text_input("방문객 성명", key="filter_name")
with text_col2:
    company_query = st.text_input("회사명", key="filter_company")
with text_col3:
    vehicle_query = st.text_input("차량번호", key="filter_vehicle")
with text_col4:
    host_name_query = st.text_input("사내 담당자", key="filter_host_name")

select_col1, select_col2, select_col3 = st.columns(3)
with select_col1:
    purpose_selected = st.multiselect("방문 목적", options=purpose_options, key="filter_purpose")
with select_col2:
    department_selected = st.multiselect("담당 부서", options=department_options, key="filter_department")
with select_col3:
    location_selected = st.multiselect("방문 장소", options=location_options, key="filter_location")

if st.button("필터 초기화"):
    for key in (
        "filter_date_start",
        "filter_date_end",
        "filter_name",
        "filter_company",
        "filter_vehicle",
        "filter_host_name",
        "filter_purpose",
        "filter_department",
        "filter_location",
    ):
        st.session_state.pop(key, None)
    st.rerun()


def contains_ci(series: pd.Series, query: str) -> pd.Series:
    """대소문자와 앞뒤 공백 영향을 없애고 일부 단어만 입력해도 검색되도록 하는 포함 검색."""
    query = query.strip().lower()
    if not query:
        return pd.Series(True, index=series.index)
    return series.str.lower().str.contains(query, na=False, regex=False)


filtered_df = log_df.copy()
filtered_df = filtered_df[contains_ci(filtered_df["visitor_name"], name_query)]
filtered_df = filtered_df[contains_ci(filtered_df["visitor_company"], company_query)]
filtered_df = filtered_df[contains_ci(filtered_df["vehicle_number"], vehicle_query)]
filtered_df = filtered_df[contains_ci(filtered_df["host_name"], host_name_query)]

if purpose_selected:
    filtered_df = filtered_df[filtered_df["visit_purpose"].isin(purpose_selected)]
if department_selected:
    filtered_df = filtered_df[filtered_df["host_department"].isin(department_selected)]
if location_selected:
    filtered_df = filtered_df[filtered_df["visit_location"].isin(location_selected)]

if date_start:
    filtered_df = filtered_df[filtered_df["registered_at_dt"].dt.date >= date_start]
if date_end:
    filtered_df = filtered_df[filtered_df["registered_at_dt"].dt.date <= date_end]

st.divider()

# ---------------------------------------------------------------------------
# 방문 기록 표 (퇴실일시만 편집 가능)
# ---------------------------------------------------------------------------

st.subheader("방문 기록")

display_df = filtered_df.sort_values("registered_at_dt", ascending=False, na_position="last")

st.write(f"조회된 방문 기록: **{len(display_df)}건**")

COLUMN_LABELS = {
    "registered_at": "등록일시",
    "checkout_at": "퇴실일시",
    "report_sent_at": "보고 발송일시",
    "visit_id": "방문 ID",
    "visitor_name": "방문객 성명",
    "visitor_company": "회사명",
    "visitor_phone": "연락처",
    "vehicle_number": "차량번호",
    "visit_purpose": "방문 목적",
    "host_department": "담당 부서",
    "host_name": "사내 담당자",
    "visit_location": "방문 장소",
    "privacy_consent": "개인정보 동의 여부",
}
TABLE_COLUMNS = list(COLUMN_LABELS.keys())
CHECKOUT_LABEL = COLUMN_LABELS["checkout_at"]
VISIT_ID_LABEL = COLUMN_LABELS["visit_id"]

view_df = display_df.copy()
view_df["privacy_consent"] = view_df["privacy_consent"].apply(lambda v: "동의" if v == "Y" else "미동의")
# 퇴실일시는 편집 위젯(DatetimeColumn)이 실제 날짜/시간 값을 받도록 파싱된 열로 바꿔치기한다.
view_df["checkout_at"] = view_df["checkout_at_dt"]

table_view = view_df[TABLE_COLUMNS].rename(columns=COLUMN_LABELS)

edited_view = st.data_editor(
    table_view,
    width="stretch",
    hide_index=True,
    num_rows="fixed",  # 행 추가/삭제를 허용하지 않는다.
    disabled=[label for key, label in COLUMN_LABELS.items() if key != "checkout_at"],
    column_config={
        CHECKOUT_LABEL: st.column_config.DatetimeColumn(
            CHECKOUT_LABEL,
            format="YYYY-MM-DD HH:mm",
            step=60,  # 초 단위 입력을 막고 분 단위까지만 입력받는다.
            help="방문객이 나간 일시를 입력하세요. 비워두면 아직 퇴실하지 않은 것으로 간주합니다.",
        ),
    },
    key="visitor_table_editor",
)

save_clicked = st.button("퇴실일시 저장", type="primary")

if save_clicked:
    changes = {
        table_view.iloc[i][VISIT_ID_LABEL]: edited_view.iloc[i][CHECKOUT_LABEL]
        for i in range(len(table_view))
    }
    try:
        save_result = save_checkout_times(changes)
    except RuntimeError as e:
        st.error(str(e))
    else:
        if save_result["updated_count"] == 0 and not save_result["skipped"]:
            st.info("변경된 내용이 없습니다.")
        else:
            if save_result["updated_count"] > 0:
                st.success(f"퇴실일시 {save_result['updated_count']}건이 저장되었습니다.")
            if save_result["skipped"]:
                st.warning(
                    "다음 방문 기록은 값이 올바르지 않거나 원본에서 찾을 수 없어 저장되지 않았습니다: "
                    + ", ".join(save_result["skipped"])
                )
            if save_result["updated_count"] > 0:
                st.rerun()

st.divider()

# ---------------------------------------------------------------------------
# 엑셀 다운로드 (현재 검색·필터 결과 기준) / 이메일 발송 테스트 (미발송 기록 전체 기준)
# ---------------------------------------------------------------------------

download_col, email_col = st.columns(2)

with download_col:
    st.subheader("엑셀 다운로드")
    st.caption("현재 검색·필터 결과를 다운로드합니다.")

    if display_df.empty:
        st.info("현재 조회된 방문 기록이 없어 다운로드할 파일이 없습니다.")
        st.download_button(
            "엑셀 다운로드 (.xlsx)",
            data=b"",
            file_name="방문기록.xlsx",
            disabled=True,
        )
    else:
        # 화면 표(퇴실일시가 편집 반영 전 원본 값)와 동일한 기준으로 다운로드 내용을 만든다.
        download_source_df = view_df.copy()
        download_source_df["checkout_at"] = display_df["checkout_at"]
        excel_bytes = build_excel_bytes(download_source_df, TABLE_COLUMNS, COLUMN_LABELS)
        file_name = f"방문기록_{datetime.now(KST).strftime('%Y%m%d')}.xlsx"
        st.download_button(
            "엑셀 다운로드 (.xlsx)",
            data=excel_bytes,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with email_col:
    st.subheader("이메일 발송 테스트")
    st.caption("현재 필터와 무관하게, 아직 보고에 포함되지 않은 방문 기록 전체가 대상입니다.")

    try:
        recipients_df = load_report_recipients()
    except RuntimeError as e:
        st.error(str(e))
        recipients_df = None

    if recipients_df is not None:
        recipient_label_map = {
            row["recipient_id"]: f"{row['name']} | {row['department']}"
            for _, row in recipients_df.iterrows()
        }
        selected_recipient_ids = st.multiselect(
            "이메일 수신자",
            options=recipients_df["recipient_id"].tolist(),
            format_func=lambda rid: recipient_label_map.get(rid, rid),
            key="report_recipient_ids",
        )

        if st.button("이메일 발송(테스트)"):
            if not selected_recipient_ids:
                st.warning("이메일 수신자를 한 명 이상 선택해주세요.")
            else:
                try:
                    unsent_df = get_unsent_records()
                except RuntimeError as e:
                    st.error(str(e))
                    unsent_df = None

                if unsent_df is not None:
                    if unsent_df.empty:
                        st.info("새로 발송할 방문 차량 기록이 없습니다.")
                    else:
                        sorted_unsent = unsent_df.sort_values(
                            "registered_at_dt", ascending=True, na_position="last"
                        )

                        EMAIL_ATTACHMENT_LABELS = {
                            "registered_at": "등록일시",
                            "checkout_at": "퇴실일시",
                            "visitor_name": "성명",
                            "visitor_company": "회사명",
                            "visitor_phone": "연락처",
                            "vehicle_number": "차량번호",
                            "visit_purpose": "방문목적",
                            "host_name": "담당자",
                            "visit_location": "출입장소",
                        }
                        EMAIL_ATTACHMENT_COLUMNS = list(EMAIL_ATTACHMENT_LABELS.keys())
                        totals_row = {
                            EMAIL_ATTACHMENT_LABELS["registered_at"]: "합계",
                            EMAIL_ATTACHMENT_LABELS["vehicle_number"]: f"총 {len(sorted_unsent)}대",
                        }

                        try:
                            attachment_bytes = build_excel_bytes(
                                sorted_unsent,
                                EMAIL_ATTACHMENT_COLUMNS,
                                EMAIL_ATTACHMENT_LABELS,
                                sheet_name="방문차량기록",
                                totals_row=totals_row,
                            )
                        except Exception as e:
                            st.error("첨부 엑셀 생성 중 오류가 발생했습니다. 서버 로그를 확인해주세요.")
                            logger.exception("이메일 첨부 엑셀 생성 실패: %s", e)
                        else:
                            executed_at = datetime.now(KST)
                            selected_recipients = recipients_df[
                                recipients_df["recipient_id"].isin(selected_recipient_ids)
                            ].to_dict("records")

                            # 테스트 모드: 아래 simulate_report_email() 은 실제 SMTP 연결/발송을
                            # 절대 수행하지 않으며, report_sent_at 값도 변경하지 않는다.
                            results = simulate_report_email(
                                selected_recipients,
                                unsent_count=len(sorted_unsent),
                                vehicle_count=len(sorted_unsent),
                                executed_at=executed_at,
                            )

                            st.success(
                                f"[테스트 모드] {len(results)}명에게 발송될 내용을 아래에서 확인하세요. "
                                "실제 이메일은 발송되지 않았고, 발송 완료 처리(보고 발송일시 갱신)도 되지 않았습니다."
                            )
                            for info in results:
                                with st.expander(f"{info['recipient_name']} ({info['recipient_email']})"):
                                    st.write(f"**수신자 이름**: {info['recipient_name']}")
                                    st.write(f"**수신자 부서**: {info['recipient_department']}")
                                    st.write(f"**수신자 이메일**: {info['recipient_email']}")
                                    st.write(f"**메일 제목**: {info['subject']}")
                                    st.write(f"**첨부파일명**: {info['attachment_filename']}")
                                    st.write(f"**테스트 실행 시각**: {info['executed_at']}")
                                    st.write(f"**첨부 대상 방문 기록 건수**: {info['unsent_count']}건")
                                    st.write(f"**총 방문 차량 수**: {info['vehicle_count']}대")
