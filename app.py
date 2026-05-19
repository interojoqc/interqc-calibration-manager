from __future__ import annotations

import os
from datetime import date
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from cloud_integrations import (
    cloud_status,
    extract_text_with_vision,
    parse_certificate_candidates,
    replace_sheet_with_dataframe,
    upload_file_to_drive,
)
from backend import (
    BACKEND_NAME,
    DEFAULT_EXCEL_PATH,
    add_calibration_record,
    calculate_corrected,
    calibration_history_df,
    contacts_df,
    dashboard_metrics,
    dataframe_to_xlsx_bytes,
    due_items,
    export_internal_certificate,
    get_import_log,
    get_instrument,
    guess_document_no,
    import_excel,
    init_db,
    instruments_df,
    make_kakao_message,
    mark_disposed,
    parse_cycle_months,
    save_uploaded_file,
    update_correction,
    update_instrument_master,
    upsert_contact,
    upsert_instrument,
)


st.set_page_config(page_title="계측기 검교정 관리", page_icon="QC", layout="wide")
init_db()


def is_qc() -> bool:
    return st.session_state.get("role") == "QC" and st.session_state.get("qc_ok", False)


def require_qc() -> bool:
    if not is_qc():
        st.warning("QC 권한에서만 입력/수정할 수 있습니다. 사용부서는 열람만 가능합니다.")
        return False
    return True


def date_str(value: date | None) -> str:
    return value.isoformat() if value else ""


def file_download_button(label: str, path_value: str | None, key: str) -> None:
    if not path_value:
        st.caption("등록 파일 없음")
        return
    if path_value.startswith("http://") or path_value.startswith("https://"):
        st.link_button(label, path_value)
        return
    path = Path(path_value)
    if not path.exists():
        st.caption("파일 경로 오류")
        return
    with open(path, "rb") as file:
        st.download_button(label, file, file_name=path.name, key=key)


def selected_instrument(df: pd.DataFrame, label: str = "계측기 선택") -> dict | None:
    if df.empty:
        st.info("등록된 계측기가 없습니다.")
        return None
    options = {f"{r.management_no} | {r.name} | {r.location}": int(r.id) for r in df.itertuples()}
    selected = st.selectbox(label, list(options.keys()))
    return get_instrument(options[selected])


def save_certificate_file(uploaded_file, management_no: str, category: str) -> str:
    if uploaded_file is None:
        return ""
    status = cloud_status()
    if status.google_libs and status.credentials and status.drive_folder_id:
        try:
            return upload_file_to_drive(uploaded_file.name, uploaded_file.getvalue(), uploaded_file.type or "application/octet-stream")
        except Exception as exc:
            st.warning(f"Google Drive 업로드 실패로 로컬에 저장합니다: {exc}")
    return save_uploaded_file(uploaded_file, category, management_no)


st.sidebar.title("계측기 관리")
role = st.sidebar.radio("접속 역할", ["사용부서", "QC"], key="role")
if role == "QC":
    password = st.sidebar.text_input("QC 암호", type="password")
    expected = os.getenv("QC_APP_PASSWORD", "QC2026")
    st.session_state["qc_ok"] = password == expected
    if st.session_state["qc_ok"]:
        st.sidebar.success("QC 수정 권한 활성화")
    elif password:
        st.sidebar.error("암호가 맞지 않습니다.")
else:
    st.session_state["qc_ok"] = False

qc_pages = [
    "대시보드",
    "계측기 대장",
    "신규 등록",
    "계측기 수정",
    "검교정/보정 입력",
    "알림 문구",
    "폐기 계측기 관리",
    "담당자 관리",
    "클라우드 설정",
    "데이터 가져오기",
]
user_pages = ["대시보드", "계측기 대장", "폐기 계측기 관리"]
available_pages = qc_pages if is_qc() else user_pages
if role == "QC" and not is_qc():
    st.sidebar.info("QC 암호를 입력하면 QC 전용 메뉴가 표시됩니다.")
page = st.sidebar.radio(
    "메뉴",
    available_pages,
)

st.title("계측기 검교정 관리 시스템")
st.caption(f"현재 저장소: {BACKEND_NAME}")

if page == "대시보드":
    metrics = dashboard_metrics()
    all_df = instruments_df(include_disposed=True)
    active_df = all_df[all_df["status"] != "폐기"] if not all_df.empty else all_df
    due_30 = due_items(30)
    due_90 = due_items(90)
    overdue = due_90[due_90["남은일수"] < 0] if not due_90.empty else due_90
    disposed_df = all_df[all_df["status"] == "폐기"] if not all_df.empty else all_df
    disposal_missing = disposed_df[disposed_df["disposal_report_file_path"].fillna("") == ""] if not disposed_df.empty else disposed_df
    cert_missing = active_df[
        (active_df["last_record_id"].fillna("") != "")
        & (active_df["last_certificate_file_path"].fillna("") == "")
    ] if not active_df.empty else active_df

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("전체 계측기", f"{metrics['total']:,}")
    c2.metric("사용 계측기", f"{metrics['active']:,}")
    c3.metric("폐기 계측기", f"{metrics['disposed']:,}")
    c4.metric("기한 초과", f"{metrics['overdue']:,}")
    c5.metric("90일 내 도래", f"{metrics['due_90']:,}")

    s1, s2, s3, s4 = st.columns(4)
    s1.metric("30일 내 도래", f"{len(due_30):,}")
    s2.metric("폐기보고서 미등록", f"{len(disposal_missing):,}")
    s3.metric("성적서 파일 미등록", f"{len(cert_missing):,}")
    s4.metric("검교정 미등록", f"{int(active_df['last_record_id'].isna().sum()) if not active_df.empty else 0:,}")

    st.subheader("우선 조치 리스트")
    priority_tabs = st.tabs(["기한 초과", "30일 내 도래", "폐기보고서 미등록"])
    with priority_tabs[0]:
        if overdue.empty:
            st.success("기한 초과 계측기가 없습니다.")
        else:
            st.dataframe(
                overdue[["management_no", "name", "department", "location", "process", "next_due_date", "남은일수", "department_owner"]].rename(
                    columns={
                        "management_no": "관리번호",
                        "name": "계측기명",
                        "department": "담당부서",
                        "location": "위치",
                        "process": "공정",
                        "next_due_date": "차기 교정일",
                        "department_owner": "담당자",
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )
    with priority_tabs[1]:
        if due_30.empty:
            st.success("30일 내 도래 계측기가 없습니다.")
        else:
            st.dataframe(
                due_30[["management_no", "name", "department", "location", "process", "next_due_date", "남은일수", "department_owner"]].rename(
                    columns={
                        "management_no": "관리번호",
                        "name": "계측기명",
                        "department": "담당부서",
                        "location": "위치",
                        "process": "공정",
                        "next_due_date": "차기 교정일",
                        "department_owner": "담당자",
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )
    with priority_tabs[2]:
        if disposal_missing.empty:
            st.success("폐기보고서 미등록 계측기가 없습니다.")
        else:
            st.dataframe(
                disposal_missing[["management_no", "name", "department", "location", "process", "remark"]].rename(
                    columns={
                        "management_no": "관리번호",
                        "name": "계측기명",
                        "department": "담당부서",
                        "location": "위치",
                        "process": "공정",
                        "remark": "비고",
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

    due = due_90
    st.subheader("향후 3개월 검교정 도래")
    if due.empty:
        st.info("도래 대상이 없습니다. 먼저 데이터를 가져오거나 검교정 이력을 입력하세요.")
    else:
        view = due[
            [
                "management_no",
                "name",
                "serial_no",
                "department",
                "location",
                "last_calibration_type",
                "last_calibration_date",
                "next_due_date",
                "남은일수",
                "department_owner",
            ]
        ].rename(
            columns={
                "management_no": "관리번호",
                "name": "계측기명",
                "serial_no": "제작 일련번호",
                "department": "사용부서",
                "location": "설치 위치",
                "last_calibration_type": "검교정 구분",
                "last_calibration_date": "최근 교정일",
                "next_due_date": "차기 교정일",
                "department_owner": "부서 담당자",
            }
        )
        st.dataframe(view, use_container_width=True, hide_index=True)
        by_dept = due.groupby("department", dropna=False).size().reset_index(name="건수")
        fig = px.bar(by_dept, x="department", y="건수", title="사용부서별 90일 내 도래 건수")
        st.plotly_chart(fig, use_container_width=True)

elif page == "계측기 대장":
    df = instruments_df(include_disposed=True)
    st.caption("보정값(더하기)은 측정값에 더하는 값이고, 보정계수(곱하기)는 더한 뒤 곱하는 계수입니다. 보정 적용값 = (측정값 + 보정값) x 보정계수")
    c1, c2, c3 = st.columns(3)
    dept_options = sorted([x for x in df.get("department", pd.Series(dtype=str)).dropna().unique()])
    process_options = sorted([x for x in df.get("process", pd.Series(dtype=str)).dropna().unique()])
    location_options = sorted([x for x in df.get("location", pd.Series(dtype=str)).dropna().unique()])
    if is_qc():
        status_filter = c1.multiselect("상태", ["사용", "폐기"], default=["사용"], help="기본값은 사용 계측기만 표시합니다. 폐기까지 보려면 폐기를 추가하세요.")
        cal_filter = c2.multiselect("최근 검교정 구분", ["내부", "외부", "미등록"], default=[])
        due_filter = c3.selectbox("차기 교정일", ["전체", "기한 초과", "30일 내", "90일 내", "미등록"])
        c4, c5, c6 = st.columns(3)
        dept_filter = c4.multiselect("담당부서", dept_options)
        process_filter = c5.multiselect("공정", process_options)
        location_filter = c6.multiselect("위치", location_options)
    else:
        status_filter = ["사용"]
        cal_filter = []
        due_filter = c1.selectbox("차기 교정일", ["전체", "기한 초과", "30일 내", "90일 내"])
        dept_filter = c2.multiselect("담당부서", dept_options)
        process_filter = c3.multiselect("공정", process_options)
        location_filter = []
    search = st.text_input("검색", placeholder="관리번호, 계측기명, 일련번호, 담당자")
    filtered = df[df["status"].isin(status_filter)] if not df.empty else df
    if dept_filter:
        filtered = filtered[filtered["department"].isin(dept_filter)]
    if process_filter:
        filtered = filtered[filtered["process"].isin(process_filter)]
    if location_filter:
        filtered = filtered[filtered["location"].isin(location_filter)]
    if cal_filter:
        cal_mask = pd.Series(False, index=filtered.index)
        if "미등록" in cal_filter:
            cal_mask |= filtered["last_calibration_type"].isna() | (filtered["last_calibration_type"].fillna("") == "")
        selected_cal = [x for x in cal_filter if x != "미등록"]
        if selected_cal:
            cal_mask |= filtered["last_calibration_type"].isin(selected_cal)
        filtered = filtered[cal_mask]
    if due_filter != "전체":
        due_dates = pd.to_datetime(filtered["next_due_date"], errors="coerce")
        today = pd.Timestamp(date.today())
        if due_filter == "기한 초과":
            filtered = filtered[due_dates < today]
        elif due_filter == "30일 내":
            filtered = filtered[(due_dates >= today) & (due_dates <= today + pd.Timedelta(days=30))]
        elif due_filter == "90일 내":
            filtered = filtered[(due_dates >= today) & (due_dates <= today + pd.Timedelta(days=90))]
        elif due_filter == "미등록":
            filtered = filtered[due_dates.isna()]
    if search:
        terms = filtered[["management_no", "name", "serial_no", "department_owner", "department_owner2"]].fillna("").agg(" ".join, axis=1)
        filtered = filtered[terms.str.contains(search, case=False, na=False)]
    st.caption(f"조회 결과: {len(filtered):,}건")

    if is_qc():
        cols = [
            "management_no", "name", "location", "process", "department", "department_owner",
            "department_owner2", "serial_no", "cycle_text", "status", "last_calibration_type",
            "last_calibration_date", "next_due_date", "last_certificate_no", "correction_offset",
            "correction_factor", "correction_unit", "last_measured_value", "last_corrected_value", "remark",
        ]
    else:
        cols = [
            "management_no", "name", "location", "process", "department", "department_owner",
            "department_owner2", "cycle_text", "last_calibration_type", "last_calibration_date",
            "next_due_date", "remark",
        ]
    show = filtered[[c for c in cols if c in filtered.columns]].rename(
        columns={
            "management_no": "관리번호",
            "name": "계측기명",
            "location": "위치",
            "process": "공정",
            "department": "담당부서",
            "department_owner": "담당자",
            "department_owner2": "담당자 2",
            "serial_no": "제작 일련번호",
            "cycle_text": "교정주기",
            "status": "상태",
            "last_calibration_type": "최근 구분",
            "last_calibration_date": "최근 교정일",
            "next_due_date": "차기 교정일",
            "last_certificate_no": "성적서 번호",
            "correction_offset": "보정값(더하기)",
            "correction_factor": "보정계수(곱하기)",
            "correction_unit": "단위",
            "last_measured_value": "최근 측정값",
            "last_corrected_value": "최근 보정 적용값",
            "remark": "비고",
        }
    )
    st.dataframe(show, use_container_width=True, hide_index=True)
    st.download_button(
        "조회 결과 엑셀 다운로드",
        dataframe_to_xlsx_bytes(show),
        "계측기_조회결과.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.subheader("계측기별 성적서 다운로드")
    st.caption("성적서 다운로드 버튼은 `검교정/보정 입력`에서 성적서 스캔본/PDF를 업로드했거나, 내부교정 성적서를 자동 생성한 이력이 있을 때 표시됩니다.")
    instrument = selected_instrument(filtered, "성적서를 확인할 계측기")
    if instrument:
        history = calibration_history_df(instrument["id"])
        if history.empty:
            st.info("이 계측기의 검교정 이력이 아직 없습니다.")
        else:
            history_view = history[
                [
                    "calibration_type",
                    "calibration_date",
                    "next_due_date",
                    "certificate_no",
                    "certificate_file_path",
                    "measured_value",
                    "corrected_value",
                    "note",
                ]
            ].copy()
            history_view["certificate_file_path"] = history_view["certificate_file_path"].apply(lambda x: "다운로드 가능" if x else "파일 미등록")
            st.dataframe(
                history_view.rename(
                columns={
                    "calibration_type": "구분",
                    "calibration_date": "교정일",
                    "next_due_date": "차기교정일",
                    "certificate_no": "성적서 번호",
                    "certificate_file_path": "성적서 파일",
                    "measured_value": "측정값",
                    "corrected_value": "보정 적용값",
                    "note": "비고",
                }
                ),
                use_container_width=True,
                hide_index=True,
            )
            has_file = False
            for row in history.itertuples():
                if getattr(row, "certificate_file_path", ""):
                    has_file = True
                    file_download_button(
                        f"{row.calibration_date or '날짜없음'} {row.calibration_type} 성적서 다운로드",
                        row.certificate_file_path,
                        f"cert_{row.id}",
                    )
            if not has_file:
                st.warning("이 계측기에는 아직 등록된 성적서 파일이 없습니다. `검교정/보정 입력` 메뉴에서 성적서 스캔본/PDF를 업로드하면 여기에 다운로드 버튼이 표시됩니다.")

elif page == "신규 등록":
    st.subheader("신규 계측기 등록")
    disabled = not is_qc()
    with st.form("new_instrument"):
        c1, c2, c3 = st.columns(3)
        management_no = c1.text_input("관리번호", disabled=disabled)
        name = c2.text_input("계측기명", disabled=disabled)
        serial_no = c3.text_input("제작 일련번호", disabled=disabled)
        c4, c5, c6 = st.columns(3)
        cycle_text = c4.selectbox("교정주기", ["12개월", "내부 6개월", "24개월", "내부 12개월", "6개월", "기타"], disabled=disabled)
        custom_cycle = c5.text_input("기타 주기", disabled=disabled)
        location = c6.text_input("위치", disabled=disabled)
        c7, c8, c9 = st.columns(3)
        process = c7.text_input("공정", disabled=disabled)
        department = c8.text_input("담당부서", disabled=disabled)
        department_owner = c9.text_input("담당자", disabled=disabled)
        c10, c11 = st.columns(2)
        department_owner2 = c10.text_input("담당자 2", disabled=disabled)
        qc_owner = c11.text_input("QC 담당자", disabled=disabled)
        remark = st.text_area("비고", disabled=disabled)
        is_standard = st.checkbox("표준품", disabled=disabled)
        submitted = st.form_submit_button("등록", disabled=disabled)
    if submitted and require_qc():
        if not management_no or not name:
            st.error("관리번호와 계측기명은 필수입니다.")
        else:
            final_cycle = custom_cycle if cycle_text == "기타" else cycle_text
            upsert_instrument(
                {
                    "management_no": management_no,
                    "name": name,
                    "serial_no": serial_no,
                    "cycle_text": final_cycle,
                    "cycle_months": parse_cycle_months(final_cycle),
                    "location": location,
                    "process": process,
                    "department": department,
                    "department_owner": department_owner,
                    "department_owner2": department_owner2,
                    "qc_owner": qc_owner,
                    "is_standard": is_standard,
                    "status": "사용",
                    "remark": remark,
                }
            )
            st.success("신규 계측기를 등록했습니다.")

elif page == "계측기 수정":
    st.subheader("계측기별 기본정보 수정")
    df = instruments_df(include_disposed=True)
    instrument = selected_instrument(df, "수정할 계측기 선택")
    disabled = not is_qc()
    if instrument:
        with st.form("edit_instrument"):
            c1, c2, c3 = st.columns(3)
            c1.text_input("관리번호", value=instrument.get("management_no", ""), disabled=True)
            name = c2.text_input("계측기명", value=instrument.get("name", ""), disabled=disabled)
            serial_no = c3.text_input("제작 일련번호", value=instrument.get("serial_no", ""), disabled=disabled)
            c4, c5, c6 = st.columns(3)
            cycle_text = c4.text_input("교정주기", value=instrument.get("cycle_text", ""), disabled=disabled)
            location = c5.text_input("위치", value=instrument.get("location", ""), disabled=disabled)
            process = c6.text_input("공정", value=instrument.get("process", ""), disabled=disabled)
            c7, c8, c9 = st.columns(3)
            department = c7.text_input("담당부서", value=instrument.get("department", ""), disabled=disabled)
            department_owner = c8.text_input("담당자", value=instrument.get("department_owner", ""), disabled=disabled)
            department_owner2 = c9.text_input("담당자 2", value=instrument.get("department_owner2", ""), disabled=disabled)
            c10, c11, c12 = st.columns(3)
            qc_owner = c10.text_input("QC 담당자", value=instrument.get("qc_owner", ""), disabled=disabled)
            status = c11.selectbox("상태", ["사용", "폐기"], index=1 if instrument.get("status") == "폐기" else 0, disabled=disabled)
            is_standard = c12.checkbox("표준품", value=bool(int(instrument.get("is_standard") or 0)), disabled=disabled)
            remark = st.text_area("비고", value=instrument.get("remark", ""), disabled=disabled)
            submitted = st.form_submit_button("수정 저장", disabled=disabled)
        if submitted and require_qc():
            update_instrument_master(
                int(instrument["id"]),
                {
                    "name": name,
                    "serial_no": serial_no,
                    "cycle_text": cycle_text,
                    "cycle_months": parse_cycle_months(cycle_text),
                    "location": location,
                    "process": process,
                    "department": department,
                    "department_owner": department_owner,
                    "department_owner2": department_owner2,
                    "qc_owner": qc_owner,
                    "status": status,
                    "is_standard": is_standard,
                    "remark": remark,
                },
            )
            st.success("계측기 기본정보를 수정했습니다.")

elif page == "검교정/보정 입력":
    st.subheader("검교정 이력, 성적서, 보정값 입력")
    df = instruments_df(include_disposed=False)
    instrument = selected_instrument(df)
    if instrument:
        status = cloud_status()
        if status.vision_ready:
            st.success("Google Vision OCR 설정이 감지되었습니다. 이미지 성적서에서 번호/날짜 후보를 추출할 수 있습니다.")
        else:
            st.info("Google Vision OCR 설정 전에는 파일명에서 번호 후보만 가져옵니다. `클라우드 설정` 메뉴에서 필요한 설정을 확인하세요.")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("현재 보정값(더하기)", instrument.get("correction_offset") or 0)
        c2.metric("현재 보정계수(곱하기)", instrument.get("correction_factor") or 1)
        c3.metric("단위", instrument.get("correction_unit") or "-")
        c4.metric("최근 보정 적용값", instrument.get("last_corrected_value") if instrument.get("last_corrected_value") is not None else "-")

        disabled = not is_qc()
        st.markdown("**보정 기준**")
        with st.form("correction_rule"):
            r1, r2, r3 = st.columns(3)
            offset = r1.number_input("보정값(더하기)", value=float(instrument.get("correction_offset") or 0), format="%.6f", disabled=disabled)
            factor = r2.number_input("보정계수(곱하기)", value=float(instrument.get("correction_factor") or 1), format="%.6f", disabled=disabled)
            unit = r3.text_input("단위", value=instrument.get("correction_unit") or "", disabled=disabled)
            correction_note = st.text_area("보정 기준/메모", value=instrument.get("correction_note") or "", disabled=disabled)
            save_rule = st.form_submit_button("보정 기준 저장", disabled=disabled)
        if save_rule and require_qc():
            update_correction(instrument["id"], offset, factor, unit, correction_note)
            st.success("보정 기준을 저장했습니다.")
            instrument = get_instrument(instrument["id"])

        st.markdown("**검교정 완료 입력**")
        uploaded = st.file_uploader("성적서 스캔본/PDF 업로드", type=["pdf", "png", "jpg", "jpeg", "xlsx"], disabled=disabled)
        guessed_no = guess_document_no(uploaded)
        ocr_candidates = {}
        if uploaded and status.vision_ready and uploaded.type and uploaded.type.startswith("image/"):
            if st.button("OCR로 성적서 후보 읽기", disabled=disabled):
                try:
                    text = extract_text_with_vision(uploaded.name, uploaded.getvalue())
                    ocr_candidates = parse_certificate_candidates(text)
                    st.session_state["ocr_candidates"] = ocr_candidates
                    with st.expander("OCR 원문 확인"):
                        st.text_area("추출 텍스트", text, height=220)
                except Exception as exc:
                    st.error(f"OCR 처리 실패: {exc}")
        ocr_candidates = st.session_state.get("ocr_candidates", {}) if uploaded else {}
        if ocr_candidates:
            st.write("OCR 후보:", ocr_candidates)
            guessed_no = ocr_candidates.get("certificate_no", guessed_no)
        with st.form("calibration_input"):
            c1, c2, c3 = st.columns(3)
            cal_type = c1.selectbox("검교정 구분", ["내부", "외부"], disabled=disabled)
            cal_date = c2.date_input("교정일자", value=date.today(), disabled=disabled)
            due_date = c3.date_input("차기 교정일", value=date.today(), disabled=disabled)
            c4, c5 = st.columns(2)
            result = c4.selectbox("판정", ["적합", "부적합", "수리", "기존 대장", "기타"], disabled=disabled)
            certificate_no = c5.text_input("성적서 번호", value=guessed_no, disabled=disabled)
            measured = st.number_input("측정값", format="%.6f", disabled=disabled)
            corrected = calculate_corrected(
                measured,
                float(instrument.get("correction_offset") or 0),
                float(instrument.get("correction_factor") or 1),
            )
            st.metric("보정 적용값", f"{corrected:,.6f} {instrument.get('correction_unit') or ''}")
            note = st.text_area("입력 내용/비고", disabled=disabled)
            make_internal_cert = st.checkbox("내부교정 성적서 파일 자동 생성", value=False, disabled=disabled)
            submitted = st.form_submit_button("검교정 이력 저장", disabled=disabled)
        if submitted and require_qc():
            cert_path = save_certificate_file(uploaded, instrument["management_no"], "certificates") if uploaded else ""
            record = {
                "instrument_id": instrument["id"],
                "calibration_type": cal_type,
                "calibration_date": date_str(cal_date),
                "next_due_date": date_str(due_date),
                "result": result,
                "certificate_no": certificate_no,
                "certificate_file_path": cert_path,
                "measured_value": measured,
                "corrected_value": corrected,
                "correction_snapshot": f"offset={instrument.get('correction_offset')}, factor={instrument.get('correction_factor')}",
                "note": note,
            }
            if cal_type == "내부" and make_internal_cert:
                generated = export_internal_certificate(instrument, record)
                record["certificate_file_path"] = str(generated)
                if not record["certificate_no"]:
                    record["certificate_no"] = generated.stem
            add_calibration_record(record)
            st.success("검교정 이력을 저장했습니다.")

        st.subheader("선택 계측기 이력")
        history = calibration_history_df(instrument["id"])
        st.dataframe(history, use_container_width=True, hide_index=True)
        has_file = False
        for row in history.itertuples():
            if getattr(row, "certificate_file_path", ""):
                has_file = True
                file_download_button(
                    f"{row.calibration_date or '날짜없음'} {row.calibration_type} 성적서 다운로드",
                    row.certificate_file_path,
                    f"hist_cert_{row.id}",
                )
        if not history.empty and not has_file:
            st.caption("등록된 성적서 파일이 있는 이력이 없어서 다운로드 버튼은 표시되지 않습니다.")

elif page == "알림 문구":
    st.subheader("카카오톡 알림 문구")
    c1, c2 = st.columns(2)
    days = c1.slider("도래 기준일", 30, 180, 90, 15)
    cal_filter = c2.selectbox("검교정 구분", ["전체", "내부", "외부"])
    due = due_items(days, calibration_filter=cal_filter)
    contacts = contacts_df()
    if due.empty:
        st.info("알림 대상이 없습니다.")
    else:
        departments = ["전체"] + sorted(due["department"].fillna("미지정").unique())
        department = st.selectbox("발송 대상", departments)
        rows = due if department == "전체" else due[due["department"].fillna("미지정") == department]
        owner = ""
        if department != "전체" and not contacts.empty:
            contact = contacts[contacts["department"] == department]
            if not contact.empty:
                owner = str(contact.iloc[0].get("owner_name") or "")
        message = make_kakao_message(department, rows, owner)
        st.dataframe(
            rows[["department", "management_no", "name", "location", "process", "department_owner", "department_owner2", "last_calibration_type", "next_due_date", "남은일수"]].rename(
                columns={
                    "department": "사용부서",
                    "management_no": "관리번호",
                    "name": "계측기명",
                    "location": "위치",
                    "process": "공정",
                    "department_owner": "담당자",
                    "department_owner2": "담당자 2",
                    "last_calibration_type": "구분",
                    "next_due_date": "차기 교정일",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )
        st.caption(f"문구 포함 건수: {len(rows):,}건")
        st.text_area("복사해서 보낼 문구", value=message, height=420)

elif page == "폐기 계측기 관리":
    st.subheader("폐기 계측기 관리")
    df = instruments_df(include_disposed=True)
    disposed = df[df["status"] == "폐기"] if not df.empty else df
    if not disposed.empty:
        c1, c2, c3 = st.columns(3)
        report_filter = c1.selectbox("폐기보고서 등록 여부", ["전체", "등록", "미등록"])
        dept_options = sorted([x for x in disposed.get("department", pd.Series(dtype=str)).dropna().unique()])
        dept_filter = c2.multiselect("담당부서", dept_options)
        search = c3.text_input("폐기 계측기 검색", placeholder="관리번호, 계측기명, 일련번호")
        if report_filter == "등록":
            disposed = disposed[disposed["disposal_report_file_path"].fillna("") != ""]
        elif report_filter == "미등록":
            disposed = disposed[disposed["disposal_report_file_path"].fillna("") == ""]
        if dept_filter:
            disposed = disposed[disposed["department"].isin(dept_filter)]
        if search:
            terms = disposed[["management_no", "name", "serial_no"]].fillna("").agg(" ".join, axis=1)
            disposed = disposed[terms.str.contains(search, case=False, na=False)]
        missing_reports = disposed[disposed["disposal_report_file_path"].fillna("") == ""]
        st.caption(f"조회 결과: {len(disposed):,}건 / 폐기보고서 미등록: {len(missing_reports):,}건")
    if disposed.empty:
        st.info("조건에 맞는 폐기 계측기가 없습니다.")
    else:
        show = disposed[
            [
                "management_no",
                "name",
                "serial_no",
                "location",
                "process",
                "department",
                "department_owner",
                "department_owner2",
                "disposal_report_no",
                "disposal_report_file_path",
                "remark",
            ]
        ].rename(
            columns={
                "management_no": "관리번호",
                "name": "계측기명",
                "serial_no": "일련번호",
                "location": "위치",
                "process": "공정",
                "department": "담당부서",
                "department_owner": "담당자",
                "department_owner2": "담당자 2",
                "disposal_report_no": "폐기 보고서 번호",
                "disposal_report_file_path": "폐기 보고서 등록여부",
                "remark": "비고",
            }
        )
        show["폐기 보고서 등록여부"] = show["폐기 보고서 등록여부"].apply(lambda x: "등록" if x else "미등록")
        st.dataframe(show, use_container_width=True, hide_index=True)
        if not missing_reports.empty:
            with st.expander("폐기보고서 미등록 리스트"):
                st.dataframe(
                    missing_reports[["management_no", "name", "location", "process", "department", "remark"]].rename(
                        columns={
                            "management_no": "관리번호",
                            "name": "계측기명",
                            "location": "위치",
                            "process": "공정",
                            "department": "담당부서",
                            "remark": "비고",
                        }
                    ),
                    use_container_width=True,
                    hide_index=True,
                )
        for row in disposed.itertuples():
            if getattr(row, "disposal_report_file_path", ""):
                file_download_button(f"{row.management_no} 폐기보고서 다운로드", row.disposal_report_file_path, f"disposal_{row.id}")

    if require_qc() and not df.empty:
        active = df[df["status"] != "폐기"]
        options = {f"{r.management_no} | {r.name}": int(r.id) for r in active.itertuples()}
        if options:
            selected = st.selectbox("폐기 처리할 계측기", list(options.keys()))
            instrument = get_instrument(options[selected])
            report_file = st.file_uploader("폐기 보고서 스캔본/PDF 업로드", type=["pdf", "png", "jpg", "jpeg"], key="disposal_upload")
            guessed_report_no = guess_document_no(report_file)
            status = cloud_status()
            if report_file and status.vision_ready and report_file.type and report_file.type.startswith("image/"):
                if st.button("OCR로 폐기 보고서 번호 후보 읽기"):
                    try:
                        text = extract_text_with_vision(report_file.name, report_file.getvalue())
                        candidates = parse_certificate_candidates(text)
                        guessed_report_no = candidates.get("certificate_no", guessed_report_no)
                        st.write("OCR 후보:", candidates)
                    except Exception as exc:
                        st.error(f"OCR 처리 실패: {exc}")
            report_no = st.text_input("폐기 보고서 번호", value=guessed_report_no)
            note = st.text_area("폐기 사유/비고")
            st.caption("Google Vision OCR 설정이 있으면 이미지에서 번호 후보를 읽고, 없으면 파일명에 들어간 번호를 후보로 가져옵니다.")
            if st.button("폐기 처리"):
                file_path = save_certificate_file(report_file, instrument["management_no"], "disposal_reports") if report_file else ""
                mark_disposed(options[selected], note or "폐기 처리", report_no, file_path)
                st.success("폐기 계측기로 분리했습니다.")

elif page == "담당자 관리":
    st.subheader("사용부서 담당자 관리")
    st.info("맞습니다. 여기서 사용부서별 담당자와 카카오톡 대상/방 이름을 입력하고 저장하면, 알림 문구 생성에 반영됩니다.")
    st.dataframe(contacts_df(), use_container_width=True, hide_index=True)
    disabled = not is_qc()
    with st.form("contact_form"):
        c1, c2, c3 = st.columns(3)
        department = c1.text_input("사용부서", disabled=disabled)
        owner_name = c2.text_input("담당자명", disabled=disabled)
        kakao_target = c3.text_input("카카오톡 대상/방 이름", disabled=disabled)
        phone = st.text_input("연락처", disabled=disabled)
        note = st.text_area("비고", disabled=disabled)
        submitted = st.form_submit_button("담당자 저장", disabled=disabled)
    if submitted and require_qc():
        if department:
            upsert_contact(department, owner_name, kakao_target, phone, note)
            st.success("담당자 정보를 저장했습니다.")
        else:
            st.error("사용부서를 입력하세요.")

elif page == "클라우드 설정":
    st.subheader("Google Cloud / Drive 운영 설정")
    status = cloud_status()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Google 라이브러리", "OK" if status.google_libs else "미설치")
    c2.metric("서비스 계정", "OK" if status.credentials else "미설정")
    c3.metric("Drive 폴더", "OK" if status.drive_folder_id else "미설정")
    c4.metric("Sheets 문서", "OK" if status.spreadsheet_id else "미설정")
    st.info(status.message)

    st.markdown("**필요한 Streamlit secrets / Cloud Run 환경변수**")
    st.code(
        """QC_APP_PASSWORD=QC2026
GOOGLE_DRIVE_FOLDER_ID=...
GOOGLE_SHEET_ID=...
GOOGLE_SERVICE_ACCOUNT_JSON={...서비스 계정 JSON 전체...}""",
        language="text",
    )
    st.caption("서비스 계정 이메일을 Google Drive 폴더와 Google Sheet에 편집자로 공유해야 앱이 파일/데이터를 저장할 수 있습니다.")

    if require_qc():
        st.markdown("**현재 로컬 DB를 Google Sheets로 내보내기**")
        if st.button("계측기/이력/담당자 시트를 Google Sheets로 동기화"):
            if not (status.google_libs and status.credentials and status.spreadsheet_id):
                st.error("Google Sheets 연동 설정이 아직 완료되지 않았습니다.")
            else:
                try:
                    replace_sheet_with_dataframe("instruments", instruments_df(include_disposed=True))
                    replace_sheet_with_dataframe("calibration_records", calibration_history_df())
                    replace_sheet_with_dataframe("department_contacts", contacts_df())
                    st.success("Google Sheets 동기화를 완료했습니다.")
                except Exception as exc:
                    st.error(f"Google Sheets 동기화 실패: {exc}")

elif page == "데이터 가져오기":
    st.subheader("기존 엑셀 대장 가져오기")
    st.caption(f"기본 경로: {DEFAULT_EXCEL_PATH}")
    st.info("반복 업로드해도 같은 계측기/구분/교정일/차기교정일/비고의 이력은 중복 추가하지 않고 기존 이력을 보완합니다. 전체 초기화가 필요할 때만 아래 체크박스를 사용하세요.")
    st.dataframe(get_import_log(), use_container_width=True, hide_index=True)
    if require_qc():
        reset = st.checkbox("기존 앱 DB를 비우고 다시 가져오기", help="초기 데이터 전체를 새 엑셀 기준으로 갈아엎을 때만 사용하세요.")
        uploaded = st.file_uploader("엑셀 파일 업로드", type=["xlsx"])
        if uploaded and st.button("업로드 파일 가져오기", type="primary"):
            temp_path = Path(__file__).parent / "data" / uploaded.name
            temp_path.parent.mkdir(parents=True, exist_ok=True)
            temp_path.write_bytes(uploaded.getbuffer())
            with st.spinner("엑셀 데이터를 가져오는 중입니다. 완료 메시지가 뜰 때까지 다시 누르지 마세요."):
                summary = import_excel(temp_path, reset=reset)
            st.success(f"{summary.instrument_count}개 계측기, {summary.record_count}개 이력을 처리했습니다. 기존 동일 이력은 중복 없이 보완됩니다.")
        if st.button("기본 경로 파일 가져오기"):
            if DEFAULT_EXCEL_PATH.exists():
                with st.spinner("기본 경로 엑셀 데이터를 가져오는 중입니다."):
                    summary = import_excel(DEFAULT_EXCEL_PATH, reset=reset)
                st.success(f"{summary.instrument_count}개 계측기, {summary.record_count}개 이력을 처리했습니다. 기존 동일 이력은 중복 없이 보완됩니다.")
            else:
                st.error("기본 경로에서 파일을 찾을 수 없습니다. 파일 업로드를 사용하세요.")
