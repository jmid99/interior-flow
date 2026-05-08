from datetime import datetime, timedelta, date
from io import BytesIO
import os
import uuid

import pandas as pd
import plotly.express as px
import streamlit as st
from PIL import Image
from supabase import create_client


# =========================================================
# Interior Flow v3 - Supabase DB + Storage 저장형
# 실행: streamlit run app.py
# 필요 패키지: streamlit, pandas, plotly, pillow, openpyxl, supabase
# =========================================================

APP_TITLE = "🏠 Interior Flow v3 - Supabase 저장형 다현장 협업 대시보드"
DEFAULT_BUCKET = "interior-photos"

st.set_page_config(page_title="Interior Flow v3", layout="wide")
st.title(APP_TITLE)
st.caption("여러 인테리어 현장의 공정, 이슈, 추가공사, 사진 기록을 Supabase DB와 Storage에 저장합니다.")


# -----------------------------
# Supabase 연결
# -----------------------------
@st.cache_resource
def get_supabase_client():
    try:
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
        return create_client(url, key)
    except Exception:
        st.error("Supabase 연결 정보가 없습니다. Streamlit Secrets에 SUPABASE_URL, SUPABASE_KEY, SUPABASE_BUCKET을 등록해 주세요.")
        st.stop()


supabase = get_supabase_client()
BUCKET = st.secrets.get("SUPABASE_BUCKET", DEFAULT_BUCKET)


# -----------------------------
# Supabase 유틸
# -----------------------------
def sb_select(table: str, order_col: str = "id") -> pd.DataFrame:
    try:
        res = supabase.table(table).select("*").order(order_col).execute()
        return pd.DataFrame(res.data or [])
    except Exception as e:
        st.error(f"{table} 데이터를 불러오지 못했습니다: {e}")
        return pd.DataFrame()


def sb_insert(table: str, row: dict):
    return supabase.table(table).insert(row).execute()


def sb_update(table: str, row_id: int, values: dict):
    return supabase.table(table).update(values).eq("id", row_id).execute()


def sb_delete(table: str, row_id: int):
    return supabase.table(table).delete().eq("id", row_id).execute()


def to_int(value, default=0):
    try:
        if pd.isna(value):
            return default
        return int(value)
    except Exception:
        return default


def format_currency(value):
    """금액을 1,000원 단위 콤마가 포함된 원화 문자열로 변환합니다."""
    try:
        if value is None or value == "":
            return "0원"
        cleaned = str(value).replace(",", "").replace("원", "").strip()
        if cleaned == "":
            return "0원"
        return f"{int(float(cleaned)):,}원"
    except Exception:
        return "0원"


def parse_currency(value):
    """50,000,000원 또는 50000000 형태의 입력값을 정수로 변환합니다."""
    try:
        if value is None or value == "":
            return 0
        cleaned = str(value).replace(",", "").replace("원", "").strip()
        if cleaned == "":
            return 0
        return int(float(cleaned))
    except Exception:
        return 0


def normalize_currency_input(key):
    """금액 입력칸에서 Tab/Enter/포커스 이동 시 50,000,000원 형태로 자동 정리합니다."""
    current_value = st.session_state.get(key, "0")
    st.session_state[key] = format_currency(current_value)


def currency_input(label, value=0, key=None):
    """숫자를 입력한 뒤 Tab/Enter를 누르면 콤마와 '원'이 자동 적용되는 금액 입력칸입니다."""
    if key is None:
        key = f"currency_{label}"

    if key not in st.session_state:
        st.session_state[key] = format_currency(value)

    st.text_input(
        label,
        key=key,
        on_change=normalize_currency_input,
        args=(key,),
        help="숫자만 입력해도 됩니다. 입력 후 Tab 또는 Enter를 누르면 50,000,000원 형태로 자동 정리됩니다.",
    )
    return parse_currency(st.session_state.get(key, "0"))


def format_money_columns(df, columns):
    """데이터프레임의 금액 컬럼을 콤마 포함 원화 문자열로 변환합니다."""
    if df is None or df.empty:
        return df
    result = df.copy()
    for col in columns:
        if col in result.columns:
            result[col] = result[col].apply(format_currency)
    return result


def to_date_str(value):
    if value in [None, "", pd.NaT]:
        return None
    try:
        return str(pd.to_datetime(value).date())
    except Exception:
        return None


def safe_date(value, default=None):
    if default is None:
        default = datetime.now().date()
    try:
        parsed = pd.to_datetime(value, errors="coerce")
        if pd.isna(parsed):
            return default
        return parsed.date()
    except Exception:
        return default


# -----------------------------
# 기본 공정표
# -----------------------------
def default_process_rows(project_id: int, start_date: date):
    processes = [
        ("철거", 0, 3, "김팀장"),
        ("전기/배선", 3, 8, "박기사"),
        ("설비/배관", 5, 12, "이설비"),
        ("미장/창호", 10, 18, "최미장"),
        ("타일/바닥", 16, 24, "정타일"),
        ("목공/가구", 22, 30, "윤목공"),
        ("도장/마감", 28, 35, "한도장"),
        ("조명/기구", 34, 39, "송조명"),
        ("청소 및 검수", 39, 45, "관리자"),
    ]
    rows = []
    for idx, (process, start_offset, end_offset, manager) in enumerate(processes):
        status = "대기"
        progress = 0
        actual_start = None
        actual_end = None
        next_action = ""
        approval_required = "아니오"
        approval_status = "해당없음"

        if idx == 0:
            status = "완료"
            progress = 100
            actual_start = str(start_date)
            actual_end = str(start_date + timedelta(days=3))
        elif idx == 1:
            status = "진행중"
            progress = 65
            actual_start = str(start_date + timedelta(days=3))
            next_action = "콘센트 위치 고객 확인"
            approval_required = "예"
            approval_status = "대기"

        rows.append(
            {
                "project_id": project_id,
                "process": process,
                "status": status,
                "manager": manager,
                "progress": progress,
                "planned_start": str(start_date + timedelta(days=start_offset)),
                "planned_end": str(start_date + timedelta(days=end_offset)),
                "actual_start": actual_start,
                "actual_end": actual_end,
                "material_status": "미확인",
                "approval_required": approval_required,
                "approval_status": approval_status,
                "delay_reason": "",
                "next_action": next_action,
                "memo": "",
            }
        )
    return rows


def create_default_project_if_empty():
    projects = sb_select("projects")
    if not projects.empty:
        return
    today = datetime.now().date()
    res = sb_insert(
        "projects",
        {
            "name": "해운대 아파트 34평 리모델링",
            "address": "부산 해운대구 ○○동",
            "client_name": "홍길동",
            "contract_amount": 50000000,
            "deposit_amount": 10000000,
            "first_payment": 10000000,
            "second_payment": 10000000,
            "third_payment": 10000000,
            "balance_amount": 10000000,
            "start_date": str(today),
            "duration_days": 45,
            "status": "진행중",
            "pm": "박민서",
            "memo": "기본 예시 현장",
        },
    )
    project_id = res.data[0]["id"]
    supabase.table("tasks").insert(default_process_rows(project_id, today)).execute()


# -----------------------------
# 데이터 로딩
# -----------------------------
def load_all_data():
    projects = sb_select("projects")
    tasks = sb_select("tasks")
    issues = sb_select("issues")
    change_orders = sb_select("change_orders")
    photos = sb_select("photos")
    return projects, tasks, issues, change_orders, photos


try:
    create_default_project_if_empty()
except Exception as e:
    st.error(f"기본 현장 생성 중 오류가 발생했습니다: {e}")
    st.stop()

projects_df, tasks_df, issues_df, change_orders_df, photos_df = load_all_data()


# -----------------------------
# 컬럼명 한글 표시용
# -----------------------------
def display_projects(df):
    if df.empty:
        return df
    return df.rename(
        columns={
            "id": "프로젝트ID",
            "name": "프로젝트명",
            "address": "현장주소",
            "client_name": "고객명",
            "contract_amount": "총계약금액",
            "deposit_amount": "계약금",
            "first_payment": "1차지급",
            "second_payment": "2차지급",
            "third_payment": "3차지급",
            "balance_amount": "잔금",
            "start_date": "공사시작일",
            "duration_days": "예상공사기간",
            "status": "상태",
            "pm": "PM",
            "memo": "메모",
        }
    )


def display_tasks(df):
    if df.empty:
        return df
    return df.rename(
        columns={
            "id": "ID",
            "project_id": "프로젝트ID",
            "process": "공정",
            "status": "진행상태",
            "manager": "담당자",
            "progress": "진행률",
            "planned_start": "시작예정",
            "planned_end": "완료예정",
            "actual_start": "실제시작일",
            "actual_end": "실제완료일",
            "material_status": "자재상태",
            "approval_required": "고객승인필요",
            "approval_status": "고객승인상태",
            "delay_reason": "지연사유",
            "next_action": "다음액션",
            "memo": "메모",
        }
    )


def display_issues(df):
    if df.empty:
        return df
    return df.rename(
        columns={
            "id": "ID",
            "project_id": "프로젝트ID",
            "process": "공정",
            "issue_type": "이슈유형",
            "priority": "중요도",
            "content": "내용",
            "manager": "담당자",
            "status": "상태",
            "registered_date": "등록일",
            "due_date": "처리기한",
            "result": "처리내용",
        }
    )


def display_changes(df):
    if df.empty:
        return df
    return df.rename(
        columns={
            "id": "ID",
            "project_id": "프로젝트ID",
            "process": "공정",
            "requester": "요청자",
            "content": "변경내용",
            "amount": "추가금액",
            "approval_status": "승인상태",
            "approval_method": "승인방식",
            "request_date": "요청일",
            "approval_date": "승인일",
            "memo": "메모",
        }
    )


def display_photos(df):
    if df.empty:
        return df
    return df.rename(
        columns={
            "id": "ID",
            "project_id": "프로젝트ID",
            "process": "공정",
            "photo_type": "사진구분",
            "file_name": "파일명",
            "storage_path": "저장경로",
            "description": "설명",
            "uploaded_at": "업로드일시",
        }
    )


def task_values_from_display(row):
    return {
        "process": str(row.get("공정", "")),
        "status": str(row.get("진행상태", "대기")),
        "manager": str(row.get("담당자", "")),
        "progress": to_int(row.get("진행률", 0)),
        "planned_start": to_date_str(row.get("시작예정")),
        "planned_end": to_date_str(row.get("완료예정")),
        "actual_start": to_date_str(row.get("실제시작일")),
        "actual_end": to_date_str(row.get("실제완료일")),
        "material_status": str(row.get("자재상태", "미확인")),
        "approval_required": str(row.get("고객승인필요", "아니오")),
        "approval_status": str(row.get("고객승인상태", "해당없음")),
        "delay_reason": str(row.get("지연사유", "")),
        "next_action": str(row.get("다음액션", "")),
        "memo": str(row.get("메모", "")),
    }


# -----------------------------
# 사이드바: 현장 선택/추가
# -----------------------------
st.sidebar.header("현장 관리")

if projects_df.empty:
    st.sidebar.warning("등록된 현장이 없습니다.")
    st.stop()

project_options = {f"[{int(row['id'])}] {row['name']}": int(row["id"]) for _, row in projects_df.iterrows()}
selected_label = st.sidebar.selectbox("현재 현장 선택", list(project_options.keys()))
selected_project_id = project_options[selected_label]
selected_project = projects_df[projects_df["id"] == selected_project_id].iloc[0]

with st.sidebar.expander("➕ 새 현장 추가", expanded=False):
    with st.form("new_project_form", clear_on_submit=True):
        new_name = st.text_input("프로젝트명", "수영 상가 인테리어")
        new_address = st.text_input("현장주소", "부산 수영구 ○○로")
        new_client = st.text_input("고객명", "고객명")
        new_amount = currency_input("총 공사대금", 30000000, key="new_amount")
        new_deposit = currency_input("계약금", 6000000, key="new_deposit")
        new_first = currency_input("1차 지급", 6000000, key="new_first")
        new_second = currency_input("2차 지급", 6000000, key="new_second")
        new_third = currency_input("3차 지급", 6000000, key="new_third")
        new_balance = currency_input("잔금", 6000000, key="new_balance")
        new_start = st.date_input("공사시작일", datetime.now().date())
        new_duration = st.slider("예상공사기간(일)", 7, 120, 30)
        new_pm = st.text_input("PM", "박민서")
        new_memo = st.text_area("메모")
        create_project = st.form_submit_button("현장 추가")

        if create_project:
            try:
                res = sb_insert(
                    "projects",
                    {
                        "name": new_name,
                        "address": new_address,
                        "client_name": new_client,
                        "contract_amount": int(new_amount),
                        "deposit_amount": int(new_deposit),
                        "first_payment": int(new_first),
                        "second_payment": int(new_second),
                        "third_payment": int(new_third),
                        "balance_amount": int(new_balance),
                        "start_date": str(new_start),
                        "duration_days": int(new_duration),
                        "status": "진행중",
                        "pm": new_pm,
                        "memo": new_memo,
                    },
                )
                new_project_id = res.data[0]["id"]
                supabase.table("tasks").insert(default_process_rows(new_project_id, new_start)).execute()
                st.success("새 현장과 기본 공정표가 Supabase에 저장되었습니다.")
                st.rerun()
            except Exception as e:
                st.error(f"현장 추가 실패: {e}")


# -----------------------------
# 선택 현장 데이터
# -----------------------------
project_tasks = tasks_df[tasks_df["project_id"] == selected_project_id] if not tasks_df.empty else pd.DataFrame()
project_issues = issues_df[issues_df["project_id"] == selected_project_id] if not issues_df.empty else pd.DataFrame()
project_changes = change_orders_df[change_orders_df["project_id"] == selected_project_id] if not change_orders_df.empty else pd.DataFrame()
project_photos = photos_df[photos_df["project_id"] == selected_project_id] if not photos_df.empty else pd.DataFrame()

project_name = str(selected_project.get("name", ""))
site_address = str(selected_project.get("address", ""))
client_name = str(selected_project.get("client_name", ""))
contract_amount = to_int(selected_project.get("contract_amount", 0))
deposit_amount = to_int(selected_project.get("deposit_amount", 0))
first_payment = to_int(selected_project.get("first_payment", 0))
second_payment = to_int(selected_project.get("second_payment", 0))
third_payment = to_int(selected_project.get("third_payment", 0))
balance_amount = to_int(selected_project.get("balance_amount", 0))
start_date = safe_date(selected_project.get("start_date"))
expected_duration = to_int(selected_project.get("duration_days", 45), 45)
expected_end_date = start_date + timedelta(days=expected_duration)


# -----------------------------
# 전체 현장 요약
# -----------------------------
def calculate_summary(projects, tasks, issues, changes):
    rows = []
    for _, p in projects.iterrows():
        pid = int(p["id"])
        t = tasks[tasks["project_id"] == pid] if not tasks.empty else pd.DataFrame()
        i = issues[issues["project_id"] == pid] if not issues.empty else pd.DataFrame()
        c = changes[changes["project_id"] == pid] if not changes.empty else pd.DataFrame()
        progress = int(pd.to_numeric(t["progress"], errors="coerce").fillna(0).mean()) if not t.empty else 0
        open_issues = len(i[i["status"].isin(["접수", "처리중"])]) if not i.empty else 0
        pending_changes = len(c[c["approval_status"] == "대기"]) if not c.empty else 0
        approved_amount = int(pd.to_numeric(c.loc[c["approval_status"] == "승인", "amount"], errors="coerce").fillna(0).sum()) if not c.empty else 0
        rows.append(
            {
                "프로젝트ID": pid,
                "프로젝트명": p.get("name", ""),
                "상태": p.get("status", ""),
                "PM": p.get("pm", ""),
                "진행률": progress,
                "미처리이슈": open_issues,
                "미승인추가공사": pending_changes,
                "승인추가금액": approved_amount,
                "총계약금액": to_int(p.get("contract_amount", 0)),
                "계약금": to_int(p.get("deposit_amount", 0)),
                "1차지급": to_int(p.get("first_payment", 0)),
                "2차지급": to_int(p.get("second_payment", 0)),
                "3차지급": to_int(p.get("third_payment", 0)),
                "잔금": to_int(p.get("balance_amount", 0)),
                "고객명": p.get("client_name", ""),
                "현장주소": p.get("address", ""),
            }
        )
    return pd.DataFrame(rows)


summary_df = calculate_summary(projects_df, tasks_df, issues_df, change_orders_df)


# -----------------------------
# 상단 요약
# -----------------------------
overall_progress = int(pd.to_numeric(project_tasks["progress"], errors="coerce").fillna(0).mean()) if not project_tasks.empty else 0
in_progress_count = len(project_tasks[project_tasks["status"] == "진행중"]) if not project_tasks.empty else 0
done_count = len(project_tasks[project_tasks["status"] == "완료"]) if not project_tasks.empty else 0
open_issues_count = len(project_issues[project_issues["status"].isin(["접수", "처리중"])]) if not project_issues.empty else 0
pending_change_count = len(project_changes[project_changes["approval_status"] == "대기"]) if not project_changes.empty else 0

if not project_tasks.empty:
    end_dates = pd.to_datetime(project_tasks["planned_end"], errors="coerce")
    max_end = end_dates.max()
    remaining_days = (max_end.date() - datetime.now().date()).days if pd.notna(max_end) else 0
else:
    remaining_days = 0

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("선택 현장 진행률", f"{overall_progress}%")
col2.metric("진행중 공정", f"{in_progress_count}개")
col3.metric("완료 공정", f"{done_count}개")
col4.metric("미처리 이슈", f"{open_issues_count}건")
col5.metric("미승인 추가공사", f"{pending_change_count}건")

st.info(
    f"현재 현장: {project_name} / 주소: {site_address} / 고객: {client_name} / "
    f"총 공사대금: {format_currency(contract_amount)} / "
    f"계약금: {format_currency(deposit_amount)} / "
    f"1차 지급: {format_currency(first_payment)} / "
    f"2차 지급: {format_currency(second_payment)} / "
    f"3차 지급: {format_currency(third_payment)} / "
    f"잔금: {format_currency(balance_amount)} / "
    f"예정 준공일: {expected_end_date} / 남은 예정일: {remaining_days}일"
)

st.divider()


# -----------------------------
# 탭 구성
# -----------------------------
tab0, tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(
    ["🏢 전체 현장", "📊 선택 현장", "📋 공정 Kanban", "📅 타임라인", "⚠️ 이슈", "💰 추가공사", "📸 사진/증거", "⚙️ 데이터 관리"]
)


# -----------------------------
# 탭0: 전체 현장
# -----------------------------
with tab0:
    st.subheader("전체 현장 요약")
    summary_display_df = format_money_columns(
        summary_df,
        ["총공사대금", "총 공사대금", "계약금", "1차지급", "1차 지급", "2차지급", "2차 지급", "3차지급", "3차 지급", "잔금", "승인추가금액", "승인 추가금액"],
    )
    st.dataframe(summary_display_df, use_container_width=True, hide_index=True)

    if not summary_df.empty:
        fig_all = px.bar(summary_df, x="프로젝트명", y="진행률", color="상태", text="진행률", hover_data=["PM", "미처리이슈", "미승인추가공사"])
        fig_all.update_layout(height=420)
        st.plotly_chart(fig_all, use_container_width=True)

    st.subheader("선택 현장 기본정보 수정")
    with st.form("edit_project_form"):
        edit_name = st.text_input("프로젝트명", project_name)
        edit_address = st.text_input("현장주소", site_address)
        edit_client = st.text_input("고객명", client_name)
        edit_amount = currency_input("총 공사대금", contract_amount, key="edit_amount")
        edit_deposit = currency_input("계약금", deposit_amount, key="edit_deposit")
        edit_first = currency_input("1차 지급", first_payment, key="edit_first")
        edit_second = currency_input("2차 지급", second_payment, key="edit_second")
        edit_third = currency_input("3차 지급", third_payment, key="edit_third")
        edit_balance = currency_input("잔금", balance_amount, key="edit_balance")
        st.caption(f"지급단계 합계: {format_currency(edit_deposit + edit_first + edit_second + edit_third + edit_balance)}")
        edit_start = st.date_input("공사시작일", start_date)
        edit_duration = st.number_input("예상공사기간", min_value=1, value=expected_duration, step=1)
        edit_status = st.selectbox("상태", ["준비중", "진행중", "보류", "완료", "취소"], index=["준비중", "진행중", "보류", "완료", "취소"].index(str(selected_project.get("status", "진행중"))) if str(selected_project.get("status", "진행중")) in ["준비중", "진행중", "보류", "완료", "취소"] else 1)
        edit_pm = st.text_input("PM", str(selected_project.get("pm", "")))
        edit_memo = st.text_area("메모", str(selected_project.get("memo", "")))
        save_project = st.form_submit_button("💾 선택 현장 기본정보 저장")

        if save_project:
            try:
                sb_update(
                    "projects",
                    selected_project_id,
                    {
                        "name": edit_name,
                        "address": edit_address,
                        "client_name": edit_client,
                        "contract_amount": int(edit_amount),
                        "deposit_amount": int(edit_deposit),
                        "first_payment": int(edit_first),
                        "second_payment": int(edit_second),
                        "third_payment": int(edit_third),
                        "balance_amount": int(edit_balance),
                        "start_date": str(edit_start),
                        "duration_days": int(edit_duration),
                        "status": edit_status,
                        "pm": edit_pm,
                        "memo": edit_memo,
                    },
                )
                st.success("현장 기본정보가 저장되었습니다.")
                st.rerun()
            except Exception as e:
                st.error(f"저장 실패: {e}")


# -----------------------------
# 탭1: 선택 현장 대시보드
# -----------------------------
with tab1:
    st.subheader("선택 현장 공정별 진행률")

    if project_tasks.empty:
        st.warning("이 현장의 공정 데이터가 없습니다.")
        if st.button("현재 현장 기본 공정표 생성"):
            supabase.table("tasks").insert(default_process_rows(selected_project_id, start_date)).execute()
            st.rerun()
    else:
        display_project_tasks = display_tasks(project_tasks)
        fig = px.bar(
            display_project_tasks,
            x="공정",
            y="진행률",
            color="진행상태",
            color_discrete_map={"완료": "#2ecc71", "진행중": "#f1c40f", "대기": "#95a5a6", "보류": "#e74c3c"},
            text="진행률",
            hover_data=["담당자", "시작예정", "완료예정", "다음액션"],
        )
        fig.update_layout(height=420)
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("오늘 확인할 항목")
        check_df = display_project_tasks[
            (display_project_tasks["진행상태"].isin(["진행중", "보류"]))
            | (display_project_tasks["고객승인상태"].isin(["대기", "보류"]))
            | (display_project_tasks["지연사유"].fillna("") != "")
        ]
        if check_df.empty:
            st.success("현재 특별히 확인할 항목이 없습니다.")
        else:
            st.dataframe(check_df[["공정", "진행상태", "진행률", "담당자", "고객승인상태", "지연사유", "다음액션"]], use_container_width=True, hide_index=True)


# -----------------------------
# 탭2: Kanban + 공정 직접 수정
# -----------------------------
with tab2:
    st.subheader("공정 Kanban 보드")
    status_order = ["대기", "진행중", "완료", "보류"]
    cols = st.columns(4)

    if project_tasks.empty:
        st.info("등록된 공정이 없습니다.")
    else:
        for idx, status in enumerate(status_order):
            with cols[idx]:
                st.markdown(f"### {status}")
                filtered = project_tasks[project_tasks["status"] == status]
                if filtered.empty:
                    st.caption("해당 공정 없음")
                for _, row in filtered.iterrows():
                    row_id = int(row["id"])
                    with st.expander(f"{row['process']} / {row.get('manager','')} / {row.get('progress',0)}%"):
                        new_progress = st.slider("진행률", 0, 100, to_int(row.get("progress", 0)), key=f"prog_{row_id}")
                        new_status = st.selectbox("상태", status_order, index=status_order.index(row.get("status", "대기")) if row.get("status", "대기") in status_order else 0, key=f"stat_{row_id}")
                        next_action = st.text_input("다음 액션", str(row.get("next_action", "")), key=f"next_{row_id}")
                        delay_reason = st.text_area("지연사유", str(row.get("delay_reason", "")), key=f"delay_{row_id}")

                        if st.button("저장", key=f"save_task_{row_id}"):
                            values = {
                                "progress": int(new_progress),
                                "status": new_status,
                                "next_action": next_action,
                                "delay_reason": delay_reason,
                            }
                            if new_status == "완료" and not row.get("actual_end"):
                                values["actual_end"] = str(datetime.now().date())
                            sb_update("tasks", row_id, values)
                            st.success("공정 정보가 저장되었습니다.")
                            st.rerun()

    st.divider()
    st.subheader("선택 현장 공정표 직접 수정")
    if not project_tasks.empty:
        edit_df = display_tasks(project_tasks).copy()
        editable_cols = ["ID", "공정", "진행상태", "담당자", "진행률", "시작예정", "완료예정", "실제시작일", "실제완료일", "자재상태", "고객승인필요", "고객승인상태", "지연사유", "다음액션", "메모"]
        edit_df = edit_df[editable_cols]
        edited_tasks = st.data_editor(
            edit_df,
            use_container_width=True,
            num_rows="fixed",
            hide_index=True,
            column_config={
                "진행상태": st.column_config.SelectboxColumn("진행상태", options=["대기", "진행중", "완료", "보류"]),
                "자재상태": st.column_config.SelectboxColumn("자재상태", options=["미확인", "발주전", "발주완료", "입고완료", "문제발생"]),
                "고객승인필요": st.column_config.SelectboxColumn("고객승인필요", options=["예", "아니오"]),
                "고객승인상태": st.column_config.SelectboxColumn("고객승인상태", options=["해당없음", "대기", "승인", "거절", "보류"]),
            },
        )

        if st.button("💾 선택 현장 공정표 저장"):
            try:
                for _, row in edited_tasks.iterrows():
                    sb_update("tasks", int(row["ID"]), task_values_from_display(row))
                st.success("공정표가 Supabase에 저장되었습니다.")
                st.rerun()
            except Exception as e:
                st.error(f"공정표 저장 실패: {e}")


# -----------------------------
# 탭3: 타임라인
# -----------------------------
with tab3:
    st.subheader("선택 현장 공사 타임라인")
    if project_tasks.empty:
        st.warning("표시할 공정이 없습니다.")
    else:
        timeline_df = display_tasks(project_tasks).copy()
        timeline_df["시작예정"] = pd.to_datetime(timeline_df["시작예정"], errors="coerce")
        timeline_df["완료예정"] = pd.to_datetime(timeline_df["완료예정"], errors="coerce")
        timeline_df = timeline_df.dropna(subset=["시작예정", "완료예정"])
        if timeline_df.empty:
            st.warning("시작예정/완료예정 날짜를 확인해 주세요.")
        else:
            fig_timeline = px.timeline(timeline_df, x_start="시작예정", x_end="완료예정", y="공정", color="진행상태", hover_data=["담당자", "진행률", "다음액션"])
            fig_timeline.update_yaxes(autorange="reversed")
            fig_timeline.update_layout(height=520)
            st.plotly_chart(fig_timeline, use_container_width=True)


# -----------------------------
# 탭4: 이슈 관리
# -----------------------------
with tab4:
    st.subheader("이슈 등록")
    process_options = list(project_tasks["process"].unique()) if not project_tasks.empty else ["공통"]

    with st.form("issue_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        issue_process = c1.selectbox("공정", process_options)
        issue_type = c2.selectbox("이슈유형", ["일정지연", "자재문제", "하자", "고객요청", "작업자문제", "안전", "기타"])
        issue_level = c3.selectbox("중요도", ["낮음", "보통", "높음", "긴급"])
        issue_content = st.text_area("내용")
        c4, c5, c6 = st.columns(3)
        issue_manager = c4.text_input("담당자")
        issue_status = c5.selectbox("상태", ["접수", "처리중", "완료", "보류"])
        issue_due = c6.date_input("처리기한", datetime.now().date() + timedelta(days=3))
        issue_result = st.text_area("처리내용")

        if st.form_submit_button("이슈 저장"):
            try:
                sb_insert(
                    "issues",
                    {
                        "project_id": selected_project_id,
                        "process": issue_process,
                        "issue_type": issue_type,
                        "priority": issue_level,
                        "content": issue_content,
                        "manager": issue_manager,
                        "status": issue_status,
                        "registered_date": str(datetime.now().date()),
                        "due_date": str(issue_due),
                        "result": issue_result,
                    },
                )
                st.success("이슈가 Supabase에 저장되었습니다.")
                st.rerun()
            except Exception as e:
                st.error(f"이슈 저장 실패: {e}")

    st.divider()
    st.subheader("선택 현장 이슈 목록")
    st.dataframe(display_issues(project_issues), use_container_width=True, hide_index=True)


# -----------------------------
# 탭5: 추가공사 관리
# -----------------------------
with tab5:
    st.subheader("추가공사/변경 요청 등록")
    st.caption("분쟁 예방을 위해 변경내용, 추가금액, 승인방식, 승인일을 반드시 남기는 것을 권장합니다.")
    process_options = list(project_tasks["process"].unique()) if not project_tasks.empty else ["공통"]

    with st.form("change_order_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        co_process = c1.selectbox("공정", process_options, key="co_process")
        requester = c2.selectbox("요청자", ["고객", "시공자", "협력업체", "기타"])
        amount = c3.number_input("추가금액", min_value=0, value=0, step=100000, format="%d")
        change_content = st.text_area("변경내용")
        c4, c5, c6 = st.columns(3)
        approval_status = c4.selectbox("승인상태", ["대기", "승인", "거절", "보류"])
        approval_method = c5.selectbox("승인방식", ["미정", "구두", "카카오톡", "문자", "이메일", "서명", "계약서/견적서"])
        request_date = c6.date_input("요청일", datetime.now().date())
        approval_date = st.date_input("승인일", datetime.now().date())
        co_memo = st.text_area("메모")

        if st.form_submit_button("추가공사 저장"):
            try:
                sb_insert(
                    "change_orders",
                    {
                        "project_id": selected_project_id,
                        "process": co_process,
                        "requester": requester,
                        "content": change_content,
                        "amount": int(amount),
                        "approval_status": approval_status,
                        "approval_method": approval_method,
                        "request_date": str(request_date),
                        "approval_date": str(approval_date) if approval_status == "승인" else None,
                        "memo": co_memo,
                    },
                )
                st.success("추가공사 기록이 Supabase에 저장되었습니다.")
                st.rerun()
            except Exception as e:
                st.error(f"추가공사 저장 실패: {e}")

    st.divider()
    st.subheader("선택 현장 추가공사 목록")
    if not project_changes.empty:
        total_approved = pd.to_numeric(project_changes.loc[project_changes["approval_status"] == "승인", "amount"], errors="coerce").fillna(0).sum()
        st.metric("승인된 추가공사 합계", format_currency(total_approved))
    st.dataframe(display_changes(project_changes), use_container_width=True, hide_index=True)


# -----------------------------
# 탭6: 사진/증거 관리
# -----------------------------
with tab6:
    st.subheader("현장 사진/증거 업로드")
    st.caption("Supabase Storage에 사진을 저장하고, DB에는 사진 기록을 남깁니다.")
    process_options = list(project_tasks["process"].unique()) if not project_tasks.empty else ["공통"]

    c1, c2, c3 = st.columns(3)
    photo_process = c1.selectbox("공정", process_options, key="photo_process")
    photo_type = c2.selectbox("사진구분", ["작업 전", "작업 중", "작업 후", "하자", "자재", "고객요청", "기타"])
    photo_description = c3.text_input("사진 설명")

    uploaded_files = st.file_uploader("사진 업로드 JPG/PNG/JPEG", type=["jpg", "png", "jpeg"], accept_multiple_files=True)

    if uploaded_files and st.button("📸 사진 저장"):
        try:
            for uploaded_file in uploaded_files:
                ext = os.path.splitext(uploaded_file.name)[1].lower()
                unique_name = f"project_{selected_project_id}/{datetime.now().strftime('%Y%m%d')}/{uuid.uuid4().hex}{ext}"
                file_bytes = uploaded_file.getvalue()

                supabase.storage.from_(BUCKET).upload(
                    unique_name,
                    file_bytes,
                    {"content-type": uploaded_file.type, "upsert": "true"},
                )

                sb_insert(
                    "photos",
                    {
                        "project_id": selected_project_id,
                        "process": photo_process,
                        "photo_type": photo_type,
                        "file_name": uploaded_file.name,
                        "storage_path": unique_name,
                        "description": photo_description,
                    },
                )
            st.success("사진이 Supabase Storage에 저장되었습니다.")
            st.rerun()
        except Exception as e:
            st.error(f"사진 저장 실패: {e}")

    st.divider()
    st.subheader("선택 현장 사진 기록 목록")
    st.dataframe(display_photos(project_photos), use_container_width=True, hide_index=True)

    st.subheader("사진 미리보기")
    if project_photos.empty:
        st.info("아직 저장된 사진이 없습니다.")
    else:
        photo_cols = st.columns(3)
        recent_photos = project_photos.tail(12).iloc[::-1]
        for idx, (_, row) in enumerate(recent_photos.iterrows()):
            with photo_cols[idx % 3]:
                try:
                    public_url = supabase.storage.from_(BUCKET).get_public_url(row["storage_path"])
                    st.image(public_url, caption=f"{row.get('process', '')} / {row.get('photo_type', '')} / {row.get('description', '')}", use_container_width=True)
                except Exception:
                    st.warning("사진 URL을 불러오지 못했습니다.")


# -----------------------------
# 탭7: 데이터 관리
# -----------------------------
with tab7:
    st.subheader("전체 데이터 다운로드")

    projects_display = display_projects(projects_df)
    tasks_display = display_tasks(tasks_df)
    issues_display = display_issues(issues_df)
    changes_display = display_changes(change_orders_df)
    photos_display = display_photos(photos_df)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        projects_display.to_excel(writer, index=False, sheet_name="현장목록")
        tasks_display.to_excel(writer, index=False, sheet_name="공정현황")
        issues_display.to_excel(writer, index=False, sheet_name="이슈관리")
        changes_display.to_excel(writer, index=False, sheet_name="추가공사")
        photos_display.to_excel(writer, index=False, sheet_name="사진기록")
    output.seek(0)

    st.download_button(
        label="📥 전체 현황 엑셀 다운로드",
        data=output,
        file_name=f"Interior_Flow_v3_현황_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()
    st.subheader("연결 상태")
    st.success("Supabase DB 연결 상태: 정상")
    st.info(f"Storage bucket: {BUCKET}")
    st.warning("현재 RLS 정책은 테스트용으로 anon 전체 허용 상태입니다. 실사용 전에는 로그인, 사용자 권한, private bucket 구조로 전환해야 합니다.")

st.caption("Interior Flow v3 / Supabase DB + Storage 저장형 / 다음 단계: 로그인, 권한관리, 모바일 현장 입력 화면, AI 일일 리포트")
