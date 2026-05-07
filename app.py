import os
from io import BytesIO
from datetime import date, datetime, timedelta

import pandas as pd
import plotly.express as px
import streamlit as st
from PIL import Image


# =========================================================
# Interior Flow v2 - 여러 현장 관리 버전
# 실행 명령어: streamlit run app.py
# 배포: Streamlit Community Cloud
# =========================================================

APP_TITLE = "🏠 Interior Flow v2 - 인테리어 다현장 협업 대시보드"
DATA_DIR = "data"
PHOTO_DIR = os.path.join(DATA_DIR, "photos")
PROJECTS_FILE = os.path.join(DATA_DIR, "projects.csv")
TASKS_FILE = os.path.join(DATA_DIR, "tasks.csv")
ISSUES_FILE = os.path.join(DATA_DIR, "issues.csv")
CHANGE_ORDERS_FILE = os.path.join(DATA_DIR, "change_orders.csv")
PHOTOS_FILE = os.path.join(DATA_DIR, "photos.csv")


# -----------------------------
# 기본 설정
# -----------------------------
st.set_page_config(page_title="Interior Flow v2", layout="wide")
st.title(APP_TITLE)
st.caption("여러 인테리어 현장의 공정, 이슈, 추가공사, 사진 기록을 한 화면에서 관리하는 MVP입니다.")


# -----------------------------
# 폴더 생성
# -----------------------------
def ensure_dirs() -> None:
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(PHOTO_DIR, exist_ok=True)


# -----------------------------
# 기본 데이터 생성
# -----------------------------
def create_default_projects() -> pd.DataFrame:
    today = datetime.now().date()
    return pd.DataFrame(
        [
            {
                "프로젝트ID": 1,
                "프로젝트명": "해운대 아파트 34평 리모델링",
                "현장주소": "부산 해운대구 ○○동",
                "고객명": "홍길동",
                "계약금액": 50000000,
                "공사시작일": str(today),
                "예상공사기간": 45,
                "상태": "진행중",
                "PM": "박민서",
                "메모": "기본 예시 현장",
            }
        ]
    )


def create_default_tasks(project_id: int, project_name: str, start_date: date) -> pd.DataFrame:
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
    base_id = int(project_id) * 1000
    for idx, (process, start_offset, end_offset, manager) in enumerate(processes, start=1):
        rows.append(
            {
                "ID": base_id + idx,
                "프로젝트ID": project_id,
                "프로젝트명": project_name,
                "공정": process,
                "진행상태": "대기",
                "담당자": manager,
                "진행률": 0,
                "시작예정": str(start_date + timedelta(days=start_offset)),
                "완료예정": str(start_date + timedelta(days=end_offset)),
                "실제시작일": "",
                "실제완료일": "",
                "선행공정": "",
                "자재상태": "미확인",
                "고객승인필요": "아니오",
                "고객승인상태": "해당없음",
                "지연사유": "",
                "다음액션": "",
                "메모": "",
            }
        )

    if rows:
        rows[0]["진행상태"] = "완료"
        rows[0]["진행률"] = 100
        rows[0]["실제시작일"] = str(start_date)
        rows[0]["실제완료일"] = str(start_date + timedelta(days=3))
        rows[1]["진행상태"] = "진행중"
        rows[1]["진행률"] = 65
        rows[1]["실제시작일"] = str(start_date + timedelta(days=3))
        rows[1]["다음액션"] = "콘센트 위치 고객 확인"
        rows[1]["고객승인필요"] = "예"
        rows[1]["고객승인상태"] = "대기"

    return pd.DataFrame(rows)


def create_empty_issues() -> pd.DataFrame:
    return pd.DataFrame(
        columns=[
            "ID",
            "프로젝트ID",
            "프로젝트명",
            "공정",
            "이슈유형",
            "중요도",
            "내용",
            "담당자",
            "상태",
            "등록일",
            "처리기한",
            "처리내용",
        ]
    )


def create_empty_change_orders() -> pd.DataFrame:
    return pd.DataFrame(
        columns=[
            "ID",
            "프로젝트ID",
            "프로젝트명",
            "공정",
            "요청자",
            "변경내용",
            "추가금액",
            "승인상태",
            "승인방식",
            "요청일",
            "승인일",
            "메모",
        ]
    )


def create_empty_photos() -> pd.DataFrame:
    return pd.DataFrame(
        columns=[
            "ID",
            "프로젝트ID",
            "프로젝트명",
            "공정",
            "사진구분",
            "파일명",
            "저장경로",
            "설명",
            "업로드일시",
        ]
    )


# -----------------------------
# 저장/불러오기
# -----------------------------
def load_csv(path: str, default_df: pd.DataFrame) -> pd.DataFrame:
    if os.path.exists(path):
        try:
            return pd.read_csv(path)
        except Exception:
            return default_df
    return default_df


def save_csv(df: pd.DataFrame, path: str) -> None:
    df.to_csv(path, index=False, encoding="utf-8-sig")


def next_id(df: pd.DataFrame, id_col: str = "ID") -> int:
    if df.empty or id_col not in df.columns:
        return 1
    ids = pd.to_numeric(df[id_col], errors="coerce").dropna()
    if ids.empty:
        return 1
    return int(ids.max()) + 1


def make_excel_file(projects_df, tasks_df, issues_df, change_orders_df, photos_df) -> BytesIO:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        projects_df.to_excel(writer, index=False, sheet_name="현장목록")
        tasks_df.to_excel(writer, index=False, sheet_name="공정현황")
        issues_df.to_excel(writer, index=False, sheet_name="이슈관리")
        change_orders_df.to_excel(writer, index=False, sheet_name="추가공사")
        photos_df.to_excel(writer, index=False, sheet_name="사진기록")
    output.seek(0)
    return output


def safe_date(value, default=None):
    try:
        parsed = pd.to_datetime(value, errors="coerce")
        if pd.isna(parsed):
            return default or datetime.now().date()
        return parsed.date()
    except Exception:
        return default or datetime.now().date()


# -----------------------------
# 초기화
# -----------------------------
ensure_dirs()

if "projects" not in st.session_state:
    st.session_state["projects"] = load_csv(PROJECTS_FILE, create_default_projects())

if "tasks" not in st.session_state:
    projects_default = st.session_state["projects"]
    if projects_default.empty:
        default_tasks = pd.DataFrame()
    else:
        p = projects_default.iloc[0]
        default_tasks = create_default_tasks(int(p["프로젝트ID"]), str(p["프로젝트명"]), safe_date(p["공사시작일"]))
    st.session_state["tasks"] = load_csv(TASKS_FILE, default_tasks)

if "issues" not in st.session_state:
    st.session_state["issues"] = load_csv(ISSUES_FILE, create_empty_issues())

if "change_orders" not in st.session_state:
    st.session_state["change_orders"] = load_csv(CHANGE_ORDERS_FILE, create_empty_change_orders())

if "photos" not in st.session_state:
    st.session_state["photos"] = load_csv(PHOTOS_FILE, create_empty_photos())

projects_df = st.session_state["projects"]
tasks_df = st.session_state["tasks"]
issues_df = st.session_state["issues"]
change_orders_df = st.session_state["change_orders"]
photos_df = st.session_state["photos"]


# -----------------------------
# 사이드바: 현장 선택/추가
# -----------------------------
st.sidebar.header("현장 관리")

if projects_df.empty:
    st.sidebar.warning("등록된 현장이 없습니다. 새 현장을 추가해 주세요.")
    selected_project_id = None
    selected_project = None
else:
    project_options = {
        f"[{int(row['프로젝트ID'])}] {row['프로젝트명']}": int(row["프로젝트ID"])
        for _, row in projects_df.iterrows()
    }
    selected_label = st.sidebar.selectbox("현재 현장 선택", list(project_options.keys()))
    selected_project_id = project_options[selected_label]
    selected_project = projects_df[projects_df["프로젝트ID"] == selected_project_id].iloc[0]

with st.sidebar.expander("➕ 새 현장 추가", expanded=False):
    with st.form("new_project_form", clear_on_submit=True):
        new_name = st.text_input("프로젝트명", "수영 상가 인테리어")
        new_address = st.text_input("현장주소", "부산 수영구 ○○로")
        new_client = st.text_input("고객명", "고객명")
        new_amount = st.number_input("계약금액", min_value=0, value=30000000, step=1000000, format="%d")
        new_start = st.date_input("공사시작일", datetime.now().date())
        new_duration = st.slider("예상공사기간(일)", 7, 120, 30)
        new_pm = st.text_input("PM", "박민서")
        new_memo = st.text_area("메모")
        create_project = st.form_submit_button("현장 추가")

        if create_project:
            new_project_id = next_id(projects_df, "프로젝트ID")
            new_project = {
                "프로젝트ID": new_project_id,
                "프로젝트명": new_name,
                "현장주소": new_address,
                "고객명": new_client,
                "계약금액": new_amount,
                "공사시작일": str(new_start),
                "예상공사기간": new_duration,
                "상태": "진행중",
                "PM": new_pm,
                "메모": new_memo,
            }
            projects_df = pd.concat([projects_df, pd.DataFrame([new_project])], ignore_index=True)
            new_tasks = create_default_tasks(new_project_id, new_name, new_start)
            tasks_df = pd.concat([tasks_df, new_tasks], ignore_index=True)

            st.session_state["projects"] = projects_df
            st.session_state["tasks"] = tasks_df
            save_csv(projects_df, PROJECTS_FILE)
            save_csv(tasks_df, TASKS_FILE)
            st.success("새 현장과 기본 공정표가 생성되었습니다.")
            st.rerun()


# -----------------------------
# 선택 현장 데이터
# -----------------------------
if selected_project_id is None:
    st.stop()

project_tasks = tasks_df[tasks_df["프로젝트ID"] == selected_project_id]
project_issues = issues_df[issues_df["프로젝트ID"] == selected_project_id] if not issues_df.empty else issues_df
project_changes = change_orders_df[change_orders_df["프로젝트ID"] == selected_project_id] if not change_orders_df.empty else change_orders_df
project_photos = photos_df[photos_df["프로젝트ID"] == selected_project_id] if not photos_df.empty else photos_df

project_name = str(selected_project["프로젝트명"])
site_address = str(selected_project.get("현장주소", ""))
client_name = str(selected_project.get("고객명", ""))
contract_amount = int(pd.to_numeric(selected_project.get("계약금액", 0), errors="coerce") or 0)
start_date = safe_date(selected_project.get("공사시작일", datetime.now().date()))
expected_duration = int(pd.to_numeric(selected_project.get("예상공사기간", 45), errors="coerce") or 45)
expected_end_date = start_date + timedelta(days=expected_duration)


# -----------------------------
# 전체 현장 요약 계산
# -----------------------------
def calculate_project_summary(projects, tasks, issues, changes):
    rows = []
    for _, p in projects.iterrows():
        pid = int(p["프로젝트ID"])
        t = tasks[tasks["프로젝트ID"] == pid]
        i = issues[issues["프로젝트ID"] == pid] if not issues.empty else issues
        c = changes[changes["프로젝트ID"] == pid] if not changes.empty else changes
        progress = int(pd.to_numeric(t["진행률"], errors="coerce").fillna(0).mean()) if not t.empty else 0
        open_issues = len(i[i["상태"].isin(["접수", "처리중"])]) if not i.empty and "상태" in i.columns else 0
        pending_changes = len(c[c["승인상태"] == "대기"]) if not c.empty and "승인상태" in c.columns else 0
        approved_amount = int(pd.to_numeric(c.loc[c["승인상태"] == "승인", "추가금액"], errors="coerce").fillna(0).sum()) if not c.empty and "승인상태" in c.columns else 0
        rows.append(
            {
                "프로젝트ID": pid,
                "프로젝트명": p["프로젝트명"],
                "상태": p.get("상태", ""),
                "PM": p.get("PM", ""),
                "진행률": progress,
                "미처리이슈": open_issues,
                "미승인추가공사": pending_changes,
                "승인추가금액": approved_amount,
                "고객명": p.get("고객명", ""),
                "현장주소": p.get("현장주소", ""),
            }
        )
    return pd.DataFrame(rows)


summary_df = calculate_project_summary(projects_df, tasks_df, issues_df, change_orders_df)


# -----------------------------
# 상단 요약
# -----------------------------
col1, col2, col3, col4, col5 = st.columns(5)

overall_progress = int(pd.to_numeric(project_tasks["진행률"], errors="coerce").fillna(0).mean()) if not project_tasks.empty else 0
in_progress_count = len(project_tasks[project_tasks["진행상태"] == "진행중"]) if not project_tasks.empty else 0
done_count = len(project_tasks[project_tasks["진행상태"] == "완료"]) if not project_tasks.empty else 0
open_issues_count = len(project_issues[project_issues["상태"].isin(["접수", "처리중"])]) if not project_issues.empty and "상태" in project_issues.columns else 0
pending_change_count = len(project_changes[project_changes["승인상태"] == "대기"]) if not project_changes.empty and "승인상태" in project_changes.columns else 0

if not project_tasks.empty:
    end_dates = pd.to_datetime(project_tasks["완료예정"], errors="coerce")
    max_end = end_dates.max()
    remaining_days = (max_end.date() - datetime.now().date()).days if pd.notna(max_end) else 0
else:
    remaining_days = 0

col1.metric("선택 현장 진행률", f"{overall_progress}%")
col2.metric("진행중 공정", f"{in_progress_count}개")
col3.metric("완료 공정", f"{done_count}개")
col4.metric("미처리 이슈", f"{open_issues_count}건")
col5.metric("미승인 추가공사", f"{pending_change_count}건")

st.info(
    f"현재 현장: {project_name} / 주소: {site_address} / 고객: {client_name} / "
    f"계약금액: {contract_amount:,}원 / 예정 준공일: {expected_end_date} / 남은 예정일: {remaining_days}일"
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
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    if not summary_df.empty:
        fig_all = px.bar(summary_df, x="프로젝트명", y="진행률", color="상태", text="진행률", hover_data=["PM", "미처리이슈", "미승인추가공사"])
        fig_all.update_layout(height=420)
        st.plotly_chart(fig_all, use_container_width=True)

    st.subheader("현장 목록 직접 수정")
    edited_projects = st.data_editor(
        projects_df,
        use_container_width=True,
        num_rows="dynamic",
        hide_index=True,
        column_config={
            "상태": st.column_config.SelectboxColumn("상태", options=["준비중", "진행중", "보류", "완료", "취소"]),
        },
    )
    if st.button("💾 현장 목록 저장"):
        st.session_state["projects"] = edited_projects
        save_csv(edited_projects, PROJECTS_FILE)
        st.success("현장 목록이 저장되었습니다.")
        st.rerun()


# -----------------------------
# 탭1: 선택 현장 대시보드
# -----------------------------
with tab1:
    st.subheader("선택 현장 공정별 진행률")

    if project_tasks.empty:
        st.warning("이 현장의 공정 데이터가 없습니다.")
        if st.button("현재 현장 기본 공정표 생성"):
            new_tasks = create_default_tasks(selected_project_id, project_name, start_date)
            tasks_df = pd.concat([tasks_df, new_tasks], ignore_index=True)
            st.session_state["tasks"] = tasks_df
            save_csv(tasks_df, TASKS_FILE)
            st.rerun()
    else:
        fig = px.bar(
            project_tasks,
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
        check_df = project_tasks[
            (project_tasks["진행상태"].isin(["진행중", "보류"]))
            | (project_tasks["고객승인상태"].isin(["대기", "보류"]))
            | (project_tasks["지연사유"].fillna("") != "")
        ]
        if check_df.empty:
            st.success("현재 특별히 확인할 항목이 없습니다.")
        else:
            st.dataframe(check_df[["공정", "진행상태", "진행률", "담당자", "고객승인상태", "지연사유", "다음액션"]], use_container_width=True, hide_index=True)


# -----------------------------
# 탭2: Kanban
# -----------------------------
with tab2:
    st.subheader("공정 Kanban 보드")
    status_order = ["대기", "진행중", "완료", "보류"]
    cols = st.columns(4)

    for idx, status in enumerate(status_order):
        with cols[idx]:
            st.markdown(f"### {status}")
            filtered = project_tasks[project_tasks["진행상태"] == status]

            if filtered.empty:
                st.caption("해당 공정 없음")

            for _, row in filtered.iterrows():
                row_id = int(row["ID"])
                with st.expander(f"{row['공정']} / {row['담당자']} / {row['진행률']}%"):
                    new_progress = st.slider("진행률", 0, 100, int(row["진행률"]), key=f"prog_{row_id}")
                    new_status = st.selectbox("상태", status_order, index=status_order.index(row["진행상태"]), key=f"stat_{row_id}")
                    next_action = st.text_input("다음 액션", str(row.get("다음액션", "")), key=f"next_{row_id}")
                    delay_reason = st.text_area("지연사유", str(row.get("지연사유", "")), key=f"delay_{row_id}")

                    if st.button("저장", key=f"save_task_{row_id}"):
                        task_index = tasks_df[tasks_df["ID"] == row_id].index
                        if len(task_index) > 0:
                            i = task_index[0]
                            tasks_df.at[i, "진행률"] = new_progress
                            tasks_df.at[i, "진행상태"] = new_status
                            tasks_df.at[i, "다음액션"] = next_action
                            tasks_df.at[i, "지연사유"] = delay_reason
                            if new_status == "완료" and not str(tasks_df.at[i, "실제완료일"]):
                                tasks_df.at[i, "실제완료일"] = str(datetime.now().date())
                            save_csv(tasks_df, TASKS_FILE)
                            st.session_state["tasks"] = tasks_df
                            st.success("공정 정보가 저장되었습니다.")
                            st.rerun()

    st.divider()
    st.subheader("선택 현장 공정표 직접 수정")
    edited_project_tasks = st.data_editor(
        project_tasks,
        use_container_width=True,
        num_rows="dynamic",
        hide_index=True,
        column_config={
            "진행상태": st.column_config.SelectboxColumn("진행상태", options=["대기", "진행중", "완료", "보류"]),
            "자재상태": st.column_config.SelectboxColumn("자재상태", options=["미확인", "발주전", "발주완료", "입고완료", "문제발생"]),
            "고객승인필요": st.column_config.SelectboxColumn("고객승인필요", options=["예", "아니오"]),
            "고객승인상태": st.column_config.SelectboxColumn("고객승인상태", options=["해당없음", "대기", "승인", "거절", "보류"]),
        },
    )

    if st.button("💾 선택 현장 공정표 저장"):
        other_tasks = tasks_df[tasks_df["프로젝트ID"] != selected_project_id]
        combined_tasks = pd.concat([other_tasks, edited_project_tasks], ignore_index=True)
        st.session_state["tasks"] = combined_tasks
        save_csv(combined_tasks, TASKS_FILE)
        st.success("선택 현장 공정표가 저장되었습니다.")
        st.rerun()


# -----------------------------
# 탭3: 타임라인
# -----------------------------
with tab3:
    st.subheader("선택 현장 공사 타임라인")
    if project_tasks.empty:
        st.warning("표시할 공정이 없습니다.")
    else:
        timeline_df = project_tasks.copy()
        timeline_df["시작예정"] = pd.to_datetime(timeline_df["시작예정"], errors="coerce")
        timeline_df["완료예정"] = pd.to_datetime(timeline_df["완료예정"], errors="coerce")
        timeline_df = timeline_df.dropna(subset=["시작예정", "완료예정"])

        if timeline_df.empty:
            st.warning("시작예정/완료예정 날짜를 확인해 주세요.")
        else:
            fig_timeline = px.timeline(
                timeline_df,
                x_start="시작예정",
                x_end="완료예정",
                y="공정",
                color="진행상태",
                hover_data=["담당자", "진행률", "다음액션"],
            )
            fig_timeline.update_yaxes(autorange="reversed")
            fig_timeline.update_layout(height=520)
            st.plotly_chart(fig_timeline, use_container_width=True)


# -----------------------------
# 탭4: 이슈 관리
# -----------------------------
with tab4:
    st.subheader("이슈 등록")

    with st.form("issue_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        issue_process = c1.selectbox("공정", list(project_tasks["공정"].unique()) if not project_tasks.empty else ["공통"])
        issue_type = c2.selectbox("이슈유형", ["일정지연", "자재문제", "하자", "고객요청", "작업자문제", "안전", "기타"])
        issue_level = c3.selectbox("중요도", ["낮음", "보통", "높음", "긴급"])

        issue_content = st.text_area("내용")
        c4, c5, c6 = st.columns(3)
        issue_manager = c4.text_input("담당자")
        issue_status = c5.selectbox("상태", ["접수", "처리중", "완료", "보류"])
        issue_due = c6.date_input("처리기한", datetime.now().date() + timedelta(days=3))
        issue_result = st.text_area("처리내용")

        submitted = st.form_submit_button("이슈 저장")
        if submitted:
            new_row = {
                "ID": next_id(issues_df),
                "프로젝트ID": selected_project_id,
                "프로젝트명": project_name,
                "공정": issue_process,
                "이슈유형": issue_type,
                "중요도": issue_level,
                "내용": issue_content,
                "담당자": issue_manager,
                "상태": issue_status,
                "등록일": str(datetime.now().date()),
                "처리기한": str(issue_due),
                "처리내용": issue_result,
            }
            issues_df = pd.concat([issues_df, pd.DataFrame([new_row])], ignore_index=True)
            st.session_state["issues"] = issues_df
            save_csv(issues_df, ISSUES_FILE)
            st.success("이슈가 저장되었습니다.")
            st.rerun()

    st.divider()
    st.subheader("선택 현장 이슈 목록")
    edited_project_issues = st.data_editor(project_issues, use_container_width=True, num_rows="dynamic", hide_index=True)
    if st.button("💾 선택 현장 이슈 저장"):
        other_issues = issues_df[issues_df["프로젝트ID"] != selected_project_id] if not issues_df.empty else issues_df
        combined_issues = pd.concat([other_issues, edited_project_issues], ignore_index=True)
        st.session_state["issues"] = combined_issues
        save_csv(combined_issues, ISSUES_FILE)
        st.success("이슈 목록이 저장되었습니다.")
        st.rerun()


# -----------------------------
# 탭5: 추가공사 관리
# -----------------------------
with tab5:
    st.subheader("추가공사/변경 요청 등록")
    st.caption("분쟁 예방을 위해 변경내용, 추가금액, 승인방식, 승인일을 반드시 남기는 것을 권장합니다.")

    with st.form("change_order_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        co_process = c1.selectbox("공정", list(project_tasks["공정"].unique()) if not project_tasks.empty else ["공통"], key="co_process")
        requester = c2.selectbox("요청자", ["고객", "시공자", "협력업체", "기타"])
        amount = c3.number_input("추가금액", min_value=0, value=0, step=100000, format="%d")

        change_content = st.text_area("변경내용")
        c4, c5, c6 = st.columns(3)
        approval_status = c4.selectbox("승인상태", ["대기", "승인", "거절", "보류"])
        approval_method = c5.selectbox("승인방식", ["미정", "구두", "카카오톡", "문자", "이메일", "서명", "계약서/견적서"])
        request_date = c6.date_input("요청일", datetime.now().date())
        approval_date = st.date_input("승인일", datetime.now().date())
        co_memo = st.text_area("메모")

        co_submitted = st.form_submit_button("추가공사 저장")
        if co_submitted:
            new_row = {
                "ID": next_id(change_orders_df),
                "프로젝트ID": selected_project_id,
                "프로젝트명": project_name,
                "공정": co_process,
                "요청자": requester,
                "변경내용": change_content,
                "추가금액": amount,
                "승인상태": approval_status,
                "승인방식": approval_method,
                "요청일": str(request_date),
                "승인일": str(approval_date) if approval_status == "승인" else "",
                "메모": co_memo,
            }
            change_orders_df = pd.concat([change_orders_df, pd.DataFrame([new_row])], ignore_index=True)
            st.session_state["change_orders"] = change_orders_df
            save_csv(change_orders_df, CHANGE_ORDERS_FILE)
            st.success("추가공사 기록이 저장되었습니다.")
            st.rerun()

    st.divider()
    st.subheader("선택 현장 추가공사 목록")
    if not project_changes.empty:
        total_approved = pd.to_numeric(project_changes.loc[project_changes["승인상태"] == "승인", "추가금액"], errors="coerce").fillna(0).sum()
        st.metric("승인된 추가공사 합계", f"{int(total_approved):,}원")

    edited_project_changes = st.data_editor(project_changes, use_container_width=True, num_rows="dynamic", hide_index=True)
    if st.button("💾 선택 현장 추가공사 저장"):
        other_changes = change_orders_df[change_orders_df["프로젝트ID"] != selected_project_id] if not change_orders_df.empty else change_orders_df
        combined_changes = pd.concat([other_changes, edited_project_changes], ignore_index=True)
        st.session_state["change_orders"] = combined_changes
        save_csv(combined_changes, CHANGE_ORDERS_FILE)
        st.success("추가공사 목록이 저장되었습니다.")
        st.rerun()


# -----------------------------
# 탭6: 사진/증거 관리
# -----------------------------
with tab6:
    st.subheader("현장 사진/증거 업로드")
    st.caption("공정별 작업 전·중·후 사진을 남기면 하자, 추가공사, 일정 지연 분쟁 대응에 도움이 됩니다.")

    c1, c2, c3 = st.columns(3)
    photo_process = c1.selectbox("공정", list(project_tasks["공정"].unique()) if not project_tasks.empty else ["공통"], key="photo_process")
    photo_type = c2.selectbox("사진구분", ["작업 전", "작업 중", "작업 후", "하자", "자재", "고객요청", "기타"])
    photo_description = c3.text_input("사진 설명")

    uploaded_files = st.file_uploader("사진 업로드 JPG/PNG/JPEG", type=["jpg", "png", "jpeg"], accept_multiple_files=True)

    if uploaded_files and st.button("📸 사진 저장"):
        new_rows = []
        for uploaded_file in uploaded_files:
            ext = os.path.splitext(uploaded_file.name)[1].lower()
            safe_project = project_name.replace("/", "_").replace("\\", "_").replace(" ", "_")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            saved_filename = f"P{selected_project_id}_{safe_project}_{photo_process}_{timestamp}{ext}"
            saved_path = os.path.join(PHOTO_DIR, saved_filename)

            image = Image.open(uploaded_file)
            image.save(saved_path)

            new_rows.append(
                {
                    "ID": next_id(photos_df) + len(new_rows),
                    "프로젝트ID": selected_project_id,
                    "프로젝트명": project_name,
                    "공정": photo_process,
                    "사진구분": photo_type,
                    "파일명": uploaded_file.name,
                    "저장경로": saved_path,
                    "설명": photo_description,
                    "업로드일시": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }
            )

        photos_df = pd.concat([photos_df, pd.DataFrame(new_rows)], ignore_index=True)
        st.session_state["photos"] = photos_df
        save_csv(photos_df, PHOTOS_FILE)
        st.success("사진이 저장되었습니다.")
        st.rerun()

    st.divider()
    st.subheader("선택 현장 사진 기록 목록")
    st.dataframe(project_photos, use_container_width=True, hide_index=True)

    st.subheader("사진 미리보기")
    if project_photos.empty:
        st.info("아직 저장된 사진이 없습니다.")
    else:
        recent_photos = project_photos.tail(12).iloc[::-1]
        photo_cols = st.columns(3)
        for idx, (_, row) in enumerate(recent_photos.iterrows()):
            with photo_cols[idx % 3]:
                path = str(row.get("저장경로", ""))
                if os.path.exists(path):
                    st.image(path, caption=f"{row.get('공정', '')} / {row.get('사진구분', '')} / {row.get('설명', '')}", use_container_width=True)
                else:
                    st.warning(f"파일을 찾을 수 없습니다: {path}")


# -----------------------------
# 탭7: 데이터 관리
# -----------------------------
with tab7:
    st.subheader("전체 데이터 백업 및 다운로드")

    c1, c2 = st.columns(2)

    with c1:
        if st.button("💾 모든 데이터 저장"):
            save_csv(st.session_state["projects"], PROJECTS_FILE)
            save_csv(st.session_state["tasks"], TASKS_FILE)
            save_csv(st.session_state["issues"], ISSUES_FILE)
            save_csv(st.session_state["change_orders"], CHANGE_ORDERS_FILE)
            save_csv(st.session_state["photos"], PHOTOS_FILE)
            st.success("모든 데이터가 저장되었습니다.")

    with c2:
        excel_data = make_excel_file(
            st.session_state["projects"],
            st.session_state["tasks"],
            st.session_state["issues"],
            st.session_state["change_orders"],
            st.session_state["photos"],
        )
        st.download_button(
            label="📥 전체 현황 엑셀 다운로드",
            data=excel_data,
            file_name=f"Interior_Flow_v2_현황_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.divider()
    st.subheader("주의사항")
    st.warning(
        "현재 Streamlit Community Cloud의 파일 저장은 영구 DB가 아닙니다. 앱이 재시작되거나 배포가 갱신되면 CSV/사진 파일이 초기화될 수 있습니다. "
        "실사용 단계에서는 Supabase 같은 외부 DB와 Storage 연동이 필요합니다."
    )

st.caption("Interior Flow v2 / 다현장 관리 MVP / 다음 단계: Supabase DB, 로그인, 모바일 입력 최적화, AI 일일 리포트")
