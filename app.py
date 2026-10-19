import csv
import io
import re
import uuid
from datetime import date, datetime
from typing import Any

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError, SpreadsheetNotFound, WorksheetNotFound

# =========================
# Google Sheets 설정: 여기에 실제 스프레드시트 URL을 입력하세요.
# =========================
SHEET_URL = "https://docs.google.com/spreadsheets/d/여기에_시트_ID/edit"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SHEETS = {
    "기자재목록": [
        "관리번호", "기자재명", "분류", "모델명", "제조사", "보유수량",
        "대여가능수량", "보관위치", "구입일", "구입가격", "상태", "담당자",
        "비고", "등록일", "최소재고수량", "활성여부",
    ],
    "대여반납": [
        "대여번호", "관리번호", "기자재명", "대여자", "소속", "대여수량",
        "대여일", "반납예정일", "실제반납일", "반납수량", "상태", "담당자",
        "비고", "누적반납수량", "파손수량", "분실수량", "반납기록",
    ],
    "소모품이력": [
        "기록번호", "품목관리번호", "품목명", "처리유형", "수량", "처리일",
        "처리자", "비고", "현재재고", "최소재고수량",
    ],
    "변경이력": [
        "이력번호", "처리시각", "작업유형", "대상관리번호", "대상명", "변경내용", "처리자",
    ],
    "설정": ["설정항목", "설정값"],
}

CATEGORIES = [
    "전동 공구", "수공구", "측정 장비", "목공 장비", "금속 가공 장비",
    "전기·전자 기자재", "안전 장비", "소모품", "기타",
]
ITEM_STATES = ["사용 가능", "수리 중", "폐기"]
LOAN_STATES = ["대여 중", "부분 반납", "반납 완료", "연체", "분실 처리", "파손 처리"]


st.set_page_config(page_title="기술실 기자재 관리", page_icon="🧰", layout="wide")
st.title("🧰 기술실 기자재 관리")
st.caption("기자재 · 대여/반납 · 소모품 · 변경 이력")

def now_text() -> str:
    return datetime.now().astimezone().strftime("%Y-%m-%d %H:%M:%S")

def new_id(prefix: str) -> str:
    return f"{prefix}-{datetime.now().strftime('%y%m%d')}-{uuid.uuid4().hex[:6].upper()}"

def safe_int(value: Any, default: int = 0) -> int:
    try:
        if value is None or str(value).strip() == "":
            return default
        return int(float(str(value).replace(",", "").strip()))
    except (TypeError, ValueError):
        return default

def safe_float(value: Any, default: float = 0.0) -> float:
    try:
        if value is None or str(value).strip() == "":
            return default
        return float(str(value).replace(",", "").strip())
    except (TypeError, ValueError):
        return default

def normalize_date(value: Any) -> str:
    if value is None or str(value).strip() == "":
        return ""
    if isinstance(value, (date, datetime)):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    try:
        return pd.to_datetime(text).strftime("%Y-%m-%d")
    except Exception:
        return text

def get_client():
    if not SHEET_URL or "여기에_시트_ID" in SHEET_URL:
        raise ValueError("app.py 상단의 SHEET_URL에 실제 Google Sheets 주소를 입력해 주세요.")
    if "gcp_service_account" not in st.secrets:
        raise ValueError("Streamlit Secrets에 [gcp_service_account] 인증 설정이 없습니다.")
    info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)

@st.cache_resource(show_spinner=False)
def get_spreadsheet():
    client = get_client()
    return client.open_by_url(SHEET_URL)

def get_worksheet(spreadsheet, name: str):
    try:
        ws = spreadsheet.worksheet(name)
    except WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=name, rows=1000, cols=max(20, len(SHEETS[name])))
    expected = SHEETS[name]
    existing = ws.row_values(1)
    if not existing:
        ws.update(range_name="A1", values=[expected])
    else:
        # 기존 헤더와 데이터는 보존하고, 빠진 컬럼은 오른쪽에 추가한다.
        missing = [col for col in expected if col not in existing]
        if missing:
            ws.update_cell(1, len(existing) + 1, existing[-1] if existing else expected[0]) if False else None
            ws.update(range_name=f"{gspread.utils.rowcol_to_a1(1, len(existing)+1)}", values=[missing], value_input_option="USER_ENTERED")
    return ws

def read_sheet(name: str) -> pd.DataFrame:
    ss = get_spreadsheet()
    ws = get_worksheet(ss, name)
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    for col in SHEETS[name]:
        if col not in df.columns:
            df[col] = ""
    return df[SHEETS[name]]

def append_record(name: str, record: dict):
    ss = get_spreadsheet()
    ws = get_worksheet(ss, name)
    headers = ws.row_values(1)
    row = [record.get(h, "") for h in headers]
    ws.append_row(row, value_input_option="USER_ENTERED")
    st.cache_data.clear()

def update_record(name: str, key_col: str, key_value: str, changes: dict):
    ss = get_spreadsheet()
    ws = get_worksheet(ss, name)
    headers = ws.row_values(1)
    if key_col not in headers:
        raise ValueError(f"{name} 시트에 '{key_col}' 컬럼이 없습니다.")
    key_idx = headers.index(key_col) + 1
    cell = ws.find(str(key_value), in_column=key_idx)
    if not cell:
        raise ValueError("수정할 기록을 찾을 수 없습니다. 새로고침 후 다시 시도해 주세요.")
    for col, value in changes.items():
        if col in headers:
            ws.update_cell(cell.row, headers.index(col) + 1, value)
    st.cache_data.clear()

def log_change(action: str, target_id: str, target_name: str, detail: str, actor: str):
    append_record("변경이력", {
        "이력번호": new_id("LOG"), "처리시각": now_text(), "작업유형": action,
        "대상관리번호": target_id, "대상명": target_name, "변경내용": detail, "처리자": actor or "미입력",
    })

def cached_read(name: str) -> pd.DataFrame:
    # 읽기 캐시를 사용하지 않고 매번 최신 데이터를 읽어 재고 변경 후 오래된 값이 남지 않게 한다.
    return read_sheet(name)

def initialize():
    ss = get_spreadsheet()
    for name in SHEETS:
        get_worksheet(ss, name)

def refresh():
    st.cache_data.clear()
    st.rerun()

def require_actor(actor: str) -> bool:
    if not actor.strip():
        st.error("처리 담당자를 입력해 주세요.")
        return False
    return True

def export_csv(df: pd.DataFrame, filename: str, label: str):
    data = df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(label, data=data, file_name=filename, mime="text/csv", use_container_width=True)

def item_row(items: pd.DataFrame, item_id: str):
    rows = items[items["관리번호"].astype(str) == str(item_id)]
    return None if rows.empty else rows.iloc[0].to_dict()

def update_item_quantities(item_id: str, available: int, state: str | None = None):
    changes = {"대여가능수량": available}
    if state is not None:
        changes["상태"] = state
    update_record("기자재목록", "관리번호", item_id, changes)

def get_open_loans(loans: pd.DataFrame) -> pd.DataFrame:
    if loans.empty:
        return loans
    return loans[loans["상태"].astype(str).isin(["대여 중", "부분 반납", "연체"])].copy()

try:
    initialize()
    items = cached_read("기자재목록")
    loans = cached_read("대여반납")
    consumables = cached_read("소모품이력")
    history = cached_read("변경이력")
except ValueError as e:
    st.error(str(e))
    st.info("1) app.py의 SHEET_URL을 수정하고 2) Streamlit Secrets에 서비스 계정 정보를 설정하세요.")
    st.stop()
except Exception as e:
    message = str(e).lower()
    if "permission" in message or "403" in message:
        st.error("Google Sheets 접근 권한이 없습니다. 서비스 계정 이메일을 해당 스프레드시트의 편집자로 추가했는지 확인하세요.")
    elif "not found" in message or "404" in message:
        st.error("스프레드시트를 찾을 수 없습니다. SHEET_URL과 공유 권한을 확인하세요.")
    else:
        st.error("Google Sheets 연결 또는 초기화에 실패했습니다. URL, Secrets, API 활성화 여부를 확인하세요.")
    with st.expander("기술 정보"):
        st.code(type(e).__name__)
    st.stop()

with st.sidebar:
    st.header("메뉴")
    page = st.radio(
        "화면 선택",
        ["대시보드", "기자재 목록", "기자재 등록", "대여 및 반납", "소모품 관리", "변경 이력", "설정 및 도움말"],
        label_visibility="collapsed",
    )
    st.divider()
    actor_default = st.text_input("현재 처리 담당자", placeholder="예: 홍길동")
    st.caption("현재 버전은 별도 로그인/권한 인증을 제공하지 않습니다.")
    if st.button("데이터 새로고침", use_container_width=True):
        refresh()

# Load fresh data for every page render.
items = cached_read("기자재목록")
loans = cached_read("대여반납")
consumables = cached_read("소모품이력")
history = cached_read("변경이력")

# Normalize numbers and dates
for col in ["보유수량", "대여가능수량", "최소재고수량", "구입가격"]:
    if col in items:
        items[col] = items[col].apply(safe_int)
if not loans.empty:
    for col in ["대여수량", "반납수량", "누적반납수량", "파손수량", "분실수량"]:
        loans[col] = loans[col].apply(safe_int)
    loans["반납예정일"] = loans["반납예정일"].apply(normalize_date)
    loans["대여일"] = loans["대여일"].apply(normalize_date)
    loans["실제반납일"] = loans["실제반납일"].apply(normalize_date)
if not consumables.empty:
    for col in ["수량", "현재재고", "최소재고수량"]:
        consumables[col] = consumables[col].apply(safe_int)

if page == "대시보드":
    st.subheader("기술실 현황")
    active_items = items[items["활성여부"].astype(str).str.lower().isin(["true", "1", "yes", "y", ""])].copy()
    total_types = len(active_items[active_items["상태"] != "폐기"])
    total_stock = int(active_items["보유수량"].sum()) if not active_items.empty else 0
    available = int(active_items.loc[active_items["상태"] == "사용 가능", "대여가능수량"].sum()) if not active_items.empty else 0
    repair = int(active_items.loc[active_items["상태"] == "수리 중", "보유수량"].sum()) if not active_items.empty else 0
    disposed = int(items.loc[items["상태"] == "폐기", "보유수량"].sum()) if not items.empty else 0
    open_loans = get_open_loans(loans)
    borrowed = int((open_loans["대여수량"] - open_loans["누적반납수량"]).clip(lower=0).sum()) if not open_loans.empty else 0
    today = date.today().strftime("%Y-%m-%d")
    overdue = int(((open_loans["반납예정일"] < today) & (open_loans["반납예정일"] != "")).sum()) if not open_loans.empty else 0
    low_stock = 0
    if not consumables.empty:
        latest = consumables.sort_values("처리일").groupby("품목관리번호", as_index=False).tail(1)
        low_stock = int((latest["현재재고"] <= latest["최소재고수량"]).sum())
    metrics = [
        ("등록 기자재 종류", total_types), ("전체 보유 수량", total_stock),
        ("대여 중 수량", borrowed), ("사용 가능 수량", available),
        ("수리 중 수량", repair), ("폐기 수량", disposed),
        ("연체 대여 건수", overdue), ("재고 부족 소모품", low_stock),
    ]
    for i in range(0, len(metrics), 4):
        cols = st.columns(4)
        for col, (label, value) in zip(cols, metrics[i:i+4]):
            col.metric(label, f"{value:,}")
    left, right = st.columns(2)
    with left:
        st.markdown("#### 기자재 상태별 수량")
        if items.empty:
            st.info("아직 등록된 기자재가 없습니다.")
        else:
            state_data = items.groupby("상태", dropna=False)["보유수량"].sum().reset_index()
            st.bar_chart(state_data.set_index("상태"))
    with right:
        st.markdown("#### 분류별 보유 수량")
        if items.empty:
            st.info("기자재를 등록하면 분류별 차트가 표시됩니다.")
        else:
            category_data = items.groupby("분류", dropna=False)["보유수량"].sum().reset_index()
            st.bar_chart(category_data.set_index("분류"))
    st.markdown("#### 반납 예정일이 지난 대여")
    if overdue:
        show = open_loans[(open_loans["반납예정일"] < today) & (open_loans["반납예정일"] != "")]
        st.dataframe(show, use_container_width=True, hide_index=True)
    else:
        st.success("현재 확인된 연체 대여가 없습니다.")

elif page == "기자재 목록":
    st.subheader("기자재 목록 및 검색")
    c1, c2, c3 = st.columns([2, 1, 1])
    query = c1.text_input("검색", placeholder="기자재명 또는 관리번호")
    cats = ["전체"] + sorted([str(x) for x in items["분류"].dropna().unique() if str(x).strip()])
    states = ["전체"] + ITEM_STATES
    cat_filter = c2.selectbox("분류", cats)
    state_filter = c3.selectbox("상태", states)
    c4, c5, c6 = st.columns(3)
    locations = ["전체"] + sorted([str(x) for x in items["보관위치"].dropna().unique() if str(x).strip()])
    loc_filter = c4.selectbox("보관 위치", locations)
    low_only = c5.checkbox("재고 부족만 보기")
    borrowed_only = c6.checkbox("대여 중인 항목만 보기")
    filtered = items.copy()
    if query:
        filtered = filtered[
            filtered["기자재명"].astype(str).str.contains(query, case=False, na=False)
            | filtered["관리번호"].astype(str).str.contains(query, case=False, na=False)
        ]
    if cat_filter != "전체": filtered = filtered[filtered["분류"] == cat_filter]
    if state_filter != "전체": filtered = filtered[filtered["상태"] == state_filter]
    if loc_filter != "전체": filtered = filtered[filtered["보관위치"] == loc_filter]
    if low_only: filtered = filtered[filtered["대여가능수량"] <= filtered["최소재고수량"]]
    if borrowed_only:
        borrowed_ids = set(get_open_loans(loans)["관리번호"].astype(str)) if not loans.empty else set()
        filtered = filtered[filtered["관리번호"].astype(str).isin(borrowed_ids)]
    st.caption(f"검색 결과 {len(filtered)}건")
    st.dataframe(filtered, use_container_width=True, hide_index=True)
    export_csv(filtered, "기자재목록.csv", "검색 결과 CSV 다운로드")
    st.markdown("#### 기자재 정보 수정 / 폐기")
    if not items.empty:
        selected = st.selectbox("수정할 기자재", items["관리번호"].astype(str) + " · " + items["기자재명"].astype(str))
        selected_id = selected.split(" · ", 1)[0]
        row = item_row(items, selected_id)
        with st.form("edit_item_form"):
            ec1, ec2, ec3 = st.columns(3)
            name = ec1.text_input("기자재명 *", value=str(row.get("기자재명", "")))
            category = ec2.selectbox("분류", CATEGORIES, index=CATEGORIES.index(row["분류"]) if row.get("분류") in CATEGORIES else len(CATEGORIES)-1)
            model = ec3.text_input("모델명", value=str(row.get("모델명", "")))
            ec4, ec5, ec6 = st.columns(3)
            maker = ec4.text_input("제조사", value=str(row.get("제조사", "")))
            location = ec5.text_input("보관 위치", value=str(row.get("보관위치", "")))
            state = ec6.selectbox("관리 상태", ITEM_STATES, index=ITEM_STATES.index(row["상태"]) if row.get("상태") in ITEM_STATES else 0)
            ec7, ec8, ec9 = st.columns(3)
            stock = ec7.number_input("보유 수량", min_value=0, value=max(0, safe_int(row.get("보유수량"))), step=1)
            available_value = ec8.number_input("대여 가능 수량", min_value=0, max_value=int(stock), value=min(int(stock), max(0, safe_int(row.get("대여가능수량")))), step=1)
            minimum = ec9.number_input("최소 재고 수량", min_value=0, value=max(0, safe_int(row.get("최소재고수량"))), step=1)
            ec10, ec11 = st.columns(2)
            purchase_date = ec10.text_input("구입일 (YYYY-MM-DD)", value=str(row.get("구입일", "")))
            price = ec11.number_input("구입 가격(원)", min_value=0, value=max(0, safe_int(row.get("구입가격"))), step=1000)
            manager = st.text_input("담당자", value=str(row.get("담당자", "")))
            note = st.text_area("비고", value=str(row.get("비고", "")))
            submit_edit = st.form_submit_button("변경사항 저장", use_container_width=True)
        if submit_edit:
            if not name.strip():
                st.error("기자재명은 필수입니다.")
            elif safe_int(row.get("보유수량")) - safe_int(row.get("대여가능수량")) > int(stock) - int(available_value):
                st.error("대여 중 수량과 불일치할 수 있습니다. 보유/대여 가능 수량을 다시 확인하세요.")
            elif state == "폐기" and safe_int(row.get("보유수량")) != safe_int(row.get("대여가능수량")):
                st.error("대여 중인 기자재가 있으면 폐기 상태로 변경할 수 없습니다.")
            else:
                changes = {
                    "기자재명": name.strip(), "분류": category, "모델명": model.strip(),
                    "제조사": maker.strip(), "보유수량": int(stock), "대여가능수량": int(available_value),
                    "보관위치": location.strip(), "구입일": normalize_date(purchase_date),
                    "구입가격": int(price), "상태": state, "담당자": manager.strip(),
                    "비고": note.strip(), "최소재고수량": int(minimum),
                }
                update_record("기자재목록", "관리번호", selected_id, changes)
                log_change("기자재 수정", selected_id, name, "기본 정보 또는 수량 수정", actor_default)
                st.success("기자재 정보가 저장되었습니다.")
                refresh()
    else:
        st.info("수정할 기자재가 없습니다. 먼저 기자재를 등록하세요.")

elif page == "기자재 등록":
    st.subheader("새 기자재 등록")
    with st.form("new_item_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        custom_id = c1.text_input("관리번호 (비우면 자동 생성)", placeholder="예: TOOL-001")
        name = c2.text_input("기자재명 *")
        category = c3.selectbox("분류 *", CATEGORIES)
        c4, c5, c6 = st.columns(3)
        model = c4.text_input("모델명 또는 규격")
        maker = c5.text_input("제조사")
        location = c6.text_input("보관 위치")
        c7, c8, c9 = st.columns(3)
        stock = c7.number_input("보유 수량 *", min_value=0, value=1, step=1)
        minimum = c8.number_input("최소 재고 수량", min_value=0, value=0, step=1)
        price = c9.number_input("구입 가격(원)", min_value=0, value=0, step=1000)
        c10, c11 = st.columns(2)
        purchase_date = c10.date_input("구입일", value=None)
        state = c11.selectbox("관리 상태", ITEM_STATES, index=0)
        manager = st.text_input("관리 담당자")
        note = st.text_area("비고")
        submit = st.form_submit_button("기자재 등록", use_container_width=True)
    if submit:
        item_id = custom_id.strip() or new_id("EQ")
        if not name.strip():
            st.error("기자재명은 필수입니다.")
        elif stock < 0 or minimum < 0 or price < 0:
            st.error("수량과 가격은 음수일 수 없습니다.")
        elif (items["관리번호"].astype(str) == item_id).any():
            st.error("이미 사용 중인 관리번호입니다. 다른 번호를 입력하세요.")
        elif not require_actor(actor_default or manager):
            pass
        else:
            record = {
                "관리번호": item_id, "기자재명": name.strip(), "분류": category,
                "모델명": model.strip(), "제조사": maker.strip(), "보유수량": int(stock),
                "대여가능수량": int(stock) if state == "사용 가능" else 0,
                "보관위치": location.strip(), "구입일": purchase_date.strftime("%Y-%m-%d") if purchase_date else "",
                "구입가격": int(price), "상태": state, "담당자": (actor_default or manager).strip(),
                "비고": note.strip(), "등록일": date.today().strftime("%Y-%m-%d"),
                "최소재고수량": int(minimum), "활성여부": "TRUE",
            }
            append_record("기자재목록", record)
            log_change("기자재 등록", item_id, name.strip(), "신규 기자재 등록", actor_default or manager)
            st.success(f"기자재 '{name}'이(가) 등록되었습니다. 관리번호: {item_id}")
            refresh()

elif page == "대여 및 반납":
    st.subheader("대여 처리")
    available_items = items[(items["상태"] == "사용 가능") & (items["대여가능수량"] > 0)]
    if available_items.empty:
        st.info("현재 대여 가능한 기자재가 없습니다.")
    else:
        with st.form("loan_form", clear_on_submit=True):
            choices = available_items["관리번호"].astype(str) + " · " + available_items["기자재명"].astype(str) + " (가능 " + available_items["대여가능수량"].astype(str) + ")"
            selected = st.selectbox("기자재 선택", choices)
            selected_id = selected.split(" · ", 1)[0]
            selected_row = item_row(items, selected_id) or {}
            max_qty = max(1, safe_int(selected_row.get("대여가능수량")))
            c1, c2, c3 = st.columns(3)
            borrower = c1.text_input("대여자 이름 *")
            affiliation = c2.text_input("학급 또는 소속")
            qty = c3.number_input("대여 수량", min_value=1, max_value=max_qty, value=1, step=1)
            c4, c5 = st.columns(2)
            loan_date = c4.date_input("대여일", value=date.today())
            due_date = c5.date_input("반납 예정일", value=date.today())
            manager = st.text_input("대여 담당자", value=actor_default)
            note = st.text_area("비고")
            submit_loan = st.form_submit_button("대여 승인", use_container_width=True)
        if submit_loan:
            fresh_items = cached_read("기자재목록")
            fresh = item_row(fresh_items, selected_id)
            if not borrower.strip():
                st.error("대여자 이름을 입력해 주세요.")
            elif not fresh or safe_int(fresh.get("대여가능수량")) < int(qty) or fresh.get("상태") != "사용 가능":
                st.error("현재 대여 가능 수량이 부족하거나 상태가 변경되었습니다. 새로고침 후 다시 시도하세요.")
            elif due_date < loan_date:
                st.error("반납 예정일은 대여일보다 빠를 수 없습니다.")
            elif not require_actor(manager):
                pass
            else:
                loan_id = new_id("LOAN")
                append_record("대여반납", {
                    "대여번호": loan_id, "관리번호": selected_id, "기자재명": fresh.get("기자재명", ""),
                    "대여자": borrower.strip(), "소속": affiliation.strip(), "대여수량": int(qty),
                    "대여일": loan_date.strftime("%Y-%m-%d"), "반납예정일": due_date.strftime("%Y-%m-%d"),
                    "실제반납일": "", "반납수량": 0, "상태": "대여 중", "담당자": manager.strip(),
                    "비고": note.strip(), "누적반납수량": 0, "파손수량": 0, "분실수량": 0, "반납기록": "",
                })
                update_item_quantities(selected_id, safe_int(fresh.get("대여가능수량")) - int(qty))
                log_change("대여", selected_id, str(fresh.get("기자재명")), f"{borrower}에게 {int(qty)}개 대여 (대여번호 {loan_id})", manager)
                st.success(f"대여 처리 완료: {loan_id}")
                refresh()
    st.divider()
    st.subheader("반납 처리")
    open_loans = get_open_loans(loans)
    if open_loans.empty:
        st.info("반납 처리할 대여 기록이 없습니다.")
    else:
        loan_labels = open_loans.apply(lambda r: f"{r['대여번호']} · {r['기자재명']} · {r['대여자']} (미반납 {max(0, safe_int(r['대여수량'])-safe_int(r['누적반납수량']))})", axis=1)
        selected_loan_label = st.selectbox("반납할 대여 기록", loan_labels, key="return_loan")
        loan_id = selected_loan_label.split(" · ", 1)[0]
        loan_row = open_loans[open_loans["대여번호"].astype(str) == loan_id].iloc[0].to_dict()
        remaining = max(0, safe_int(loan_row.get("대여수량")) - safe_int(loan_row.get("누적반납수량")))
        with st.form("return_form"):
            c1, c2, c3 = st.columns(3)
            return_qty = c1.number_input("이번 반납 수량", min_value=1, max_value=max(1, remaining), value=max(1, remaining), step=1)
            broken = c2.number_input("파손 수량", min_value=0, max_value=max(1, remaining), value=0, step=1)
            lost = c3.number_input("분실 수량", min_value=0, max_value=max(1, remaining), value=0, step=1)
            actual_date = st.date_input("실제 반납일", value=date.today())
            return_note = st.text_area("반납 비고")
            return_actor = st.text_input("반납 담당자", value=actor_default)
            submit_return = st.form_submit_button("반납 처리", use_container_width=True)
        if submit_return:
            fresh_loans = cached_read("대여반납")
            matches = fresh_loans[fresh_loans["대여번호"].astype(str) == loan_id]
            if matches.empty:
                st.error("대여 기록을 찾을 수 없습니다.")
            else:
                fresh_loan = matches.iloc[0].to_dict()
                fresh_remaining = max(0, safe_int(fresh_loan.get("대여수량")) - safe_int(fresh_loan.get("누적반납수량")))
                if fresh_remaining <= 0 or int(return_qty) > fresh_remaining:
                    st.error("이미 반납되었거나 반납 가능 수량을 초과했습니다. 새로고침 후 다시 시도하세요.")
                elif int(broken) + int(lost) > int(return_qty):
                    st.error("파손 수량과 분실 수량의 합은 이번 반납 수량을 초과할 수 없습니다.")
                elif not require_actor(return_actor):
                    pass
                else:
                    item = item_row(cached_read("기자재목록"), str(fresh_loan.get("관리번호")))
                    if not item:
                        st.error("연결된 기자재 정보를 찾을 수 없습니다.")
                    else:
                        cumulative = safe_int(fresh_loan.get("누적반납수량")) + int(return_qty)
                        total_broken = safe_int(fresh_loan.get("파손수량")) + int(broken)
                        total_lost = safe_int(fresh_loan.get("분실수량")) + int(lost)
                        new_state = "반납 완료" if cumulative >= safe_int(fresh_loan.get("대여수량")) else "부분 반납"
                        note_history = str(fresh_loan.get("반납기록", ""))
                        event = f"{actual_date:%Y-%m-%d}: 반납 {int(return_qty)}, 파손 {int(broken)}, 분실 {int(lost)}"
                        if return_note.strip(): event += f", 비고: {return_note.strip()}"
                        note_history = (note_history + " | " + event).strip(" |")
                        update_record("대여반납", "대여번호", loan_id, {
                            "누적반납수량": cumulative, "반납수량": cumulative,
                            "실제반납일": actual_date.strftime("%Y-%m-%d") if new_state == "반납 완료" else "",
                            "상태": new_state, "파손수량": total_broken, "분실수량": total_lost,
                            "반납기록": note_history,
                        })
                        good_return = int(return_qty) - int(broken) - int(lost)
                        new_available = safe_int(item.get("대여가능수량")) + good_return
                        update_item_quantities(str(fresh_loan.get("관리번호")), new_available)
                        log_change("반납", str(fresh_loan.get("관리번호")), str(fresh_loan.get("기자재명")), f"대여번호 {loan_id}: 반납 {int(return_qty)}, 정상 복귀 {good_return}, 파손 {int(broken)}, 분실 {int(lost)}", return_actor)
                        if broken or lost:
                            log_change("파손/분실", str(fresh_loan.get("관리번호")), str(fresh_loan.get("기자재명")), f"파손 {int(broken)}, 분실 {int(lost)}", return_actor)
                        st.success("반납 처리가 완료되었습니다.")
                        refresh()
    st.divider()
    st.subheader("대여·반납 이력")
    st.dataframe(loans.sort_values("대여일", ascending=False) if not loans.empty else loans, use_container_width=True, hide_index=True)
    export_csv(loans, "대여반납내역.csv", "대여·반납 CSV 다운로드")

elif page == "소모품 관리":
    st.subheader("소모품 입고 및 사용")
    with st.form("consumable_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        item_id = c1.text_input("품목 관리번호 *", placeholder="예: CONS-001")
        item_name = c2.text_input("품목명 *", placeholder="예: 사포")
        c3, c4, c5 = st.columns(3)
        movement = c3.selectbox("입출고 구분", ["입고", "사용"])
        qty = c4.number_input("수량", min_value=1, value=1, step=1)
        minimum = c5.number_input("최소 재고 수량", min_value=0, value=0, step=1)
        processed_date = st.date_input("처리일", value=date.today())
        processor = st.text_input("처리자", value=actor_default)
        note = st.text_area("비고")
        submit_cons = st.form_submit_button("입출고 기록 저장", use_container_width=True)
    if submit_cons:
        if not item_id.strip() or not item_name.strip():
            st.error("품목 관리번호와 품목명은 필수입니다.")
        elif not require_actor(processor):
            pass
        else:
            existing = consumables[consumables["품목관리번호"].astype(str) == item_id.strip()]
            current = safe_int(existing.sort_values("처리일").iloc[-1]["현재재고"]) if not existing.empty else 0
            new_stock = current + int(qty) if movement == "입고" else current - int(qty)
            if new_stock < 0:
                st.error(f"재고가 부족합니다. 현재 재고: {current}")
            else:
                append_record("소모품이력", {
                    "기록번호": new_id("CONS"), "품목관리번호": item_id.strip(), "품목명": item_name.strip(),
                    "처리유형": movement, "수량": int(qty), "처리일": processed_date.strftime("%Y-%m-%d"),
                    "처리자": processor.strip(), "비고": note.strip(), "현재재고": new_stock,
                    "최소재고수량": int(minimum),
                })
                log_change(f"소모품 {movement}", item_id.strip(), item_name.strip(), f"{movement} {int(qty)}개, 재고 {new_stock}", processor)
                st.success(f"{movement} 기록 저장 완료. 현재 재고: {new_stock}")
                refresh()
    st.divider()
    st.subheader("현재 소모품 재고")
    if consumables.empty:
        st.info("아직 소모품 입출고 기록이 없습니다.")
    else:
        latest = consumables.sort_values("처리일").groupby("품목관리번호", as_index=False).tail(1).sort_values("품목명")
        latest["재고 상태"] = latest.apply(lambda r: "재고 부족" if safe_int(r["현재재고"]) <= safe_int(r["최소재고수량"]) else "정상", axis=1)
        st.dataframe(latest[["품목관리번호", "품목명", "현재재고", "최소재고수량", "재고 상태", "처리일"]], use_container_width=True, hide_index=True)
        st.markdown("#### 소모품 처리 이력")
        st.dataframe(consumables.sort_values("처리일", ascending=False), use_container_width=True, hide_index=True)
        export_csv(consumables, "소모품이력.csv", "소모품 이력 CSV 다운로드")

elif page == "변경 이력":
    st.subheader("변경 이력 조회")
    c1, c2, c3 = st.columns(3)
    q = c1.text_input("관리번호·대상명·처리자 검색")
    action_options = ["전체"] + sorted([str(x) for x in history["작업유형"].dropna().unique() if str(x).strip()])
    action_filter = c2.selectbox("작업 유형", action_options)
    date_filter = c3.date_input("시작일 필터", value=None)
    view = history.copy()
    if q:
        mask = pd.Series(False, index=view.index)
        for col in ["대상관리번호", "대상명", "처리자", "변경내용"]:
            mask |= view[col].astype(str).str.contains(q, case=False, na=False)
        view = view[mask]
    if action_filter != "전체": view = view[view["작업유형"] == action_filter]
    if date_filter:
        view = view[view["처리시각"].astype(str) >= date_filter.strftime("%Y-%m-%d")]
    if not view.empty:
        view = view.sort_values("처리시각", ascending=False)
    st.dataframe(view, use_container_width=True, hide_index=True)
    export_csv(view, "변경이력.csv", "변경 이력 CSV 다운로드")
    st.caption("이력은 일반 화면에서 수정·삭제하지 않도록 설계했습니다.")

elif page == "설정 및 도움말":
    st.subheader("설정 및 도움말")
    st.markdown("""
    **Google Sheets 연결**
    - `app.py` 상단의 `SHEET_URL`에 실제 스프레드시트 주소를 입력합니다.
    - Streamlit Secrets에 `gcp_service_account` 값을 등록합니다.
    - 서비스 계정 이메일을 해당 스프레드시트의 편집자로 공유해야 합니다.

    **상태와 수량**
    - `대여가능수량`은 지금 대여할 수 있는 수량입니다.
    - 대여 시 대여가능수량이 감소하고, 정상 반납된 수량만 증가합니다.
    - 파손·분실 수량은 대여가능수량에 다시 더하지 않습니다.
    - 수리 중 또는 폐기 상태인 기자재는 대여할 수 없습니다.

    **보안 유의사항**
    - 현재 앱에는 사용자 로그인, 역할별 권한, 감사용 인증이 없습니다.
    - 배포 주소를 아는 사람이 데이터를 변경할 수 있으므로 공개 배포 전에 인증 기능을 추가하세요.
    - 학교 내부의 개인정보를 최소한으로 기록하고, 스프레드시트 공유 권한을 제한하세요.

    **동시 사용 한계**
    - Google Sheets는 트랜잭션 데이터베이스가 아닙니다. 동시에 여러 사용자가 같은 품목을 대여·반납하면 경합이 발생할 수 있습니다.
    - 이 앱은 저장 직전 재고를 다시 확인하지만, 완전한 원자성은 보장하지 않습니다. 다중 사용자 실사용에는 데이터베이스 또는 별도의 잠금 서비스가 더 적합합니다.
    """)
    st.markdown("#### 연결된 시트")
    for name in SHEETS:
        st.write(f"- `{name}`: {len(SHEETS[name])}개 컬럼")
    if st.button("시트 구조 다시 확인", use_container_width=True):
        initialize()
        st.success("필요한 시트와 헤더를 확인했습니다. 기존 데이터는 삭제하지 않았습니다.")
