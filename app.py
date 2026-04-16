import calendar
from datetime import datetime
import html
from io import BytesIO

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas as rl_canvas


# ── 공통 디자인 ─────────────────────────
PASTEL = ["#f6c1d1", "#fde6a7", "#c7f1c9", "#b7dcff"]

pdfmetrics.registerFont(UnicodeCIDFont("HYGothic-Medium"))
KR_FONT = "HYGothic-Medium"


# ── datetime 그대로 파싱 ─────────────────
def _coerce_dt(value):
    if pd.isna(value):
        return None
    ts = pd.to_datetime(value, errors="coerce")
    return None if pd.isna(ts) else ts


# ── 핵심 파싱 (시간 포함 datetime 그대로 사용) ─────
def parse_schedule_df(df, date_col, todo_col=None):
    rows_by_date = {}

    for idx, parsed in df[date_col].map(_coerce_dt).items():
        if parsed is None:
            continue

        sdt = parsed  # ⭐ 핵심: 그대로 사용

        day = sdt.date()
        rows_by_date.setdefault(day, [])

        text = ""
        if todo_col:
            raw = df.at[idx, todo_col]
            if not pd.isna(raw):
                text = str(raw).strip()

        rows_by_date[day].append((sdt, text))

    # 날짜 내 시간 정렬
    todos = {}
    for day, entries in rows_by_date.items():
        entries.sort(key=lambda x: x[0])
        todos[day] = [
            f"{dt.strftime('%H:%M')} {todo}".strip()
            for dt, todo in entries
        ]

    return sorted(todos.keys()), todos


# ── HTML 달력 ─────────────────────────
def render_html_calendar(year, month, todos):
    cal = calendar.Calendar(firstweekday=6)
    weeks = cal.monthdatescalendar(year, month)

    rows = ""
    for week in weeks:
        rows += "<tr>"
        for d in week:
            items = todos.get(d, [])
            inner = ""

            if items:
                shown = items[:4]

                lis = "".join(
                    f'<li>{html.escape(t)}</li>'
                    for t in shown
                )

                more = f"+{len(items)-4}개 더" if len(items) > 4 else ""
                inner = f"<ul>{lis}</ul>{more}"

            rows += f"<td><b>{d.day}</b>{inner}</td>"
        rows += "</tr>"

    return f"<table border='1'>{rows}</table>"


# ── PDF ─────────────────────────
def render_pdf_calendar(year, month, todos):
    cal = calendar.Calendar(firstweekday=6)
    weeks = cal.monthdatescalendar(year, month)

    buf = BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=landscape(A4))

    c.setFont(KR_FONT, 14)
    c.drawString(30, 550, f"{year}년 {month}월")

    y = 500
    for week in weeks:
        x = 30
        for d in week:
            c.drawString(x, y, str(d.day))
            items = todos.get(d, [])
            for i, t in enumerate(items[:3]):
                c.setFont(KR_FONT, 8)
                c.drawString(x, y - (i+1)*10, t)
            x += 100
        y -= 80

    c.save()
    return buf.getvalue()


# ── UI ─────────────────────────
st.set_page_config(layout="wide")
st.title("엑셀 달력 생성기 (멀티시트 버전)")

uploaded_file = st.file_uploader("엑셀 업로드", type=["xlsx", "xls"])

if "all_todos" not in st.session_state:
    st.session_state.all_todos = {}

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)

    sheet = st.selectbox("시트 선택", xls.sheet_names, key="sheet_select")

    df = pd.read_excel(uploaded_file, sheet_name=sheet)

    st.write("미리보기", df.head())

    cols = df.columns.tolist()

    date_col = st.selectbox("날짜 컬럼", cols, key="date_col")
    todo_col = st.selectbox("할일 컬럼", cols, key="todo_col")

    if st.button("➕ 데이터 추가"):
        _, todos = parse_schedule_df(df, date_col, todo_col)

        for d, items in todos.items():
            st.session_state.all_todos.setdefault(d, [])
            st.session_state.all_todos[d].extend(items)

        st.success(f"{sheet} 추가 완료!")


if st.button("🗑 초기화"):
    st.session_state.all_todos = {}


# ── 출력 ─────────────────────────
all_dates = sorted(st.session_state.all_todos.keys())

if all_dates:
    month_options = sorted({(d.year, d.month) for d in all_dates})
    labels = [f"{y}-{m}" for y, m in month_options]

    selected = st.selectbox("월 선택", labels)
    y, m = month_options[labels.index(selected)]

    components.html(
        render_html_calendar(y, m, st.session_state.all_todos),
        height=800
    )

    pdf_bytes = render_pdf_calendar(y, m, st.session_state.all_todos)

    st.download_button(
        "PDF 다운로드",
        pdf_bytes,
        file_name=f"calendar_{y}_{m}.pdf"
    )