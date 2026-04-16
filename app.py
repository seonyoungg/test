import calendar
from datetime import date, datetime, time, timedelta
import html
from io import BytesIO

import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas as rl_canvas


# ── 공통 디자인 ─────────────────────────────────────────────
PASTEL = ["#f6c1d1", "#fde6a7", "#c7f1c9", "#b7dcff"]

def hex_to_rgb01(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i+2], 16) / 255 for i in (0, 2, 4))

pdfmetrics.registerFont(UnicodeCIDFont("HYGothic-Medium"))
KR_FONT = "HYGothic-Medium"


# ── 날짜/시간 파싱 ─────────────────────────────────────────
def _coerce_dt(value):
    if pd.isna(value):
        return None
    ts = pd.to_datetime(value, errors="coerce")
    return None if pd.isna(ts) else ts


def _coerce_time(value):
    if pd.isna(value):
        return None
    if isinstance(value, time):
        return value
    ts = pd.to_datetime(str(value), errors="coerce")
    return None if pd.isna(ts) else ts.to_pydatetime().time()


# ── 핵심: df 기반 파싱 ─────────────────────────────────────
def parse_schedule_df(df, date_col, todo_col=None, time_col=None):
    rows_by_date = {}

    for idx, parsed in df[date_col].map(_coerce_dt).items():
        if parsed is None:
            continue

        sdt = parsed

        if time_col:
            t = _coerce_time(df.at[idx, time_col])
            if t:
                sdt = pd.Timestamp.combine(parsed.date(), t)

        day = sdt.date()
        rows_by_date.setdefault(day, [])

        text = ""
        if todo_col:
            raw = df.at[idx, todo_col]
            if not pd.isna(raw):
                text = str(raw).strip()

        rows_by_date[day].append((sdt, text))

    todos = {}
    for day, entries in rows_by_date.items():
        entries.sort(key=lambda x: x[0])
        todos[day] = [
            f"{dt.strftime('%H:%M')} {todo}".strip()
            for dt, todo in entries
        ]

    return sorted(todos.keys()), todos


# ── HTML 달력 ─────────────────────────────────────────────
def render_html_calendar(year, month, todos):
    cal = calendar.Calendar(firstweekday=6)
    weeks = cal.monthdatescalendar(year, month)

    rows = ""
    for week in weeks:
        rows += "<tr>"
        for d in week:
            cls = ' class="out"' if d.month != month else ""
            items = todos.get(d, [])

            inner = ""
            if items:
                lis = "".join(
                    f'<li><span class="tp h{i%4+1}">{html.escape(t)}</span></li>'
                    for i, t in enumerate(items[:4])
                )
                more = f'<div class="more">+{len(items)-4}개 더</div>' if len(items) > 4 else ""
                inner = f'<ul class="tl">{lis}</ul>{more}'

            rows += f'<td{cls}><b>{d.day}</b>{inner}</td>'
        rows += "</tr>"

    return f"""
    <table border="1" style="width:100%;table-layout:fixed">
    <tr><th>일</th><th>월</th><th>화</th><th>수</th><th>목</th><th>금</th><th>토</th></tr>
    {rows}
    </table>
    """


# ── PDF 달력 ─────────────────────────────────────────────
def render_pdf_calendar(year, month, todos):
    cal = calendar.Calendar(firstweekday=6)
    weeks = cal.monthdatescalendar(year, month)

    buf = BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=landscape(A4))

    c.setFont(KR_FONT, 16)
    c.drawString(30, 550, f"{year}년 {month}월")

    y = 500
    for week in weeks:
        x = 30
        for d in week:
            c.rect(x, y, 100, 60)
            c.setFont(KR_FONT, 8)
            c.drawString(x + 2, y + 45, str(d.day))

            items = todos.get(d, [])
            for i, txt in enumerate(items[:3]):
                c.drawString(x + 2, y + 30 - (i * 10), txt[:10])

            x += 100
        y -= 60

    c.save()
    return buf.getvalue()


# ── Streamlit UI ─────────────────────────────────────────
st.set_page_config(layout="wide")
st.title("엑셀 달력 생성기 (멀티 시트 + 누적)")

uploaded_file = st.file_uploader("엑셀 업로드", type=["xlsx", "xls", "csv"])

if "all_todos" not in st.session_state:
    st.session_state.all_todos = {}

if uploaded_file:
    # CSV / Excel 분기
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        xls = pd.ExcelFile(uploaded_file)
        sheet = st.selectbox("시트 선택", xls.sheet_names)
        df = pd.read_excel(uploaded_file, sheet_name=sheet)

    st.write("데이터 미리보기", df.head())

    cols = df.columns.tolist()

    date_col = st.selectbox("날짜 컬럼", cols)
    todo_col = st.selectbox("할일 컬럼", cols)
    time_col = st.selectbox("시간 컬럼", ["없음"] + cols)

    if time_col == "없음":
        time_col = None

    if st.button("➕ 데이터 추가"):
        _, todos = parse_schedule_df(df, date_col, todo_col, time_col)

        for d, items in todos.items():
            st.session_state.all_todos.setdefault(d, [])
            st.session_state.all_todos[d].extend(items)

        st.success("추가 완료!")

if st.button("🗑 초기화"):
    st.session_state.all_todos = {}

# ── 달력 출력 ─────────────────────────────────────────────
all_dates = sorted(st.session_state.all_todos.keys())

if all_dates:
    month_options = sorted({(d.year, d.month) for d in all_dates})
    labels = [f"{y}-{m}" for y, m in month_options]

    selected = st.selectbox("월 선택", labels)
    y, m = month_options[labels.index(selected)]

    st.markdown(
        render_html_calendar(y, m, st.session_state.all_todos),
        unsafe_allow_html=True,
    )

    pdf_bytes = render_pdf_calendar(y, m, st.session_state.all_todos)

    st.download_button(
        "PDF 다운로드",
        pdf_bytes,
        file_name=f"calendar_{y}_{m}.pdf"
    )