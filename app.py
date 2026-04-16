import calendar
from datetime import date, datetime, time, timedelta
import html
from io import BytesIO

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

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


# ── 날짜 파싱 (시간 포함 그대로 사용) ───────────────────────
def _coerce_dt(value):
    if pd.isna(value):
        return None
    ts = pd.to_datetime(value, errors="coerce")
    return None if pd.isna(ts) else ts


# ── 핵심: df 기반 파싱 (시간 따로 안씀) ─────────────────────
def parse_schedule_df(df, date_col, todo_col=None):
    rows_by_date = {}

    for idx, parsed in df[date_col].map(_coerce_dt).items():
        if parsed is None:
            continue

        sdt = parsed  # ⭐ 이미 datetime (시간 포함)

        day = sdt.date()
        rows_by_date.setdefault(day, [])

        text = ""
        if todo_col:
            raw = df.at[idx, todo_col]
            if not pd.isna(raw):
                text = str(raw).strip()

        rows_by_date[day].append((sdt, text))

    # ⭐ 날짜 안에서 시간 정렬
    todos = {}
    for day, entries in rows_by_date.items():
        entries.sort(key=lambda x: x[0])  # datetime 기준 정렬

        todos[day] = [
            f"{dt.strftime('%H:%M')} {todo}".strip()
            for dt, todo in entries
        ]

    return sorted(todos.keys()), todos


# ── HTML 달력 ─────────────────────────────────────────────
def render_html_calendar(year, month, todos):
    cal = calendar.Calendar(firstweekday=6)
    weeks = cal.monthdatescalendar(year, month)

    style = """
    <style>
      .cal{border-collapse:collapse;width:100%;table-layout:fixed}
      .cal th,.cal td{border:1px solid #ddd;vertical-align:top;padding:4px 6px;word-break:break-word}
      .cal th{background:#f5f5f5;height:28px;font-size:12px;text-align:center}
      .cal td{height:110px;font-size:13px}
      .out{color:#bbb}
      .dn{font-weight:700;margin-bottom:4px;font-size:13px}
      .tl{margin:0;padding:0;list-style:none}
      .tl li{margin:0 0 3px}
      .tp{
        display:block;
        padding:2px 5px;
        border-radius:5px;
        white-space:nowrap;
        overflow:hidden;
        text-overflow:ellipsis;
        font-size:11px;
        line-height:1.3
      }
      .h1{background:#f6c1d1}
      .h2{background:#fde6a7}
      .h3{background:#c7f1c9}
      .h4{background:#b7dcff}
      .more{font-size:11px;color:#666;margin-top:2px}
    </style>
    """

    rows = ""
    for week in weeks:
        rows += "<tr>"
        for d in week:
            cls = ' class="out"' if d.month != month else ""
            items = todos.get(d, [])

            inner = ""
            if items:
                shown = items[:4]

                lis = "".join(
                    f'<li><span class="tp h{i%4+1}">{html.escape(t)}</span></li>'
                    for i, t in enumerate(shown)
                )

                more = (
                    f'<div class="more">+{len(items)-len(shown)}개 더</div>'
                    if len(items) > len(shown)
                    else ""
                )

                inner = f'<ul class="tl">{lis}</ul>{more}'

            rows += f"""
            <td{cls}>
              <div class="dn">{d.day}</div>
              {inner}
            </td>
            """
        rows += "</tr>"

    return f"""
    {style}
    <table class="cal">
      <thead>
        <tr>
          <th>일</th><th>월</th><th>화</th>
          <th>수</th><th>목</th><th>금</th><th>토</th>
        </tr>
      </thead>
      <tbody>
        {rows}
      </tbody>
    </table>
    """


# ── PDF 달력 (기존 그대로 유지) ───────────────────────────
def render_pdf_calendar(year, month, todos):
    cal = calendar.Calendar(firstweekday=6)
    weeks = cal.monthdatescalendar(year, month)

    PAGE_W, PAGE_H = landscape(A4)
    PAD = 20
    COLS = 7

    grid_x = PAD
    grid_w = PAGE_W - PAD * 2
    cell_w = grid_w / COLS

    TITLE_H = 30
    HEADER_H = 20

    cell_area_h = PAGE_H - PAD * 2 - TITLE_H - HEADER_H
    cell_h = cell_area_h / len(weeks)

    buf = BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=landscape(A4))

    c.setFont(KR_FONT, 16)
    c.drawString(grid_x, PAGE_H - PAD - 20, f"{year}년 {month}월")

    for r, week in enumerate(weeks):
        for col, d in enumerate(week):
            cx = grid_x + col * cell_w
            cy = PAGE_H - PAD - TITLE_H - HEADER_H - (r + 1) * cell_h

            c.rect(cx, cy, cell_w, cell_h)

            c.setFont(KR_FONT, 10)
            c.drawString(cx + 4, cy + cell_h - 14, str(d.day))

            items = todos.get(d, [])
            for i, txt in enumerate(items[:4]):
                py = cy + cell_h - 28 - i * 16
                c.setFont(KR_FONT, 8)
                c.drawString(cx + 6, py + 3, txt)

    c.save()
    return buf.getvalue()


# ── Streamlit UI ─────────────────────────────────────────
st.set_page_config(layout="wide")
st.title("엑셀 달력 생성기")

uploaded_file = st.file_uploader("엑셀 업로드", type=["xlsx", "xls", "csv"])

if "all_todos" not in st.session_state:
    st.session_state.all_todos = {}

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.write(df.head())

    cols = df.columns.tolist()

    date_col = st.selectbox("날짜 컬럼", cols)
    todo_col = st.selectbox("할일 컬럼", cols)

    if st.button("➕ 데이터 추가"):
        _, todos = parse_schedule_df(df, date_col, todo_col)

        for d, items in todos.items():
            st.session_state.all_todos.setdefault(d, [])
            st.session_state.all_todos[d].extend(items)

        st.success("추가 완료!")


# ── 출력 ─────────────────────────────────────────────
all_dates = sorted(st.session_state.all_todos.keys())

if all_dates:
    month_options = sorted({(d.year, d.month) for d in all_dates})
    labels = [f"{y}-{m}" for y, m in month_options]

    selected = st.selectbox("월 선택", labels)
    y, m = month_options[labels.index(selected)]

    components.html(
        render_html_calendar(y, m, st.session_state.all_todos),
        height=800,
        scrolling=True
    )

    pdf_bytes = render_pdf_calendar(y, m, st.session_state.all_todos)

    st.download_button(
        "PDF 다운로드",
        pdf_bytes,
        file_name=f"calendar_{y}_{m}.pdf"
    )