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

# ── PDF 달력 ─────────────────────────────────────────────
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

    # ── 제목 ─────────────────
    c.setFont(KR_FONT, 16)
    c.setFillColorRGB(0.1, 0.1, 0.1)
    c.drawString(grid_x, PAGE_H - PAD - 20, f"{year}년 {month}월")

    # ── 요일 헤더 ────────────
    weekdays = ["일", "월", "화", "수", "목", "금", "토"]

    for i, name in enumerate(weekdays):
        x = grid_x + i * cell_w
        y = PAGE_H - PAD - TITLE_H - HEADER_H

        # 배경
        c.setFillColorRGB(0.96, 0.96, 0.96)
        c.rect(x, y, cell_w, HEADER_H, fill=1, stroke=0)

        # 테두리
        c.setStrokeColorRGB(0.85, 0.85, 0.85)
        c.rect(x, y, cell_w, HEADER_H, fill=0, stroke=1)

        # 텍스트
        c.setFillColorRGB(0.2, 0.2, 0.2)
        c.setFont(KR_FONT, 10)
        c.drawCentredString(x + cell_w/2, y + 5, name)

    # ── 날짜 셀 ──────────────
    for r, week in enumerate(weeks):
        for col, d in enumerate(week):
            cx = grid_x + col * cell_w
            cy = PAGE_H - PAD - TITLE_H - HEADER_H - (r + 1) * cell_h

            in_month = d.month == month

            # 셀 테두리
            c.setStrokeColorRGB(0.85, 0.85, 0.85)
            c.rect(cx, cy, cell_w, cell_h, fill=0, stroke=1)

            # 날짜 숫자
            c.setFont(KR_FONT, 10)
            if in_month:
                c.setFillColorRGB(0.1, 0.1, 0.1)
            else:
                c.setFillColorRGB(0.7, 0.7, 0.7)

            c.drawString(cx + 4, cy + cell_h - 14, str(d.day))

            # 일정
            items = todos.get(d, [])
            if not items:
                continue

            max_show = 4
            shown = items[:max_show]

            for i, txt in enumerate(shown):
                py = cy + cell_h - 28 - i * 16

                r_, g_, b_ = hex_to_rgb01(PASTEL[i % 4])

                # pill 배경
                c.setFillColorRGB(r_, g_, b_)
                c.roundRect(cx + 4, py, cell_w - 8, 12, 4, fill=1, stroke=0)

                # 텍스트
                c.setFillColorRGB(0.1, 0.1, 0.1)
                c.setFont(KR_FONT, 8)

                max_chars = int((cell_w - 10) / 5)
                display = txt if len(txt) <= max_chars else txt[:max_chars-1] + "…"

                c.drawString(cx + 6, py + 3, display)

            # +N개 더
            if len(items) > max_show:
                c.setFont(KR_FONT, 7)
                c.setFillColorRGB(0.4, 0.4, 0.4)
                c.drawString(cx + 4, cy + 4, f"+{len(items)-max_show}개 더")

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