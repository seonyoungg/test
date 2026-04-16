import calendar
from datetime import date
import html
from datetime import datetime, time, timedelta
from io import BytesIO
import os
import tempfile
import unicodedata

import pandas as pd
import pdfkit
import streamlit as st


st.set_page_config(page_title="엑셀 날짜 달력 표시기", layout="wide")
st.title("엑셀 날짜 달력 표시기")
st.caption("엑셀/CSV에 입력된 날짜를 읽어 달력에 표시합니다.")


# ────────────────────────────────────────────────
# 날짜/시간 파싱 유틸
# ────────────────────────────────────────────────

def _coerce_datetime_with_fallback_year(value) -> pd.Timestamp | None:
    if pd.isna(value):
        return None
    ts = pd.to_datetime(value, errors="coerce")
    if pd.isna(ts):
        return None
    try:
        if ts.year < 1970:
            today = date.today()
            ts = ts.replace(year=today.year)
    except Exception:
        pass
    return ts


def _coerce_time(value) -> time | None:
    if pd.isna(value):
        return None
    if isinstance(value, time):
        return value
    if isinstance(value, datetime):
        return value.time()
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime().time()
    if isinstance(value, (int, float)):
        if value < 0:
            return None
        seconds = int(round(float(value) * 24 * 60 * 60))
        seconds = seconds % (24 * 60 * 60)
        return (datetime(2000, 1, 1) + timedelta(seconds=seconds)).time()
    s = str(value).strip()
    if not s:
        return None
    ts = pd.to_datetime(s, errors="coerce")
    if not pd.isna(ts):
        return ts.to_pydatetime().time()
    return None


def parse_schedule(file, file_name: str | None = None) -> tuple[list[date], dict[date, list[str]]]:
    if file_name and str(file_name).lower().endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)
    if df.empty:
        return [], {}

    date_col = None
    for col in df.columns:
        if str(col).strip().lower() in {"date", "날짜", "일시", "datetime"}:
            date_col = col
            break
    if date_col is None:
        date_col = df.columns[0]

    todo_col = None
    for col in df.columns:
        if str(col).strip().lower() in {"메모", "할일", "todo", "task"}:
            todo_col = col
            break

    time_col = None
    for col in df.columns:
        if str(col).strip().lower() in {"시간", "time"}:
            time_col = col
            break

    parsed_dates = df[date_col].map(_coerce_datetime_with_fallback_year)
    rows_by_date: dict[date, list[tuple[pd.Timestamp, str]]] = {}

    for idx, parsed in parsed_dates.items():
        if parsed is None:
            continue
        schedule_dt = parsed
        if time_col is not None:
            t = _coerce_time(df.at[idx, time_col])
            if t is not None:
                schedule_dt = pd.Timestamp.combine(parsed.date(), t)

        day = schedule_dt.date()
        rows_by_date.setdefault(day, [])

        text = ""
        if todo_col is not None:
            raw_todo = df.at[idx, todo_col]
            if not pd.isna(raw_todo):
                text = str(raw_todo).strip()

        if text:
            lines = []
            for chunk in text.splitlines():
                for part in chunk.split(","):
                    item = part.strip().lstrip("-").strip()
                    if item:
                        lines.append(item)
            for item in lines:
                rows_by_date[day].append((schedule_dt, item))
        else:
            rows_by_date[day].append((schedule_dt, ""))

    todos_by_date: dict[date, list[str]] = {}
    for day, entries in rows_by_date.items():
        entries.sort(key=lambda x: x[0])
        formatted: list[str] = []
        for dt_value, todo in entries:
            time_label = dt_value.strftime("%H:%M")
            if todo:
                formatted.append(f"{time_label} {todo}")
            else:
                formatted.append(time_label)
        todos_by_date[day] = formatted

    return sorted(todos_by_date.keys()), todos_by_date


# ────────────────────────────────────────────────
# HTML 달력 렌더러 (웹 보기 & PDF 공용)
# ────────────────────────────────────────────────

def render_month_calendar(
    year: int,
    month: int,
    todos_by_date: dict[date, list[str]],
    for_pdf: bool = False,
) -> str:
    """
    for_pdf=True  → 완전한 HTML 문서(<!DOCTYPE html>…) 반환 – wkhtmltopdf 입력용
    for_pdf=False → <style>+<table> 스니펫 반환 – Streamlit st.markdown() 용
    """
    cal = calendar.Calendar(firstweekday=6)  # 일요일 시작
    weeks = cal.monthdatescalendar(year, month)

    # PDF용 추가 스타일: @page 로 A4 가로 설정, 구글 폰트 불러오기
    pdf_head = ""
    if for_pdf:
        pdf_head = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<style>
  @page {{
    size: A4 landscape;
    margin: 12mm 14mm;
  }}
  body {{
    margin: 0;
    padding: 0;
    font-family: 'Noto Sans KR', 'Apple SD Gothic Neo', '맑은 고딕', sans-serif;
  }}
  h1 {{
    font-size: 20pt;
    font-weight: 700;
    margin: 0 0 6px 0;
    color: #111;
  }}
</style>
</head>
<body>
<h1>{year}년 {month}월</h1>
"""

    # 공용 CSS
    style = """
    <style>
      .calendar {border-collapse: collapse; width: 100%; table-layout: fixed;}
      .calendar th, .calendar td {
        border: 1px solid #ddd;
        vertical-align: top;
        font-size: 13px;
        padding: 4px 6px;
        word-break: break-word;
      }
      .calendar th {
        background: #f5f5f5;
        height: 28px;
        padding: 2px 6px;
        font-size: 12px;
        text-align: center;
      }
      .calendar td {height: 110px;}
      .out-month {color: #bbb;}
      .day-number {font-weight: 700; margin-bottom: 4px; font-size: 13px;}
      .todo-list {margin: 0; padding-left: 0; list-style: none;}
      .todo-list li {margin: 0 0 3px 0;}
      .todo-pill {
        display: block;
        padding: 2px 5px;
        border-radius: 5px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        font-size: 11px;
        line-height: 1.3;
      }
      .hl-1 {background: #f6c1d1;}
      .hl-2 {background: #fde6a7;}
      .hl-3 {background: #c7f1c9;}
      .hl-4 {background: #b7dcff;}
      .more {font-size: 11px; color: #666; margin-top: 2px;}
    </style>
    """

    table = """
    <table class="calendar">
      <thead>
        <tr>
          <th>일</th><th>월</th><th>화</th><th>수</th><th>목</th><th>금</th><th>토</th>
        </tr>
      </thead>
      <tbody>
    """

    for week in weeks:
        table += "<tr>"
        for d in week:
            classes = []
            if d.month != month:
                classes.append("out-month")
            class_attr = f' class="{" ".join(classes)}"' if classes else ""

            todos = todos_by_date.get(d, [])
            items_html = ""
            if todos:
                shown = todos[:4]
                list_items = "".join(
                    f'<li><span class="todo-pill hl-{(i % 4) + 1}">{html.escape(todo)}</span></li>'
                    for i, todo in enumerate(shown)
                )
                more_html = (
                    f'<div class="more">+{len(todos) - len(shown)}개 더</div>'
                    if len(todos) > len(shown)
                    else ""
                )
                items_html = f'<ul class="todo-list">{list_items}</ul>{more_html}'

            table += (
                f'<td{class_attr}>'
                f'<div class="day-number">{d.day}</div>'
                f'{items_html}'
                f'</td>'
            )
        table += "</tr>"

    table += "</tbody></table>"

    if for_pdf:
        return pdf_head + style + table + "</body></html>"
    else:
        return style + table


# ────────────────────────────────────────────────
# HTML → PDF 변환 (wkhtmltopdf via pdfkit)
# ────────────────────────────────────────────────

def html_to_pdf(html_str: str) -> bytes:
    """
    완전한 HTML 문서를 받아 A4 가로 PDF bytes로 반환.
    @page CSS에서 페이지 크기/여백을 지정하므로 pdfkit 옵션은 최소화.
    """
    options = {
        "enable-local-file-access": "",
        "encoding": "UTF-8",
        "quiet": "",
        # wkhtmltopdf가 @page CSS를 존중하도록
        "page-size": "A4",
        "orientation": "Landscape",
        "margin-top": "0mm",
        "margin-right": "0mm",
        "margin-bottom": "0mm",
        "margin-left": "0mm",
        "no-outline": "",
    }
    with tempfile.NamedTemporaryFile(suffix=".html", delete=False, mode="w", encoding="utf-8") as f:
        f.write(html_str)
        tmp_path = f.name
    try:
        pdf_bytes = pdfkit.from_file(tmp_path, False, options=options)
    finally:
        os.unlink(tmp_path)
    return pdf_bytes


# ────────────────────────────────────────────────
# Streamlit UI
# ────────────────────────────────────────────────

uploaded_file = st.file_uploader(
    "엑셀/CSV 파일 업로드 (.xlsx, .xls, .csv)", type=["xlsx", "xls", "csv"]
)

if not uploaded_file:
    st.info(
        "엑셀/CSV 파일을 업로드해 주세요. 날짜 컬럼명은 `날짜`/`date`/`일시`를 권장합니다.\n\n"
        "- 날짜 칸에 `5/1 11:00`처럼 **날짜+시간을 한 번에** 넣어도 됩니다.\n"
        "- 또는 `시간`/`time` 컬럼을 별도로 둬도 됩니다."
    )
    st.stop()

try:
    file_name = getattr(uploaded_file, "name", None)
    dates, todos_by_date = parse_schedule(uploaded_file, file_name)
except Exception as exc:
    st.error(f"엑셀을 읽는 중 오류가 발생했습니다: {exc}")
    st.stop()

if not dates:
    st.warning("유효한 날짜를 찾지 못했습니다. 날짜 형식(예: 2026-04-15)을 확인해 주세요.")
    st.stop()

month_options = sorted({(d.year, d.month) for d in dates})
option_labels = [f"{y}년 {m}월" for y, m in month_options]
selected_label = st.selectbox("표시할 월 선택", option_labels, index=0)
selected_idx = option_labels.index(selected_label)
selected_year, selected_month = month_options[selected_idx]

# ── 웹 달력 표시 ──
st.markdown(
    render_month_calendar(selected_year, selected_month, todos_by_date, for_pdf=False),
    unsafe_allow_html=True,
)

# ── PDF 다운로드 ──
with st.spinner("PDF 생성 중…"):
    pdf_html = render_month_calendar(selected_year, selected_month, todos_by_date, for_pdf=True)
    pdf_bytes = html_to_pdf(pdf_html)

st.download_button(
    label="📄 PDF 다운로드 (A4 가로)",
    data=pdf_bytes,
    file_name=f"calendar_{selected_year}_{selected_month:02d}_A4_landscape.pdf",
    mime="application/pdf",
)

st.write(f"총 날짜 개수: **{len(dates)}개**")
st.write(
    f"선택 월 표시 날짜: "
    f"**{len([d for d in dates if d.year == selected_year and d.month == selected_month])}개**"
)
