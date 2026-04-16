import calendar
from datetime import date, datetime, time, timedelta
import html
from io import BytesIO
import unicodedata

import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas as rl_canvas


# ── 공통 디자인 상수 ──────────────────────────────────────────
PASTEL = ["#f6c1d1", "#fde6a7", "#c7f1c9", "#b7dcff"]

def hex_to_rgb01(h: str):
    h = h.lstrip("#")
    return tuple(int(h[i:i+2], 16) / 255 for i in (0, 2, 4))

# ── ReportLab 한글 CID 폰트 등록 ──────────────────────────────
pdfmetrics.registerFont(UnicodeCIDFont("HYGothic-Medium"))
KR_FONT = "HYGothic-Medium"


# ────────────────────────────────────────────────────────────────
# 날짜/시간 파싱
# ────────────────────────────────────────────────────────────────
def _coerce_dt(value):
    if pd.isna(value):
        return None
    ts = pd.to_datetime(value, errors="coerce")
    if pd.isna(ts):
        return None
    try:
        if ts.year < 1970:
            ts = ts.replace(year=date.today().year)
    except Exception:
        pass
    return ts


def _coerce_time(value):
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
        sec = int(round(float(value) * 86400)) % 86400
        return (datetime(2000, 1, 1) + timedelta(seconds=sec)).time()
    ts = pd.to_datetime(str(value).strip(), errors="coerce")
    return None if pd.isna(ts) else ts.to_pydatetime().time()


def parse_schedule(file, file_name=None):
    df = pd.read_csv(file) if (file_name or "").lower().endswith(".csv") else pd.read_excel(file)
    if df.empty:
        return [], {}

    date_col = next(
        (c for c in df.columns if str(c).strip().lower() in {"date","날짜","일시","datetime"}),
        df.columns[0],
    )
    todo_col = next(
        (c for c in df.columns if str(c).strip().lower() in {"메모","할일","todo","task"}),
        None,
    )
    time_col = next(
        (c for c in df.columns if str(c).strip().lower() in {"시간","time"}),
        None,
    )

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
        if todo_col is not None:
            raw = df.at[idx, todo_col]
            if not pd.isna(raw):
                text = str(raw).strip()
        if text:
            for chunk in text.splitlines():
                for part in chunk.split(","):
                    item = part.strip().lstrip("-").strip()
                    if item:
                        rows_by_date[day].append((sdt, item))
        else:
            rows_by_date[day].append((sdt, ""))

    todos = {}
    for day, entries in rows_by_date.items():
        entries.sort(key=lambda x: x[0])
        todos[day] = [f"{dt.strftime('%H:%M')} {todo}".strip() for dt, todo in entries]
    return sorted(todos.keys()), todos


# ────────────────────────────────────────────────────────────────
# HTML 달력 렌더러 (웹 표시용)
# ────────────────────────────────────────────────────────────────
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
      .tp{display:block;padding:2px 5px;border-radius:5px;white-space:nowrap;
          overflow:hidden;text-overflow:ellipsis;font-size:11px;line-height:1.3}
      .h1{background:#f6c1d1}.h2{background:#fde6a7}
      .h3{background:#c7f1c9}.h4{background:#b7dcff}
      .more{font-size:11px;color:#666;margin-top:2px}
    </style>"""

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
                more = f'<div class="more">+{len(items)-len(shown)}개 더</div>' if len(items) > len(shown) else ""
                inner = f'<ul class="tl">{lis}</ul>{more}'
            rows += f'<td{cls}><div class="dn">{d.day}</div>{inner}</td>'
        rows += "</tr>"

    return f"""{style}
    <table class="cal">
      <thead><tr>
        <th>일</th><th>월</th><th>화</th><th>수</th><th>목</th><th>금</th><th>토</th>
      </tr></thead>
      <tbody>{rows}</tbody>
    </table>"""


# ────────────────────────────────────────────────────────────────
# ReportLab PDF 달력 (웹과 동일한 색상·구조)
# ────────────────────────────────────────────────────────────────
def render_pdf_calendar(year, month, todos):
    cal = calendar.Calendar(firstweekday=6)
    weeks = cal.monthdatescalendar(year, month)

    PAGE_W, PAGE_H = landscape(A4)
    PAD    = 20
    TITLE  = 28
    HDR_H  = 18
    COLS   = 7

    grid_x = PAD
    grid_w = PAGE_W - PAD * 2
    cell_w = grid_w / COLS

    # 헤더·제목 제외한 셀 영역
    cell_area_h = PAGE_H - PAD * 2 - TITLE - 6 - HDR_H
    cell_h = cell_area_h / len(weeks)

    # y 좌표: reportlab은 아래→위 방향
    title_y    = PAGE_H - PAD - TITLE + 4
    hdr_top_y  = title_y - 10          # 제목 아래 여백
    grid_top_y = hdr_top_y - HDR_H    # 헤더 아래 = 첫 번째 셀 top

    PILL_H   = 13.0
    PILL_GAP = 2.0
    FONT_S   = 8.5
    FONT_DAY = 10
    FONT_TTL = 16
    FONT_HDR = 10

    buf = BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=landscape(A4))
    c.setTitle(f"{year}년 {month}월 달력")

    # ── 제목 ──────────────────────────────────────────────────
    c.setFont(KR_FONT, FONT_TTL)
    c.setFillColorRGB(0.07, 0.07, 0.07)
    c.drawString(grid_x, title_y, f"{year}년 {month}월")

    # ── 요일 헤더 ─────────────────────────────────────────────
    weekdays = ["일", "월", "화", "수", "목", "금", "토"]
    c.setFont(KR_FONT, FONT_HDR)
    for i, name in enumerate(weekdays):
        x = grid_x + i * cell_w
        y = hdr_top_y - HDR_H
        c.setFillColorRGB(0.96, 0.96, 0.96)
        c.rect(x, y, cell_w, HDR_H, fill=1, stroke=0)
        c.setStrokeColorRGB(0.87, 0.87, 0.87)
        c.setLineWidth(0.5)
        c.rect(x, y, cell_w, HDR_H, fill=0, stroke=1)
        c.setFillColorRGB(0.13, 0.13, 0.13)
        c.drawCentredString(x + cell_w / 2, y + HDR_H * 0.25, name)

    # ── 날짜 셀 ───────────────────────────────────────────────
    for r, week in enumerate(weeks):
        for col, d in enumerate(week):
            cx = grid_x + col * cell_w
            cy = grid_top_y - (r + 1) * cell_h   # 셀 bottom-left y

            in_month = d.month == month

            # 셀 배경 + 테두리
            c.setFillColorRGB(1, 1, 1)
            c.setStrokeColorRGB(0.87, 0.87, 0.87)
            c.setLineWidth(0.5)
            c.rect(cx, cy, cell_w, cell_h, fill=1, stroke=1)

            # 날짜 숫자
            c.setFont(KR_FONT, FONT_DAY)
            c.setFillColorRGB(*((.07, .07, .07) if in_month else (.74, .74, .74)))
            c.drawString(cx + 4, cy + cell_h - FONT_DAY - 3, str(d.day))

            # 일정 pills
            items = todos.get(d, [])
            if not items:
                continue

            inner_top = cy + cell_h - FONT_DAY - 9
            inner_bot = cy + 4
            avail_h   = inner_top - inner_bot
            max_show  = min(4, max(1, int(avail_h / (PILL_H + PILL_GAP))))
            shown     = items[:max_show]

            for i, txt in enumerate(shown):
                py     = inner_top - i * (PILL_H + PILL_GAP) - PILL_H
                pill_w = cell_w - 8
                r_, g_, b_ = hex_to_rgb01(PASTEL[i % 4])
                c.setFillColorRGB(r_, g_, b_)
                c.roundRect(cx + 4, py, pill_w, PILL_H, 3, fill=1, stroke=0)
                c.setFillColorRGB(0.07, 0.07, 0.07)
                c.setFont(KR_FONT, FONT_S)
                max_chars = max(4, int(pill_w / (FONT_S * 0.58)))
                display   = txt if len(txt) <= max_chars else txt[:max_chars - 1] + "…"
                c.drawString(cx + 7, py + PILL_H * 0.2, display)

            if len(items) > len(shown):
                c.setFont(KR_FONT, 7.5)
                c.setFillColorRGB(0.4, 0.4, 0.4)
                c.drawString(cx + 4, cy + 2, f"+{len(items)-len(shown)}개 더")

    c.showPage()
    c.save()
    return buf.getvalue()


# ────────────────────────────────────────────────────────────────
# Streamlit UI
# ────────────────────────────────────────────────────────────────
st.set_page_config(page_title="엑셀 날짜 달력 표시기", layout="wide")
st.title("엑셀 날짜 달력 표시기")
st.caption("엑셀/CSV에 입력된 날짜를 읽어 달력에 표시합니다.")

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

month_options  = sorted({(d.year, d.month) for d in dates})
option_labels  = [f"{y}년 {m}월" for y, m in month_options]
selected_label = st.selectbox("표시할 월 선택", option_labels, index=0)
sel_year, sel_month = month_options[option_labels.index(selected_label)]

# 웹 달력 표시
st.markdown(
    render_html_calendar(sel_year, sel_month, todos_by_date),
    unsafe_allow_html=True,
)

# PDF 다운로드 버튼
pdf_bytes = render_pdf_calendar(sel_year, sel_month, todos_by_date)
st.download_button(
    label="📄 PDF 다운로드 (A4 가로)",
    data=pdf_bytes,
    file_name=f"calendar_{sel_year}_{sel_month:02d}_A4_landscape.pdf",
    mime="application/pdf",
)

st.write(f"총 날짜 개수: **{len(dates)}개**")
st.write(
    f"선택 월 표시 날짜: "
    f"**{len([d for d in dates if d.year == sel_year and d.month == sel_month])}개**"
)
