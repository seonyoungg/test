import calendar
from datetime import date
import html
from datetime import datetime, time, timedelta
from io import BytesIO
import os
from pathlib import Path
import platform
import unicodedata

import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
from matplotlib import font_manager
import requests


st.set_page_config(page_title="엑셀 날짜 달력 표시기", layout="wide")
st.title("엑셀 날짜 달력 표시기")
st.caption("엑셀/CSV에 입력된 날짜를 읽어 달력에 표시합니다.")


def _coerce_datetime_with_fallback_year(value) -> pd.Timestamp | None:
    """
    다양한 날짜/시간 입력을 최대한 관대하게 Timestamp로 변환.
    - 연도 없이 '4/1 11:00' 같은 형태는 현재 연도로 보정
    """
    if pd.isna(value):
        return None

    ts = pd.to_datetime(value, errors="coerce")
    if pd.isna(ts):
        return None

    # 연도가 비정상적으로 과거로 잡히는 경우(예: 1900) 현재 연도로 보정
    try:
        if ts.year < 1970:
            today = date.today()
            ts = ts.replace(year=today.year)
    except Exception:
        pass

    return ts


def _coerce_time(value) -> time | None:
    """
    엑셀 시간(소수), 문자열('11:00'), time, Timestamp 등을 time으로 변환.
    """
    if pd.isna(value):
        return None

    if isinstance(value, time):
        return value

    if isinstance(value, datetime):
        return value.time()

    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime().time()

    # 엑셀에서 시간만 입력한 셀은 float(하루의 소수)로 들어올 수 있음
    if isinstance(value, (int, float)):
        if value < 0:
            return None
        seconds = int(round(float(value) * 24 * 60 * 60))
        seconds = seconds % (24 * 60 * 60)
        return (datetime(2000, 1, 1) + timedelta(seconds=seconds)).time()

    s = str(value).strip()
    if not s:
        return None

    # '11:00', '11:00:00', '오전 11:00' 등도 최대한 처리
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

        # 기본: 날짜 컬럼 값에 시간이 포함되어 있으면 그대로 사용 (예: "5/1 11:00")
        schedule_dt = parsed
        # 옵션: 시간이 별도 컬럼으로 있으면 날짜+시간으로 조합
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


def _east_asian_weighted_len(s: str) -> float:
    # 한글/한자/일본어 등은 대체로 폭이 넓어서 가중치를 더 줌
    w = 0.0
    for ch in s:
        ea = unicodedata.east_asian_width(ch)
        w += 2.0 if ea in {"W", "F"} else 1.0
    return w


def ellipsize_to_fit(text: str, max_px: float, font_size_px: float) -> str:
    """
    '셀 가로폭' 기준으로 1줄 ellipsis 처리.
    - 정확한 글꼴 메트릭 대신, 폭이 넓은 문자(W/F)에 가중치를 두어 근사합니다.
    """
    s = (text or "").strip()
    if not s:
        return s

    # 대략적인 평균 문자 폭(픽셀) 추정치
    avg_char_px = max(6.0, font_size_px * 0.62)
    capacity = max(1.0, max_px / avg_char_px)

    if _east_asian_weighted_len(s) <= capacity:
        return s

    ell = "…"
    target = max(1.0, capacity - _east_asian_weighted_len(ell))
    out = []
    w = 0.0
    for ch in s:
        ch_w = 2.0 if unicodedata.east_asian_width(ch) in {"W", "F"} else 1.0
        if w + ch_w > target:
            break
        out.append(ch)
        w += ch_w
    return "".join(out).rstrip() + ell


def render_month_calendar(year: int, month: int, todos_by_date: dict[date, list[str]]) -> str:
    cal = calendar.Calendar(firstweekday=6)  # 일요일 시작
    weeks = cal.monthdatescalendar(year, month)

    html_output = """
    <style>
      .calendar {border-collapse: collapse; width: 100%; max-width: 1100px; margin-top: 8px;}
      .calendar th, .calendar td {
        border: 1px solid #ddd;
        width: 14.28%;
        height: 120px;
        vertical-align: top;
        font-size: 13px;
        padding: 4px 6px;
      }
      .calendar th {background: #f5f5f5; height: 34px; padding: 2px 6px; font-size: 12px;}
      .out-month {color: #bbb;}
      .day-number {font-weight: 700; margin-bottom: 4px;}
      .todo-list {margin: 0; padding-left: 0; text-align: left; list-style: none;}
      .todo-list li {line-height: 1.25; margin: 0 0 4px 0;}
      .todo-pill {
        display: block;
        padding: 2px 6px;
        border-radius: 6px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        max-width: 170px;
      }
      .hl-1 {background: #f6c1d1;} /* 분홍(파스텔) */
      .hl-2 {background: #fde6a7;} /* 노랑(파스텔) */
      .hl-3 {background: #c7f1c9;} /* 초록(파스텔) */
      .hl-4 {background: #b7dcff;} /* 파랑(파스텔) */
      .more {font-size: 12px; color: #666; margin-top: 2px;}
    </style>
    <table class="calendar">
      <thead>
        <tr>
          <th>일</th><th>월</th><th>화</th><th>수</th><th>목</th><th>금</th><th>토</th>
        </tr>
      </thead>
      <tbody>
    """

    for week in weeks:
        html_output += "<tr>"
        for d in week:
            classes = []
            if d.month != month:
                classes.append("out-month")
            class_attr = f' class="{" ".join(classes)}"' if classes else ""

            items_html = ""
            todos = todos_by_date.get(d, [])
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

            html_output += f'<td{class_attr}><div class="day-number">{d.day}</div>{items_html}</td>'
        html_output += "</tr>"

    html_output += "</tbody></table>"
    return html_output


PASTEL_COLORS = ["#f6c1d1", "#fde6a7", "#c7f1c9", "#b7dcff"]  # 분홍/노랑/초록/파랑
MAX_TODO_WIDTH_PX = 170
WINDOWS_KOREAN_FONT_PATH = r"C:\Windows\Fonts\malgun.ttf"
NOTO_SANS_KR_URL = (
    "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/Korean/NotoSansKR-Regular.otf"
)


def get_korean_font_properties() -> font_manager.FontProperties:
    """
    PDF/PNG 렌더링에서 한글이 깨지지 않도록 한국어 폰트를 확보합니다.
    우선순위:
    - Windows: 맑은 고딕
    - 설치된 폰트 탐색: Noto Sans KR / NanumGothic 등
    - 없으면 Noto Sans KR를 런타임에 다운로드(캐시) 후 사용
    """
    # 1) Windows 기본 폰트
    if platform.system().lower().startswith("win") and os.path.exists(WINDOWS_KOREAN_FONT_PATH):
        return font_manager.FontProperties(fname=WINDOWS_KOREAN_FONT_PATH)

    # 2) 설치 폰트 탐색(리눅스/맥/윈도우 공통)
    preferred_names = [
        "Noto Sans KR",
        "NotoSansKR",
        "NanumGothic",
        "Nanum Gothic",
        "AppleGothic",
        "Malgun Gothic",
    ]
    try:
        for fp in font_manager.findSystemFonts(fontext="ttf") + font_manager.findSystemFonts(fontext="otf"):
            name = Path(fp).stem.lower()
            if any(p.lower().replace(" ", "") in name.replace(" ", "") for p in preferred_names):
                return font_manager.FontProperties(fname=fp)
    except Exception:
        pass

    # 3) 런타임 다운로드(캐시)
    cache_dir = Path.home() / ".cache" / "calendar-fonts"
    cache_dir.mkdir(parents=True, exist_ok=True)
    font_path = cache_dir / "NotoSansKR-Regular.otf"

    if not font_path.exists():
        resp = requests.get(NOTO_SANS_KR_URL, timeout=30)
        resp.raise_for_status()
        font_path.write_bytes(resp.content)

    try:
        font_manager.fontManager.addfont(str(font_path))
    except Exception:
        pass

    return font_manager.FontProperties(fname=str(font_path))


def render_a4_landscape_asset(
    year: int, month: int, todos_by_date: dict[date, list[str]], output_format: str
) -> bytes:
    """
    A4 가로형(landscape) PNG/PDF로 달력을 렌더링합니다.
    - 안쪽 여백: 48px (배경 기준)
    - 일정은 같은 날짜 내에서 위->아래로, 1~4 색상 순환
    """
    dpi = 150
    a4_landscape_in = (11.69, 8.27)  # inches
    pad_px = 48
    font_prop = get_korean_font_properties()

    cal = calendar.Calendar(firstweekday=6)  # 일요일 시작
    weeks = cal.monthdatescalendar(year, month)

    fig = plt.figure(figsize=a4_landscape_in, dpi=dpi)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_axis_off()

    # 전체 배경
    ax.add_patch(Rectangle((0, 0), 1, 1, facecolor="white", edgecolor="none"))

    # 패딩을 반영한 내용 영역(0~1 좌표계)
    w_px = a4_landscape_in[0] * dpi
    h_px = a4_landscape_in[1] * dpi
    left = pad_px / w_px
    right = 1 - pad_px / w_px
    bottom = pad_px / h_px
    top = 1 - pad_px / h_px

    # 제목 영역
    title_h = 0.07
    ax.text(
        left,
        top,
        f"{year}년 {month}월",
        ha="left",
        va="top",
        fontsize=18,
        fontweight="bold",
        color="#111",
        fontproperties=font_prop,
        transform=ax.transAxes,
    )

    # 캘린더 영역
    grid_top = top - title_h
    grid_bottom = bottom
    grid_left = left
    grid_right = right

    cols = 7
    total_h = grid_top - grid_bottom
    header_h = total_h * 0.07  # 요일 헤더를 얇게
    row_h = (total_h - header_h) / len(weeks)
    cell_w = (grid_right - grid_left) / cols

    # 요일 헤더
    weekdays = ["일", "월", "화", "수", "목", "금", "토"]
    for c, name in enumerate(weekdays):
        x = grid_left + c * cell_w
        y = grid_top - header_h
        ax.add_patch(
            Rectangle((x, y), cell_w, header_h, facecolor="#f5f5f5", edgecolor="#dddddd", linewidth=1)
        )
        ax.text(
            x + cell_w / 2,
            y + header_h / 2,
            name,
            ha="center",
            va="center",
            fontsize=12,
            fontweight="bold",
            color="#222",
            fontproperties=font_prop,
            transform=ax.transAxes,
        )

    # 날짜 셀
    for r, week in enumerate(weeks, start=0):
        for c, d in enumerate(week):
            x = grid_left + c * cell_w
            y = (grid_top - header_h) - (r + 1) * row_h

            in_month = d.month == month
            face = "white"
            ax.add_patch(
                Rectangle((x, y), cell_w, row_h, facecolor=face, edgecolor="#dddddd", linewidth=1)
            )

            day_color = "#111" if in_month else "#bdbdbd"
            ax.text(
                x + 0.01,
                y + row_h - 0.02,
                str(d.day),
                ha="left",
                va="top",
                fontsize=12,
                fontweight="bold",
                color=day_color,
                fontproperties=font_prop,
                transform=ax.transAxes,
            )

            todos = todos_by_date.get(d, [])
            if not todos:
                continue

            # 일정 텍스트 영역(셀 내부)
            inner_left = x + 0.01
            inner_right = x + cell_w - 0.01
            inner_top = y + row_h - 0.055
            inner_bottom = y + 0.012
            available_h = max(0.02, inner_top - inner_bottom)
            max_lines = max(1, min(4, len(todos), int(available_h / 0.024)))
            line_h = available_h / max_lines
            shown = todos[:max_lines]

            for i, text in enumerate(shown):
                cy = inner_top - i * line_h
                color = PASTEL_COLORS[i % 4]
                max_w_axes = min((inner_right - inner_left), MAX_TODO_WIDTH_PX / w_px)
                # 형광펜(파스텔) 배경
                ax.add_patch(
                    Rectangle(
                        (inner_left, cy - line_h * 0.80),
                        max_w_axes,
                        line_h * 0.78,
                        facecolor=color,
                        edgecolor="none",
                        transform=ax.transAxes,
                    )
                )
                ax.text(
                    inner_left + 0.006,
                    cy - line_h * 0.15,
                    ellipsize_to_fit(
                        text=text,
                        max_px=MAX_TODO_WIDTH_PX,
                        font_size_px=(8.8 if len(weeks) == 6 else 9.3) * (dpi / 72.0),
                    ),
                    ha="left",
                    va="top",
                    fontsize=8.8 if len(weeks) == 6 else 9.3,
                    color="#111",
                    fontproperties=font_prop,
                    clip_on=True,
                    transform=ax.transAxes,
                )

            if len(todos) > len(shown):
                ax.text(
                    inner_left,
                    inner_bottom,
                    f"+{len(todos) - len(shown)}개 더",
                    ha="left",
                    va="bottom",
                    fontsize=9,
                    color="#666",
                    fontproperties=font_prop,
                    transform=ax.transAxes,
                )

    buf = BytesIO()
    fig.savefig(buf, format=output_format, dpi=dpi, bbox_inches="tight", pad_inches=0)
    plt.close(fig)
    return buf.getvalue()


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

tab_web, tab_a4 = st.tabs(["웹 보기", "A4 이미지(다운로드)"])

with tab_web:
    st.markdown(
        render_month_calendar(selected_year, selected_month, todos_by_date),
        unsafe_allow_html=True,
    )
    pdf_bytes = render_a4_landscape_asset(selected_year, selected_month, todos_by_date, "pdf")
    st.download_button(
        "웹 보기에서 바로 PDF 다운로드",
        data=pdf_bytes,
        file_name=f"calendar_{selected_year}_{selected_month:02d}_A4_landscape.pdf",
        mime="application/pdf",
    )

with tab_a4:
    png_bytes = render_a4_landscape_asset(selected_year, selected_month, todos_by_date, "png")
    st.image(png_bytes, caption="A4 가로형 미리보기 (PNG)")
    st.download_button(
        "A4 가로형 PNG 다운로드",
        data=png_bytes,
        file_name=f"calendar_{selected_year}_{selected_month:02d}_A4_landscape.png",
        mime="image/png",
    )
    st.download_button(
        "A4 가로형 PDF 다운로드",
        data=pdf_bytes,
        file_name=f"calendar_{selected_year}_{selected_month:02d}_A4_landscape.pdf",
        mime="application/pdf",
    )

st.write(f"총 날짜 개수: **{len(dates)}개**")
st.write(
    f"선택 월 표시 날짜: **{len([d for d in dates if d.year == selected_year and d.month == selected_month])}개**"
)
