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

# ── 다크/라이트 모드 상태 초기화 ──────────────────────────────────────────
if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = False

# ── 다크/라이트 모드 CSS 주입 ─────────────────────────────────────────────
def inject_theme_css(dark: bool):
    if dark:
        theme = """
        <style>
        /* 전체 배경 */
        .stApp, [data-testid="stAppViewContainer"] {
            background-color: #1a1a2e !important;
            color: #e0e0e0 !important;
        }
        [data-testid="stHeader"] {
            background-color: #1a1a2e !important;
        }
        /* 사이드바 */
        [data-testid="stSidebar"] {
            background-color: #16213e !important;
        }
        /* 카드/컨테이너 */
        [data-testid="stVerticalBlock"], .block-container {
            background-color: #1a1a2e !important;
        }
        /* 버튼 */
        .stButton > button {
            background-color: #0f3460 !important;
            color: #e0e0e0 !important;
            border: 1px solid #e94560 !important;
            border-radius: 8px !important;
        }
        .stButton > button:hover {
            background-color: #e94560 !important;
            color: white !important;
        }
        /* 파일 업로더 */
        [data-testid="stFileUploader"] {
            background-color: #16213e !important;
            border: 1px solid #0f3460 !important;
            border-radius: 8px !important;
        }
        /* 셀렉트박스 */
        [data-testid="stSelectbox"] > div > div {
            background-color: #16213e !important;
            color: #e0e0e0 !important;
            border-color: #0f3460 !important;
        }
        /* 탭 */
        .stTabs [data-baseweb="tab-list"] {
            background-color: #16213e !important;
        }
        .stTabs [data-baseweb="tab"] {
            color: #a0a0b0 !important;
        }
        .stTabs [aria-selected="true"] {
            color: #e94560 !important;
            border-bottom-color: #e94560 !important;
        }
        /* 텍스트 */
        h1, h2, h3, p, label, .stMarkdown {
            color: #e0e0e0 !important;
        }
        /* 캡션 */
        .stCaption {
            color: #a0a0b0 !important;
        }
        /* info/warning 박스 */
        [data-testid="stInfoBanner"] {
            background-color: #16213e !important;
            border-left-color: #0f3460 !important;
        }
        /* 달력 테이블 (HTML 렌더) */
        .calendar-dark th {
            background: #16213e !important;
            color: #a0c4ff !important;
        }
        .calendar-dark td {
            background: #1a1a2e !important;
            color: #e0e0e0 !important;
            border-color: #2a2a4e !important;
        }
        .calendar-dark .out-month { color: #444466 !important; }
        .calendar-dark .day-number { color: #c0c0e0 !important; }
        .calendar-dark .more { color: #8888aa !important; }
        /* 다운로드 버튼 */
        [data-testid="stDownloadButton"] > button {
            background-color: #0f3460 !important;
            color: #e0e0e0 !important;
            border: 1px solid #e94560 !important;
            border-radius: 8px !important;
        }
        [data-testid="stDownloadButton"] > button:hover {
            background-color: #e94560 !important;
        }
        </style>
        """
    else:
        theme = """
        <style>
        .stApp, [data-testid="stAppViewContainer"] {
            background-color: #ffffff !important;
            color: #1a1a1a !important;
        }
        [data-testid="stHeader"] {
            background-color: #ffffff !important;
        }
        .stButton > button {
            background-color: #f0f2f6 !important;
            color: #1a1a1a !important;
            border: 1px solid #d0d0d0 !important;
            border-radius: 8px !important;
        }
        .stButton > button:hover {
            background-color: #e0e2e6 !important;
        }
        </style>
        """
    st.markdown(theme, unsafe_allow_html=True)

inject_theme_css(st.session_state.dark_mode)

# ── 헤더: 제목 + 다크모드 토글 버튼 ──────────────────────────────────────
col_title, col_toggle = st.columns([8, 1])
with col_title:
    st.title("엑셀 날짜 달력 표시기")
    st.caption("엑셀/CSV에 입력된 날짜를 읽어 달력에 표시합니다.")
with col_toggle:
    st.write("")
    st.write("")
    icon = "☀️ 라이트" if st.session_state.dark_mode else "🌙 다크"
    if st.button(icon, key="theme_toggle"):
        st.session_state.dark_mode = not st.session_state.dark_mode
        st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# 이하 기존 함수들
# ─────────────────────────────────────────────────────────────────────────────

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


def _east_asian_weighted_len(s: str) -> float:
    w = 0.0
    for ch in s:
        ea = unicodedata.east_asian_width(ch)
        w += 2.0 if ea in {"W", "F"} else 1.0
    return w


def ellipsize_to_fit(text: str, max_px: float, font_size_px: float) -> str:
    s = (text or "").strip()
    if not s:
        return s
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


def render_month_calendar(year: int, month: int, todos_by_date: dict[date, list[str]], dark: bool = False) -> str:
    cal = calendar.Calendar(firstweekday=6)
    weeks = cal.monthdatescalendar(year, month)

    # 다크/라이트 색상 분기
    if dark:
        table_bg     = "#1a1a2e"
        header_bg    = "#16213e"
        header_color = "#a0c4ff"
        border_color = "#2a2a4e"
        cell_bg      = "#1a1a2e"
        text_color   = "#e0e0e0"
        out_color    = "#444466"
        num_color    = "#c0c0e0"
        more_color   = "#8888aa"
    else:
        table_bg     = "#ffffff"
        header_bg    = "#f5f5f5"
        header_color = "#333333"
        border_color = "#dddddd"
        cell_bg      = "#ffffff"
        text_color   = "#111111"
        out_color    = "#bbbbbb"
        num_color    = "#111111"
        more_color   = "#666666"

    html_output = f"""
    <style>
      .calendar {{border-collapse: collapse; width: 100%; max-width: 1100px; margin-top: 8px;}}
      .calendar th, .calendar td {{
        border: 1px solid {border_color};
        width: 14.28%;
        height: 120px;
        vertical-align: top;
        font-size: 13px;
        padding: 4px 6px;
        background-color: {cell_bg};
        color: {text_color};
      }}
      .calendar th {{
        background: {header_bg} !important;
        color: {header_color} !important;
        height: 34px;
        padding: 2px 6px;
        font-size: 12px;
      }}
      .out-month {{color: {out_color} !important;}}
      .day-number {{font-weight: 700; margin-bottom: 4px; color: {num_color};}}
      .todo-list {{margin: 0; padding-left: 0; text-align: left; list-style: none;}}
      .todo-list li {{line-height: 1.25; margin: 0 0 4px 0;}}
      .todo-pill {{
        display: block;
        padding: 2px 6px;
        border-radius: 6px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        max-width: 170px;
        color: #111 !important;
      }}
      .hl-1 {{background: #f6c1d1;}}
      .hl-2 {{background: #fde6a7;}}
      .hl-3 {{background: #c7f1c9;}}
      .hl-4 {{background: #b7dcff;}}
      .more {{font-size: 12px; color: {more_color}; margin-top: 2px;}}
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


PASTEL_COLORS = ["#f6c1d1", "#fde6a7", "#c7f1c9", "#b7dcff"]
MAX_TODO_WIDTH_PX = 170
WINDOWS_KOREAN_FONT_PATH = r"C:\Windows\Fonts\malgun.ttf"
NOTO_SANS_KR_URLS = [
    "https://github.com/notofonts/noto-cjk/raw/refs/heads/main/Sans/OTF/Korean/NotoSansKR-Regular.otf",
    "https://github.com/googlefonts/noto-cjk/raw/master/Sans/OTF/Korean/NotoSansKR-Regular.otf",
]


def get_korean_font_properties() -> font_manager.FontProperties:
    # 0) 로컬 폰트 파일 우선
    local_font = Path(__file__).parent / "fonts" / "NotoSans-Regular.ttf"
    if local_font.exists():
        return font_manager.FontProperties(fname=str(local_font))

    # 1) Windows 기본 폰트
    if platform.system().lower().startswith("win") and os.path.exists(WINDOWS_KOREAN_FONT_PATH):
        return font_manager.FontProperties(fname=WINDOWS_KOREAN_FONT_PATH)

    # 2) 설치 폰트 탐색
    preferred_names = ["Noto Sans KR", "NotoSansKR", "NanumGothic", "Nanum Gothic", "AppleGothic", "Malgun Gothic"]
    try:
        for fp in font_manager.findSystemFonts(fontext="ttf") + font_manager.findSystemFonts(fontext="otf"):
            name = Path(fp).stem.lower()
            if any(p.lower().replace(" ", "") in name.replace(" ", "") for p in preferred_names):
                return font_manager.FontProperties(fname=fp)
    except Exception:
        pass

    # 3) 런타임 다운로드 (여러 URL 시도)
    cache_dir = Path.home() / ".cache" / "calendar-fonts"
    cache_dir.mkdir(parents=True, exist_ok=True)
    font_path = cache_dir / "NotoSansKR-Regular.otf"

    if not font_path.exists():
        last_err = None
        for url in NOTO_SANS_KR_URLS:
            try:
                resp = requests.get(url, timeout=30)
                resp.raise_for_status()
                font_path.write_bytes(resp.content)
                break
            except Exception as e:
                last_err = e
                continue
        else:
            st.warning(f"한국어 폰트 다운로드 실패: {last_err}\n기본 폰트로 대체합니다.")
            return font_manager.FontProperties()

    try:
        font_manager.fontManager.addfont(str(font_path))
    except Exception:
        pass

    return font_manager.FontProperties(fname=str(font_path))


def render_a4_landscape_asset(
    year: int, month: int, todos_by_date: dict[date, list[str]], output_format: str, dark: bool = False
) -> bytes:
    dpi = 150
    a4_landscape_in = (11.69, 8.27)
    pad_px = 48
    font_prop = get_korean_font_properties()

    cal = calendar.Calendar(firstweekday=6)
    weeks = cal.monthdatescalendar(year, month)

    # 다크/라이트 색상 분기
    bg_color      = "#1a1a2e" if dark else "white"
    text_color    = "#e0e0e0" if dark else "#111111"
    out_color     = "#444466" if dark else "#bdbdbd"
    border_color  = "#2a2a4e" if dark else "#dddddd"
    header_bg     = "#16213e" if dark else "#f5f5f5"
    header_text   = "#a0c4ff" if dark else "#222222"

    fig = plt.figure(figsize=a4_landscape_in, dpi=dpi)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_axis_off()

    ax.add_patch(Rectangle((0, 0), 1, 1, facecolor=bg_color, edgecolor="none"))

    w_px = a4_landscape_in[0] * dpi
    h_px = a4_landscape_in[1] * dpi
    left   = pad_px / w_px
    right  = 1 - pad_px / w_px
    bottom = pad_px / h_px
    top    = 1 - pad_px / h_px

    title_h = 0.07
    ax.text(left, top, f"{year}년 {month}월",
            ha="left", va="top", fontsize=18, fontweight="bold",
            color=text_color, fontproperties=font_prop, transform=ax.transAxes)

    grid_top    = top - title_h
    grid_bottom = bottom
    grid_left   = left
    grid_right  = right

    cols     = 7
    total_h  = grid_top - grid_bottom
    header_h = total_h * 0.07
    row_h    = (total_h - header_h) / len(weeks)
    cell_w   = (grid_right - grid_left) / cols

    weekdays = ["일", "월", "화", "수", "목", "금", "토"]
    for c, name in enumerate(weekdays):
        x = grid_left + c * cell_w
        y = grid_top - header_h
        ax.add_patch(Rectangle((x, y), cell_w, header_h,
                                facecolor=header_bg, edgecolor=border_color, linewidth=1))
        ax.text(x + cell_w / 2, y + header_h / 2, name,
                ha="center", va="center", fontsize=12, fontweight="bold",
                color=header_text, fontproperties=font_prop, transform=ax.transAxes)

    for r, week in enumerate(weeks, start=0):
        for c, d in enumerate(week):
            x = grid_left + c * cell_w
            y = (grid_top - header_h) - (r + 1) * row_h

            in_month = d.month == month
            ax.add_patch(Rectangle((x, y), cell_w, row_h,
                                   facecolor=bg_color, edgecolor=border_color, linewidth=1))

            day_color = text_color if in_month else out_color
            ax.text(x + 0.01, y + row_h - 0.02, str(d.day),
                    ha="left", va="top", fontsize=12, fontweight="bold",
                    color=day_color, fontproperties=font_prop, transform=ax.transAxes)

            todos = todos_by_date.get(d, [])
            if not todos:
                continue

            inner_left   = x + 0.01
            inner_right  = x + cell_w - 0.01
            inner_top    = y + row_h - 0.055
            inner_bottom = y + 0.012
            available_h  = max(0.02, inner_top - inner_bottom)
            max_lines    = max(1, min(4, len(todos), int(available_h / 0.024)))
            line_h       = available_h / max_lines
            shown        = todos[:max_lines]

            for i, text in enumerate(shown):
                cy      = inner_top - i * line_h
                color   = PASTEL_COLORS[i % 4]
                max_w_axes = min((inner_right - inner_left), MAX_TODO_WIDTH_PX / w_px)
                ax.add_patch(Rectangle(
                    (inner_left, cy - line_h * 0.80), max_w_axes, line_h * 0.78,
                    facecolor=color, edgecolor="none", transform=ax.transAxes))
                ax.text(
                    inner_left + 0.006, cy - line_h * 0.15,
                    ellipsize_to_fit(text=text, max_px=MAX_TODO_WIDTH_PX,
                                     font_size_px=(8.8 if len(weeks) == 6 else 9.3) * (dpi / 72.0)),
                    ha="left", va="top",
                    fontsize=8.8 if len(weeks) == 6 else 9.3,
                    color="#111", fontproperties=font_prop,
                    clip_on=True, transform=ax.transAxes)

            if len(todos) > len(shown):
                ax.text(inner_left, inner_bottom,
                        f"+{len(todos) - len(shown)}개 더",
                        ha="left", va="bottom", fontsize=9,
                        color="#8888aa" if dark else "#666666",
                        fontproperties=font_prop, transform=ax.transAxes)

    buf = BytesIO()
    fig.savefig(buf, format=output_format, dpi=dpi, bbox_inches="tight", pad_inches=0)
    plt.close(fig)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# 메인 UI
# ─────────────────────────────────────────────────────────────────────────────

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
selected_idx   = option_labels.index(selected_label)
selected_year, selected_month = month_options[selected_idx]

dark = st.session_state.dark_mode

tab_web, tab_a4 = st.tabs(["웹 보기", "A4 이미지(다운로드)"])

with tab_web:
    st.markdown(
        render_month_calendar(selected_year, selected_month, todos_by_date, dark=dark),
        unsafe_allow_html=True,
    )
    pdf_bytes = render_a4_landscape_asset(selected_year, selected_month, todos_by_date, "pdf", dark=dark)
    st.download_button(
        "웹 보기에서 바로 PDF 다운로드",
        data=pdf_bytes,
        file_name=f"calendar_{selected_year}_{selected_month:02d}_A4_landscape.pdf",
        mime="application/pdf",
    )

with tab_a4:
    png_bytes = render_a4_landscape_asset(selected_year, selected_month, todos_by_date, "png", dark=dark)
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
