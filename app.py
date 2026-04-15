import calendar
from datetime import date
import html
from datetime import datetime, time, timedelta
from io import BytesIO
import os
from pathlib import Path
import platform
import subprocess
import tempfile
import unicodedata

import pandas as pd
import streamlit as st

st.set_page_config(page_title="엑셀 날짜 달력 표시기", layout="wide")

# ── 다크/라이트 모드 상태 초기화 ─────────────────────────────────────────
if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = False

# ── 테마 CSS 주입 ─────────────────────────────────────────────────────────
def inject_theme_css(dark: bool):
    if dark:
        theme = """
        <style>
        .stApp, [data-testid="stAppViewContainer"] { background-color: #1a1a2e !important; color: #e0e0e0 !important; }
        [data-testid="stHeader"] { background-color: #1a1a2e !important; }
        [data-testid="stSidebar"] { background-color: #16213e !important; }
        .stButton > button {
            background-color: #0f3460 !important; color: #e0e0e0 !important;
            border: 1px solid #e94560 !important; border-radius: 8px !important;
        }
        .stButton > button:hover { background-color: #e94560 !important; color: white !important; }
        [data-testid="stFileUploader"] {
            background-color: #16213e !important; border: 1px solid #0f3460 !important; border-radius: 8px !important;
        }
        [data-testid="stSelectbox"] > div > div { background-color: #16213e !important; color: #e0e0e0 !important; border-color: #0f3460 !important; }
        .stTabs [data-baseweb="tab-list"] { background-color: #16213e !important; }
        .stTabs [data-baseweb="tab"] { color: #a0a0b0 !important; }
        .stTabs [aria-selected="true"] { color: #e94560 !important; border-bottom-color: #e94560 !important; }
        h1, h2, h3, p, label, .stMarkdown { color: #e0e0e0 !important; }
        .stCaption { color: #a0a0b0 !important; }
        [data-testid="stDownloadButton"] > button {
            background-color: #0f3460 !important; color: #e0e0e0 !important;
            border: 1px solid #e94560 !important; border-radius: 8px !important;
        }
        [data-testid="stDownloadButton"] > button:hover { background-color: #e94560 !important; }
        </style>
        """
    else:
        theme = """
        <style>
        .stApp, [data-testid="stAppViewContainer"] { background-color: #ffffff !important; color: #1a1a1a !important; }
        [data-testid="stHeader"] { background-color: #ffffff !important; }
        .stButton > button {
            background-color: #f0f2f6 !important; color: #1a1a1a !important;
            border: 1px solid #d0d0d0 !important; border-radius: 8px !important;
        }
        .stButton > button:hover { background-color: #e0e2e6 !important; }
        </style>
        """
    st.markdown(theme, unsafe_allow_html=True)

inject_theme_css(st.session_state.dark_mode)

# ── 헤더 ─────────────────────────────────────────────────────────────────
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
# 데이터 파싱
# ─────────────────────────────────────────────────────────────────────────────

def _coerce_datetime_with_fallback_year(value):
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
        seconds = int(round(float(value) * 24 * 60 * 60)) % (24 * 60 * 60)
        return (datetime(2000, 1, 1) + timedelta(seconds=seconds)).time()
    s = str(value).strip()
    if not s:
        return None
    ts = pd.to_datetime(s, errors="coerce")
    if not pd.isna(ts):
        return ts.to_pydatetime().time()
    return None


def parse_schedule(file, file_name=None):
    if file_name and str(file_name).lower().endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)
    if df.empty:
        return [], {}

    date_col = next(
        (c for c in df.columns if str(c).strip().lower() in {"date", "날짜", "일시", "datetime"}),
        df.columns[0],
    )
    todo_col = next(
        (c for c in df.columns if str(c).strip().lower() in {"메모", "할일", "todo", "task"}),
        None,
    )
    time_col = next(
        (c for c in df.columns if str(c).strip().lower() in {"시간", "time"}),
        None,
    )

    parsed_dates = df[date_col].map(_coerce_datetime_with_fallback_year)
    rows_by_date: dict = {}

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
            raw = df.at[idx, todo_col]
            if not pd.isna(raw):
                text = str(raw).strip()

        if text:
            for chunk in text.splitlines():
                for part in chunk.split(","):
                    item = part.strip().lstrip("-").strip()
                    if item:
                        rows_by_date[day].append((schedule_dt, item))
        else:
            rows_by_date[day].append((schedule_dt, ""))

    todos_by_date = {}
    for day, entries in rows_by_date.items():
        entries.sort(key=lambda x: x[0])
        formatted = []
        for dt_val, todo in entries:
            label = dt_val.strftime("%H:%M")
            formatted.append(f"{label} {todo}" if todo else label)
        todos_by_date[day] = formatted

    return sorted(todos_by_date.keys()), todos_by_date


# ─────────────────────────────────────────────────────────────────────────────
# HTML 달력 렌더링 (웹 표시 + 내보내기 공용)
# ─────────────────────────────────────────────────────────────────────────────

def build_calendar_html(
    year: int,
    month: int,
    todos_by_date: dict,
    dark: bool = False,
    for_export: bool = False,
) -> str:
    """
    웹 표시(for_export=False)와 PDF/PNG 내보내기(for_export=True) 공용 HTML.
    for_export=True 이면 완전한 <html>...</html> 문서를 반환합니다.
    """
    cal = calendar.Calendar(firstweekday=6)
    weeks = cal.monthdatescalendar(year, month)

    if dark:
        bg         = "#1a1a2e"
        header_bg  = "#16213e"
        header_txt = "#a0c4ff"
        border     = "#2a2a4e"
        cell_bg    = "#1a1a2e"
        txt        = "#e0e0e0"
        out_txt    = "#444466"
        num_txt    = "#c0c0e0"
        more_txt   = "#8888aa"
    else:
        bg         = "#ffffff"
        header_bg  = "#f5f5f5"
        header_txt = "#333333"
        border     = "#dddddd"
        cell_bg    = "#ffffff"
        txt        = "#111111"
        out_txt    = "#bbbbbb"
        num_txt    = "#111111"
        more_txt   = "#666666"

    # 내보내기용: Google Fonts에서 Noto Sans KR 로드 (한글 완벽 지원)
    font_face = ""
    font_family = "sans-serif"
    if for_export:
        # wkhtmltopdf는 @import url이 안 될 수 있으므로 로컬 폰트 우선, 없으면 system sans
        local_font = Path(__file__).parent / "fonts" / "NotoSans-Regular.ttf"
        if local_font.exists():
            font_face = f"""
            @font-face {{
                font-family: 'NotoSansKR';
                src: url('{local_font.as_uri()}');
                font-weight: 400;
            }}
            """
            font_family = "'NotoSansKR', sans-serif"
        else:
            # 시스템 한글 폰트 후보들
            system_fonts = [
                "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
                "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
                "/System/Library/Fonts/AppleGothic.ttf",
            ]
            for fp in system_fonts:
                if Path(fp).exists():
                    ext = Path(fp).suffix.lstrip(".")
                    font_face = f"""
                    @font-face {{
                        font-family: 'SysKorean';
                        src: url('{Path(fp).as_uri()}') format('{ext}');
                    }}
                    """
                    font_family = "'SysKorean', sans-serif"
                    break

    css = f"""
    {font_face}
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{
        background: {bg};
        font-family: {font_family};
        padding: {'20px' if for_export else '0'};
        color: {txt};
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
    }}
    .cal-title {{
        font-size: 20px;
        font-weight: 700;
        color: {num_txt};
        margin-bottom: 10px;
    }}
    .calendar {{
        border-collapse: collapse;
        width: 100%;
        table-layout: fixed;
    }}
    .calendar th {{
        background: {header_bg} !important;
        color: {header_txt} !important;
        border: 1px solid {border};
        padding: 6px 4px;
        font-size: 12px;
        font-weight: 700;
        text-align: center;
        height: 30px;
    }}
    .calendar td {{
        background: {cell_bg} !important;
        color: {txt};
        border: 1px solid {border};
        vertical-align: top;
        width: 14.28%;
        height: 110px;
        padding: 4px 5px;
        font-size: 12px;
        overflow: hidden;
    }}
    .out-month {{ color: {out_txt} !important; }}
    .out-month .day-num {{ color: {out_txt} !important; }}
    .day-num {{
        font-weight: 700;
        font-size: 13px;
        color: {num_txt};
        margin-bottom: 3px;
        display: block;
    }}
    .todo-list {{ list-style: none; margin: 0; padding: 0; }}
    .todo-list li {{ margin-bottom: 3px; }}
    .pill {{
        display: block;
        padding: 2px 5px;
        border-radius: 5px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        font-size: 11px;
        color: #111 !important;
        line-height: 1.4;
    }}
    .hl-1 {{ background: #f6c1d1 !important; }}
    .hl-2 {{ background: #fde6a7 !important; }}
    .hl-3 {{ background: #c7f1c9 !important; }}
    .hl-4 {{ background: #b7dcff !important; }}
    .more {{ font-size: 11px; color: {more_txt}; margin-top: 2px; }}
    """

    # 테이블 본체
    rows_html = ""
    for week in weeks:
        rows_html += "<tr>"
        for d in week:
            in_month = d.month == month
            td_class = "" if in_month else ' class="out-month"'
            todos = todos_by_date.get(d, [])
            shown = todos[:4]

            pills = "".join(
                f'<li><span class="pill hl-{(i % 4) + 1}">{html.escape(t)}</span></li>'
                for i, t in enumerate(shown)
            )
            more = (
                f'<div class="more">+{len(todos) - len(shown)}개 더</div>'
                if len(todos) > len(shown) else ""
            )
            todo_block = f'<ul class="todo-list">{pills}</ul>{more}' if todos else ""

            rows_html += (
                f'<td{td_class}>'
                f'<span class="day-num">{d.day}</span>'
                f'{todo_block}'
                f'</td>'
            )
        rows_html += "</tr>"

    table_html = f"""
    <div class="cal-title">{year}년 {month}월</div>
    <table class="calendar">
      <thead>
        <tr>
          <th>일</th><th>월</th><th>화</th><th>수</th><th>목</th><th>금</th><th>토</th>
        </tr>
      </thead>
      <tbody>{rows_html}</tbody>
    </table>
    """

    if not for_export:
        return f"<style>{css}</style>{table_html}"

    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<style>{css}</style>
</head>
<body>{table_html}</body>
</html>"""


# ─────────────────────────────────────────────────────────────────────────────
# wkhtmltopdf / wkhtmltoimage 변환
# ─────────────────────────────────────────────────────────────────────────────

def _find_wk_tool(name: str):
    candidates = [
        f"/usr/bin/{name}",
        f"/usr/local/bin/{name}",
        f"/opt/homebrew/bin/{name}",
        f"/opt/render/project/.apt/usr/bin/{name}",  # Render apt
    ]
    for p in candidates:
        if Path(p).exists():
            return p
    result = subprocess.run(["which", name], capture_output=True, text=True)
    if result.returncode == 0 and result.stdout.strip():
        return result.stdout.strip()
    return None


def html_to_pdf(html_content: str) -> bytes:
    wk = _find_wk_tool("wkhtmltopdf")
    if wk is None:
        raise RuntimeError(
            "wkhtmltopdf를 찾을 수 없습니다.\n"
            "Render 배포 시 packages.txt에 'wkhtmltopdf'를 추가하세요."
        )
    with tempfile.TemporaryDirectory() as tmp:
        html_path = Path(tmp) / "cal.html"
        pdf_path  = Path(tmp) / "cal.pdf"
        html_path.write_text(html_content, encoding="utf-8")
        cmd = [
            wk,
            "--enable-local-file-access",
            "--page-size",     "A4",
            "--orientation",   "Landscape",
            "--margin-top",    "10mm",
            "--margin-bottom", "10mm",
            "--margin-left",   "10mm",
            "--margin-right",  "10mm",
            "--encoding",      "utf-8",
            "--disable-smart-shrinking",
            "--quiet",
            str(html_path),
            str(pdf_path),
        ]
        res = subprocess.run(cmd, capture_output=True, text=True)
        if res.returncode != 0:
            raise RuntimeError(f"PDF 변환 실패:\n{res.stderr}")
        return pdf_path.read_bytes()


def html_to_png(html_content: str) -> bytes:
    wk = _find_wk_tool("wkhtmltoimage")
    if wk is None:
        raise RuntimeError(
            "wkhtmltoimage를 찾을 수 없습니다.\n"
            "Render 배포 시 packages.txt에 'wkhtmltopdf'를 추가하세요."
        )
    with tempfile.TemporaryDirectory() as tmp:
        html_path = Path(tmp) / "cal.html"
        png_path  = Path(tmp) / "cal.png"
        html_path.write_text(html_content, encoding="utf-8")
        cmd = [
            wk,
            "--enable-local-file-access",
            "--width",    "1400",
            "--quality",  "95",
            "--encoding", "utf-8",
            "--quiet",
            str(html_path),
            str(png_path),
        ]
        res = subprocess.run(cmd, capture_output=True, text=True)
        if res.returncode != 0:
            raise RuntimeError(f"PNG 변환 실패:\n{res.stderr}")
        return png_path.read_bytes()


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
    web_html = build_calendar_html(
        selected_year, selected_month, todos_by_date, dark=dark, for_export=False
    )
    st.markdown(web_html, unsafe_allow_html=True)

    st.write("")
    if st.button("📄 PDF 생성", key="pdf_btn_web"):
        with st.spinner("PDF 생성 중..."):
            try:
                export_html = build_calendar_html(
                    selected_year, selected_month, todos_by_date, dark=dark, for_export=True
                )
                pdf_bytes = html_to_pdf(export_html)
                st.download_button(
                    "⬇️ PDF 다운로드",
                    data=pdf_bytes,
                    file_name=f"calendar_{selected_year}_{selected_month:02d}.pdf",
                    mime="application/pdf",
                    key="pdf_dl_web",
                )
            except Exception as e:
                st.error(str(e))

with tab_a4:
    col1, col2 = st.columns(2)

    with col1:
        if st.button("🖼️ PNG 생성", key="png_btn"):
            with st.spinner("PNG 생성 중..."):
                try:
                    export_html = build_calendar_html(
                        selected_year, selected_month, todos_by_date, dark=dark, for_export=True
                    )
                    png_bytes = html_to_png(export_html)
                    st.image(png_bytes, caption="A4 가로형 미리보기")
                    st.download_button(
                        "⬇️ PNG 다운로드",
                        data=png_bytes,
                        file_name=f"calendar_{selected_year}_{selected_month:02d}.png",
                        mime="image/png",
                        key="png_dl",
                    )
                except Exception as e:
                    st.error(str(e))

    with col2:
        if st.button("📄 PDF 생성", key="pdf_btn_a4"):
            with st.spinner("PDF 생성 중..."):
                try:
                    export_html = build_calendar_html(
                        selected_year, selected_month, todos_by_date, dark=dark, for_export=True
                    )
                    pdf_bytes = html_to_pdf(export_html)
                    st.download_button(
                        "⬇️ PDF 다운로드",
                        data=pdf_bytes,
                        file_name=f"calendar_{selected_year}_{selected_month:02d}.pdf",
                        mime="application/pdf",
                        key="pdf_dl_a4",
                    )
                except Exception as e:
                    st.error(str(e))

st.write(f"총 날짜 개수: **{len(dates)}개**")
st.write(
    f"선택 월 표시 날짜: **{len([d for d in dates if d.year == selected_year and d.month == selected_month])}개**"
)
