import calendar
from datetime import date
import html
from io import BytesIO

import pandas as pd
import streamlit as st

st.set_page_config(page_title="엑셀 날짜 달력 표시기", layout="wide")
st.title("엑셀 날짜 달력 표시기")
st.caption("엑셀/CSV에 입력된 날짜를 읽어 달력에 표시합니다.")


def parse_schedule(file):
    df = pd.read_excel(file)
    df["날짜"] = pd.to_datetime(df.iloc[:, 0], errors="coerce")

    todos_by_date = {}
    for _, row in df.iterrows():
        if pd.isna(row["날짜"]):
            continue

        d = row["날짜"].date()
        todos_by_date.setdefault(d, [])

        if len(row) > 1 and not pd.isna(row.iloc[1]):
            todos_by_date[d].append(str(row.iloc[1]))

    return todos_by_date


def render_month_calendar(year, month, todos_by_date):
    cal = calendar.Calendar(firstweekday=6)
    weeks = cal.monthdatescalendar(year, month)

    return f"""
    <html>
    <head>
    <meta charset="utf-8">
    <style>
      body {{
        font-family: 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif;
      }}
      .calendar {{
        border-collapse: collapse;
        width: 100%;
      }}
      .calendar th, .calendar td {{
        border: 1px solid #ddd;
        width: 14.28%;
        height: 120px;
        vertical-align: top;
        padding: 6px;
        font-size: 13px;
      }}
      .calendar th {{
        background: #f5f5f5;
      }}
      .day {{
        font-weight: bold;
        margin-bottom: 5px;
      }}
      .todo {{
        display: block;
        padding: 2px 6px;
        border-radius: 6px;
        margin-bottom: 4px;
        background: #fde6a7;
      }}
    </style>
    </head>
    <body>
    <h2>{year}년 {month}월</h2>
    <table class="calendar">
    <tr>
      <th>일</th><th>월</th><th>화</th><th>수</th>
      <th>목</th><th>금</th><th>토</th>
    </tr>
    {"".join([
        "<tr>" + "".join([
            f"<td><div class='day'>{d.day}</div>" +
            "".join([
                f"<div class='todo'>{html.escape(t)}</div>"
                for t in todos_by_date.get(d, [])
            ]) +
            "</td>"
            for d in week
        ]) + "</tr>"
        for week in weeks
    ])}
    </table>
    </body>
    </html>
    """


def html_to_pdf(html_str):
    pdf_io = BytesIO()
    HTML(string=html_str).write_pdf(pdf_io)
    return pdf_io.getvalue()


uploaded_file = st.file_uploader("엑셀 업로드", type=["xlsx"])

if uploaded_file:
    todos_by_date = parse_schedule(uploaded_file)

    today = date.today()
    year, month = today.year, today.month

    html_content = render_month_calendar(year, month, todos_by_date)

    # 웹 출력
    st.markdown(html_content, unsafe_allow_html=True)

    # PDF 변환
    pdf_bytes = html_to_pdf(html_content)

    st.download_button(
        "PDF 다운로드",
        data=pdf_bytes,
        file_name="calendar.pdf",
        mime="application/pdf"
    )