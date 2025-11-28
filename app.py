import os
import calendar
import textwrap
from datetime import datetime

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import matplotlib.font_manager as fm
import streamlit as st

# ========= 可調整預設值 =========
DEFAULT_EXCEL_PATH = "信義多功能教室控管表.xlsx"
DEFAULT_YEAR = 2025
DEFAULT_MONTH = 11

CHINESE_FONT_CANDIDATES = [
    "Heiti TC",
    "AppleGothic",
    "Apple SD Gothic Neo",
    "Noto Sans Gothic",
    "Microsoft JhengHei",
    "Noto Sans CJK TC",
    "STHeiti",
]
# ========= 可調整預設值 =========


def get_sheet_name_from_year_month(year: int, month: int, use_roc: bool = True) -> str:
    """
    由西元年 / 月換算出工作表名稱。
    Excel 命名方式：民國年(3位) + 月份(2位)，例如：
    2025/11 -> 11411
    2026/02 -> 11502
    """
    if use_roc:
        roc_year = year - 1911
        return f"{roc_year:03d}{month:02d}"
    else:
        return f"{year}{month:02d}"


def set_chinese_font():
    """設定 matplotlib 的中文字型，避免變成豆腐字。"""
    available = set(f.name for f in fm.fontManager.ttflist)
    for name in CHINESE_FONT_CANDIDATES:
        if name in available:
            plt.rcParams["font.sans-serif"] = [name]
            plt.rcParams["axes.unicode_minus"] = False
            return name
    return None


def format_time_range(raw):
    """把 0900-1200 轉成 09:00-12:00，比較好讀。"""
    if not isinstance(raw, str):
        return ""
    if "-" not in raw:
        return raw
    start, end = raw.split("-")
    start = start.zfill(4)
    end = end.zfill(4)
    return f"{start[:2]}:{start[2:]}-{end[:2]}:{end[2:]}"


def build_events_dict(df: pd.DataFrame):
    """
    依日期整理當天所有活動，輸出:
    {1: ["09:00-12:00 (借用) xxx", "14:00-16:00 (上課) yyy"], ...}
    """
    # 往下填滿日期/星期/地點 (同一天多筆時段會共用第一列的日期)
    df[["日期", "星期", "地點"]] = df[["日期", "星期", "地點"]].ffill()

    # 只留有時間與日期的列，避免一堆空白
    df = df[df["時間"].notna()]
    df = df[df["日期"].notna()]

    # 日期轉成 datetime
    df["日期"] = pd.to_datetime(df["日期"])

    events_by_day = {}

    for _, row in df.iterrows():
        day = int(row["日期"].day)
        time_str = format_time_range(str(row["時間"]))

        # 決定標籤：上課 / 借用 / 參訪
        tags = []
        if "上課" in df.columns and str(row.get("上課")) == "V":
            tags.append("上課")
        if "借用" in df.columns and str(row.get("借用")) == "V":
            tags.append("借用")
        if "參訪" in df.columns and str(row.get("參訪")) == "V":
            tags.append("參訪")
        tag_text = f"({ '、'.join(tags) })" if tags else ""

        # 活動名稱：先用申請事由，沒有再退回申請單位
        raw_title = row.get("申請事由")
        if pd.isna(raw_title) or str(raw_title).strip() == "":
            raw_title = row.get("申請單位")

        # 如果兩個欄位都沒有內容，就視為「沒有用到」，直接跳過
        if pd.isna(raw_title) or str(raw_title).strip() == "":
            continue

        title = str(raw_title).strip()

        # 時間如果是空的也跳過，避免畫出奇怪的空白事件
        if not time_str:
            continue

        # 組合一行字：時間 + 標籤 + 活動名稱
        event_line = f"{time_str} {tag_text} {title}".strip()
        if not event_line:
            continue

        events_by_day.setdefault(day, []).append(event_line)

    return events_by_day


def draw_month_calendar(year, month, events_by_day, title_text):
    """
    繪製月曆並回傳 matplotlib Figure 物件。
    """
    used_font = set_chinese_font()

    calendar.setfirstweekday(calendar.SUNDAY)
    month_matrix = calendar.monthcalendar(year, month)
    weeks = len(month_matrix)

    fig_width = 24
    fig_height = 18
    fig, ax = plt.subplots(figsize=(fig_width, fig_height))
    ax.set_xlim(0, 7)
    ax.set_ylim(0, weeks)
    ax.axis("off")

    # 標題
    ax.text(
        0.5, weeks + 0.3,
        title_text,
        ha="center",
        va="bottom",
        fontsize=20,
        fontweight="bold",
        transform=ax.transData
    )

    # 星期標頭
    weekdays = ["週日", "週一", "週二", "週三", "週四", "週五", "週六"]
    for i, wd in enumerate(weekdays):
        ax.text(
            i + 0.5, weeks - 0.1,
            wd,
            ha="center",
            va="top",
            fontsize=14,
            fontweight="bold"
        )

    # 畫格子 & 填入日期 / 活動
    for week_idx, week in enumerate(month_matrix):
        for day_idx, day in enumerate(week):
            x = day_idx
            y = weeks - (week_idx + 1)

            rect = patches.Rectangle(
                (x, y), 1, 1,
                fill=False,
                linewidth=1
            )
            ax.add_patch(rect)

            if day == 0:
                continue

            # 日期數字
            ax.text(
                x + 0.05, y + 0.85,
                str(day),
                ha="left",
                va="top",
                fontsize=12,
                fontweight="bold"
            )

            events = events_by_day.get(day, [])
            if not events:
                continue

            wrapped_lines = []
            for event in events:
                wrapped = textwrap.fill(event, width=14)
                wrapped_lines.append("• " + wrapped)

            cell_text = "\n".join(wrapped_lines)

            ax.text(
                x + 0.05, y + 0.75,
                cell_text,
                ha="left",
                va="top",
                fontsize=12
            )

    fig.tight_layout()
    return fig


def generate_calendar_png(excel_path: str, year: int, month: int) -> str:
    """讀取 Excel 指定年月，產出月曆 PNG，回傳檔案路徑。"""
    sheet_name = get_sheet_name_from_year_month(year, month)
    output_png = f"calendar_{sheet_name}.png"

    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"找不到檔案：{excel_path}")

    # 讀取對應工作表
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=1)
    events_by_day = build_events_dict(df)
    title_text = f"{year}年{month:02d}月 多功能教室使用情形"

    fig = draw_month_calendar(year, month, events_by_day, title_text)
    fig.savefig(output_png, dpi=300, bbox_inches="tight")
    plt.close(fig)

    return output_png


# ================= Streamlit 介面 =================

st.set_page_config(page_title="信義多功能教室行事曆", layout="wide")

st.title("信義多功能教室行事曆產生器")

with st.sidebar:
    st.header("設定")
    excel_path = st.text_input("Excel 檔案路徑", value=DEFAULT_EXCEL_PATH)

    year = st.number_input(
        "西元年",
        min_value=2024,
        max_value=2035,
        value=DEFAULT_YEAR,
        step=1,
    )
    month = st.selectbox(
        "月份",
        options=list(range(1, 13)),
        index=DEFAULT_MONTH - 1,
        format_func=lambda m: f"{m} 月",
    )

    generate_btn = st.button("產生行事曆")

if generate_btn:
    try:
        png_path = generate_calendar_png(excel_path, int(year), int(month))
        st.success(f"已產生：{png_path}")
        st.image(png_path, use_column_width=True)
    except FileNotFoundError as e:
        st.error(str(e))
    except ValueError as e:
        st.error(f"讀取 Excel 或工作表時發生錯誤：{e}")
    except Exception as e:
        st.error(f"發生未預期錯誤：{e}")
else:
    st.info("請在左側選擇年、月，並按下「產生行事曆」。")
