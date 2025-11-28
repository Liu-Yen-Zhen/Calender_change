import os
import calendar
import textwrap
from datetime import datetime
from io import BytesIO

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import matplotlib.font_manager as fm
import streamlit as st


# st.subheader("偵測到的系統字型：")

# fonts = fm.findSystemFonts(fontpaths=None)
# for f in fonts:
#     if ("heiti" in f.lower()) or ("gothic" in f.lower()):
#         st.write(f)


# ========= 字型設定 =========
CHINESE_FONT_CANDIDATES = [
    "Heiti TC",
    "AppleGothic",
    "Apple SD Gothic Neo",
    "Noto Sans Gothic",
    "Microsoft JhengHei",
    "Noto Sans CJK TC",
    "STHeiti",
]
# ===========================


import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

def set_chinese_font():
    """
    強制指定 macOS 內建可用的中文/韓文字型，避免 Streamlit + matplotlib 出現豆腐字。
    """
    font_paths = [
        "/System/Library/AssetsV2/com_apple_MobileAsset_Font7/de97612eef4e3cf8ee8e5c0ebd6fd879bbecd23a.asset/AssetData/AppleLiGothic-Medium.ttf"
        "/System/Library/Fonts/Supplemental/AppleGothic.ttf",
        "/System/Library/Fonts/AppleSDGothicNeo.ttc",
        "/System/Library/Fonts/STHeiti Medium.ttc",
        "/System/Library/Fonts/STHeiti Light.ttc",
        "/System/Library/Fonts/Supplemental/NotoSansGothic-Regular.ttf",
    ]

    for fp in font_paths:
        try:
            fm.fontManager.addfont(fp)
        except:
            pass

    # 用第一個最穩定的
    plt.rcParams["font.sans-serif"] = ["AppleGothic"]
    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["axes.unicode_minus"] = False



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
    """整理每一天的事件清單"""
    df[["日期", "星期", "地點"]] = df[["日期", "星期", "地點"]].ffill()
    df = df[df["時間"].notna()]
    df = df[df["日期"].notna()]
    df["日期"] = pd.to_datetime(df["日期"])

    events_by_day = {}

    for _, row in df.iterrows():
        day = int(row["日期"].day)
        time_str = format_time_range(str(row["時間"]))

        # 標籤
        tags = []
        if str(row.get("上課")) == "V": tags.append("上課")
        if str(row.get("借用")) == "V": tags.append("借用")
        if str(row.get("參訪")) == "V": tags.append("參訪")
        tag_text = f"({ '、'.join(tags) })" if tags else ""

        # 活動名稱
        raw_title = row.get("申請事由") or row.get("申請單位")
        if pd.isna(raw_title) or str(raw_title).strip() == "":
            continue

        title = str(raw_title).strip()
        if not time_str:
            continue

        event_line = f"{time_str} {tag_text} {title}".strip()
        events_by_day.setdefault(day, []).append(event_line)

    return events_by_day


def draw_month_calendar(year, month, events_by_day, title_text):
    """畫出月曆並回傳 PNG BytesIO"""

    set_chinese_font()
    calendar.setfirstweekday(calendar.SUNDAY)
    month_matrix = calendar.monthcalendar(year, month)
    weeks = len(month_matrix)

    fig, ax = plt.subplots(figsize=(24, 18))
    ax.set_xlim(0, 7)
    ax.set_ylim(0, weeks)
    ax.axis("off")

    # 標題
    ax.text(0.5, weeks + 0.3, title_text, ha="center",
            fontsize=24, fontweight="bold", transform=ax.transData)

    weekdays = ["週日", "週一", "週二", "週三", "週四", "週五", "週六"]
    for i, wd in enumerate(weekdays):
        ax.text(i + 0.5, weeks - 0.1, wd, ha="center",
                fontsize=16, fontweight="bold")

    for week_idx, week in enumerate(month_matrix):
        for day_idx, day in enumerate(week):
            x = day_idx
            y = weeks - (week_idx + 1)

            ax.add_patch(patches.Rectangle((x, y), 1, 1, fill=False, linewidth=1))

            if day == 0:
                continue

            ax.text(x + 0.05, y + 0.85, str(day),
                    ha="left", va="top", fontsize=14, fontweight="bold")

            events = events_by_day.get(day, [])
            if not events:
                continue

            wrapped_lines = []
            for event in events:
                wrapped = textwrap.fill(event, width=14)
                wrapped_lines.append("• " + wrapped)

            ax.text(x + 0.05, y + 0.75, "\n".join(wrapped_lines),
                    ha="left", va="top", fontsize=12)

    fig.tight_layout()

    buf = BytesIO()
    fig.savefig(buf, dpi=300, format="png", bbox_inches="tight")
    buf.seek(0)
    plt.close(fig)
    return buf


# =============== Streamlit UI ===============

st.set_page_config(page_title="教室月曆產生器", layout="wide")
st.title("信義多功能教室行事曆產生器")

uploaded_file = st.file_uploader("請上傳 Excel 檔 (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names

        st.success(f"成功讀取，共有 {len(sheet_names)} 個月份表。")

        target_sheet = st.selectbox("請選擇要產生日曆的月份工作表", sheet_names)

        generate_btn = st.button("產生行事曆")

        if generate_btn:
            df = pd.read_excel(uploaded_file, sheet_name=target_sheet, header=1)

            # 推算年月
            tmp = df[df["日期"].notna()].copy()
            first_date = pd.to_datetime(tmp["日期"].iloc[0])
            year = first_date.year
            month = first_date.month

            events = build_events_dict(df)
            title = f"{year}年{month:02d}月 多功能教室使用情形"
            png_buf = draw_month_calendar(year, month, events, title)

            st.image(png_buf, use_column_width=True)

            st.download_button(
                label="下載 PNG",
                data=png_buf,
                file_name=f"calendar_{target_sheet}.png",
                mime="image/png"
            )

    except Exception as e:
        st.error(f"讀取 Excel 時發生錯誤：{e}")

else:
    st.info("請先上傳 Excel 檔以開始。")
