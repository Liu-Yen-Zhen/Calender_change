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


# ================= 字型設定 =================

def set_chinese_font():
    """
    強制指定 macOS 內建可用的中文字型，避免 Streamlit + matplotlib 出現豆腐字。
    會依序嘗試：AppleGothic → AppleSDGothicNeo → STHeiti → NotoSansGothic
    回傳實際使用的字型 family name（除錯用）。
    """
    font_paths = [
        "/System/Library/Fonts/Supplemental/AppleGothic.ttf",
        "/System/Library/Fonts/AppleSDGothicNeo.ttc",
        "/System/Library/Fonts/STHeiti Medium.ttc",
        "/System/Library/Fonts/STHeiti Light.ttc",
        "/System/Library/Fonts/Supplemental/NotoSansGothic-Regular.ttf",
    ]

    chosen_name = None

    for fp in font_paths:
        if not os.path.exists(fp):
            continue
        try:
            # 把字型檔加入 fontManager
            fm.fontManager.addfont(fp)
            # 從檔案讀出真正的 family name
            prop = fm.FontProperties(fname=fp)
            name = prop.get_name()
            # 設定成預設字型
            plt.rcParams["font.sans-serif"] = [name]
            plt.rcParams["font.family"] = "sans-serif"
            plt.rcParams["axes.unicode_minus"] = False
            chosen_name = name
            print(f"使用字型檔：{fp} → family name = {name}")
            break
        except Exception as e:
            print(f"載入字型失敗：{fp}，錯誤：{e}")

    # 如果上面全部失敗，再用名稱 fallback
    if chosen_name is None:
        available = set(f.name for f in fm.fontManager.ttflist)
        fallback_names = [
            "Heiti TC",
            "AppleGothic",
            "Apple SD Gothic Neo",
            "STHeiti",
            "Noto Sans CJK TC",
        ]
        for name in fallback_names:
            if name in available:
                plt.rcParams["font.sans-serif"] = [name]
                plt.rcParams["font.family"] = "sans-serif"
                plt.rcParams["axes.unicode_minus"] = False
                chosen_name = name
                print(f"使用 fallback 字型名稱：{name}")
                break

    if chosen_name is None:
        print("⚠ 找不到可用中文字型，中文可能會變成豆腐字")

    return chosen_name


# ================= 資料處理與畫圖 =================

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
    """整理每一天的事件清單，回傳 {day: [事件文字, ...]}"""
    # 往下填滿日期/星期/地點，讓同一天多筆時段共用
    df[["日期", "星期", "地點"]] = df[["日期", "星期", "地點"]].ffill()

    # 只保留有時間、日期的列
    df = df[df["時間"].notna()]
    df = df[df["日期"].notna()]

    # 日期轉 datetime
    df["日期"] = pd.to_datetime(df["日期"])

    events_by_day = {}

    for _, row in df.iterrows():
        day = int(row["日期"].day)
        time_str = format_time_range(str(row["時間"]))

        # 標籤：上課 / 借用 / 參訪
        tags = []
        if str(row.get("上課")) == "V":
            tags.append("上課")
        if str(row.get("借用")) == "V":
            tags.append("借用")
        if str(row.get("參訪")) == "V":
            tags.append("參訪")
        tag_text = f"({ '、'.join(tags) })" if tags else ""

        # 活動名稱：申請事由 > 申請單位
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
    """畫出月曆並回傳 PNG 的 BytesIO 物件。"""

    used_font = set_chinese_font()
    print(f"目前使用字型：{used_font}")

    calendar.setfirstweekday(calendar.SUNDAY)
    month_matrix = calendar.monthcalendar(year, month)
    weeks = len(month_matrix)

    fig, ax = plt.subplots(figsize=(24, 18))
    ax.set_xlim(0, 7)
    ax.set_ylim(0, weeks)
    ax.axis("off")

    # 標題
    ax.text(
        0.5,
        weeks + 0.3,
        title_text,
        ha="center",
        va="bottom",
        fontsize=24,
        fontweight="bold",
        transform=ax.transData,
    )

    # 星期標頭
    weekdays = ["週日", "週一", "週二", "週三", "週四", "週五", "週六"]
    for i, wd in enumerate(weekdays):
        ax.text(
            i + 0.5,
            weeks - 0.1,
            wd,
            ha="center",
            va="top",
            fontsize=16,
            fontweight="bold",
        )

    # 每一格畫框 + 填內容
    for week_idx, week in enumerate(month_matrix):
        for day_idx, day in enumerate(week):
            x = day_idx
            y = weeks - (week_idx + 1)

            # 方框
            ax.add_patch(
                patches.Rectangle((x, y), 1, 1, fill=False, linewidth=1)
            )

            if day == 0:
                continue

            # 日期
            ax.text(
                x + 0.05,
                y + 0.85,
                str(day),
                ha="left",
                va="top",
                fontsize=14,
                fontweight="bold",
            )

            # 當天活動
            events = events_by_day.get(day, [])
            if not events:
                continue

            wrapped_lines = []
            for event in events:
                wrapped = textwrap.fill(event, width=14)
                wrapped_lines.append("• " + wrapped)

            ax.text(
                x + 0.05,
                y + 0.75,
                "\n".join(wrapped_lines),
                ha="left",
                va="top",
                fontsize=12,
            )

    fig.tight_layout()

    buf = BytesIO()
    fig.savefig(buf, dpi=300, format="png", bbox_inches="tight")
    buf.seek(0)
    plt.close(fig)
    return buf


# ================= Streamlit 介面 =================

st.set_page_config(page_title="教室月曆產生器", layout="wide")
st.title("信義多功能教室行事曆產生器")

uploaded_file = st.file_uploader("請上傳 Excel 檔 (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names

        st.success(f"成功讀取，共有 {len(sheet_names)} 個月份表。")

        target_sheet = st.selectbox(
            "請選擇要產生日曆的月份工作表",
            sheet_names,
        )

        generate_btn = st.button("產生行事曆")

        if generate_btn:
            df = pd.read_excel(uploaded_file, sheet_name=target_sheet, header=1)

            # 從資料裡推算年月
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
                mime="image/png",
            )

    except Exception as e:
        st.error(f"讀取 Excel 時發生錯誤：{e}")
else:
    st.info("請先上傳 Excel 檔以開始。")
