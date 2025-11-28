import os
import calendar
import textwrap
from datetime import datetime
from io import BytesIO
from pathlib import Path
from urllib.request import urlretrieve

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import matplotlib.font_manager as fm
import streamlit as st


# ============================================
# 1. 自動下載中文字型（Source Han Sans TC）
# ============================================

FONT_URL = (
    "https://github.com/adobe-fonts/source-han-sans/raw/release/"
    "OTF/TraditionalChinese/SourceHanSansTC-Regular.otf"
)

def set_chinese_font():
    """
    自動下載「思源黑體（繁體）」，最穩定的中文字型方案。
    能在本地 + GitHub + Streamlit Cloud 正常運作。
    """
    cache_dir = Path(".font_cache")
    cache_dir.mkdir(exist_ok=True)
    font_path = cache_dir / "SourceHanSansTC-Regular.otf"

    # 若不存在 → 自動下載
    if not font_path.exists():
        try:
            st.write("首次使用：正在下載中文字型（Source Han Sans TC），請稍候 2～3 秒…")
        except:
            pass

        try:
            urlretrieve(FONT_URL, font_path)
            print(f"已下載中文字型：{font_path}")
        except Exception as e:
            print("⚠️ 字型下載失敗，可能會變成豆腐字")
            print(e)
            return None

    # 載入字型
    try:
        fm.fontManager.addfont(str(font_path))
        prop = fm.FontProperties(fname=str(font_path))
        font_name = prop.get_name()

        plt.rcParams["font.sans-serif"] = [font_name]
        plt.rcParams["font.family"] = "sans-serif"
        plt.rcParams["axes.unicode_minus"] = False

        print(f"目前中文使用字型：{font_name}")
        return font_name

    except Exception as e:
        print("⚠️ 字型載入失敗")
        print(e)
        return None


# ============================================
# 2. 你的行事曆資料處理函式
# ============================================

def format_time_range(raw):
    """把 0900-1200 轉成 09:00-12:00"""
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



# ============================================
# 3. 生成月曆圖片
# ============================================

def draw_month_calendar(year, month, events_by_day, title_text):
    """畫出月曆並回傳 PNG BytesIO"""

    set_chinese_font()  # 重要：確保 matplot 使用中文字型

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

    # 星期
    weekdays = ["週日", "週一", "週二", "週三", "週四", "週五", "週六"]
    for i, wd in enumerate(weekdays):
        ax.text(i + 0.5, weeks - 0.1, wd, ha="center",
                fontsize=16, fontweight="bold")

    # 月曆格線
    for week_idx, week in enumerate(month_matrix):
        for day_idx, day in enumerate(week):
            x = day_idx
            y = weeks - (week_idx + 1)

            ax.add_patch(patches.Rectangle((x, y), 1, 1, fill=False, linewidth=1))

            if day == 0:
                continue

            # 日期
            ax.text(x + 0.05, y + 0.85, str(day),
                    ha="left", va="top", fontsize=14, fontweight="bold")

            # 活動內容
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



# ============================================
# 4. Streamlit UI
# ============================================

st.set_page_config(page_title="教室月曆產生器", layout="wide")
st.title("多功能教室行事曆產生器")

uploaded_file = st.file_uploader("請上傳 Excel 檔 (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names

        st.success(f"成功讀取，共有 {len(sheet_names)} 個月份表。")

        target_sheet = st.selectbox("請選擇要產生日曆的月份工作表", sheet_names)

        if st.button("產生行事曆"):
            df = pd.read_excel(uploaded_file, sheet_name=target_sheet, header=1)

            tmp = df[df["日期"].notna()].copy()
            first_date = pd.to_datetime(tmp["日期"].iloc[0])

            year = first_date.year
            month = first_date.month

            events = build_events_dict(df)
            title = f"{year}年{month:02d}月 多功能教室使用情形"

            png_buf = draw_month_calendar(year, month, events, title)

            st.image(png_buf, use_container_width=True)

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
