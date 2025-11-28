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
    自動下載「思源黑體（繁體）」並設定為 matplotlib 中文字型。
    在本機與雲端環境都能用，避免豆腐字。
    """
    cache_dir = Path(".font_cache")
    cache_dir.mkdir(exist_ok=True)
    font_path = cache_dir / "SourceHanSansTC-Regular.otf"

    # 若不存在 → 自動下載
    if not font_path.exists():
        try:
            st.write("首次使用：正在下載中文字型（Source Han Sans TC），請稍候 2～3 秒…")
        except Exception:
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
# 2. 資料處理：把 Excel 整理成 {day: [事件文字…]}
# ============================================

def format_time_range(raw: str) -> str:
    """把 0900-1200 轉成 09:00-12:00。"""
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
    將每一列轉成「某一天的某一個活動」：
    - 時間
    - 上課 / 借用 / 參訪 標籤
    - 申請事由 + 申請單位（都會顯示）
    回傳: {1: ["09:00-12:00 (借用) OO課程｜OO單位", ...], ...}
    """
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
        if str(row.get("上課")) == "V":
            tags.append("上課")
        if str(row.get("借用")) == "V":
            tags.append("借用")
        if str(row.get("參訪")) == "V":
            tags.append("參訪")
        tag_text = f"({'、'.join(tags)})" if tags else ""

        # 申請事由 + 申請單位
        reason_raw = row.get("申請事由")
        unit_raw = row.get("申請單位")

        reason = "" if pd.isna(reason_raw) else str(reason_raw).strip()
        unit = "" if pd.isna(unit_raw) else str(unit_raw).strip()

        # 兩者組合：希望單位一定出現
        if reason and unit:
            desc = f"{reason}｜{unit}"
        elif reason:
            desc = reason
        elif unit:
            desc = unit
        else:
            # 完全沒內容就略過
            continue

        if not time_str:
            continue

        # 一行文字：時間 + 標籤 + 事由/單位
        event_line = f"{time_str} {tag_text} {desc}".strip()
        events_by_day.setdefault(day, []).append(event_line)

    return events_by_day


# ============================================
# 3. 漂亮版月曆繪製
# ============================================

def draw_month_calendar(year, month, events_by_day, title_text):
    """
    繪製漂亮版月曆：
    - 淺灰底色
    - 週末淡色背景
    - 上方標題、星期列底色
    - 每筆活動前面加 •，內含「時間 + 標籤 + 申請事由｜申請單位」
    """
    set_chinese_font()

    calendar.setfirstweekday(calendar.SUNDAY)
    month_matrix = calendar.monthcalendar(year, month)
    weeks = len(month_matrix)

    fig, ax = plt.subplots(figsize=(24, 16))
    fig.patch.set_facecolor("#F7F7F9")
    ax.set_facecolor("#FFFFFF")
    ax.set_xlim(0, 7)
    ax.set_ylim(0, weeks + 0.6)
    ax.axis("off")

    # ---------- 標題區 ----------
    ax.text(
        0.02,
        weeks + 0.45,
        title_text,
        ha="left",
        va="center",
        fontsize=26,
        fontweight="bold",
        transform=ax.transData,
    )

    # 子標題：小字說明
    subtitle = "時間、申請事由與申請單位一覽"
    ax.text(
        0.02,
        weeks + 0.15,
        subtitle,
        ha="left",
        va="center",
        fontsize=14,
        color="#666666",
        transform=ax.transData,
    )

    # ---------- 星期列 ----------
    weekdays = ["週日", "週一", "週二", "週三", "週四", "週五", "週六"]
    header_height = 0.5
    for i, wd in enumerate(weekdays):
        # 星期底色
        rect = patches.Rectangle(
            (i, weeks - 0.3),
            1,
            header_height,
            linewidth=0,
            facecolor="#F0F0F5",
        )
        ax.add_patch(rect)

        ax.text(
            i + 0.5,
            weeks + header_height - 0.45,
            wd,
            ha="center",
            va="center",
            fontsize=16,
            fontweight="bold",
            color="#333333",
        )

    # ---------- 每一格（日曆） ----------
    for week_idx, week in enumerate(month_matrix):
        for day_idx, day in enumerate(week):
            x = day_idx
            y = weeks - (week_idx + 1) - 0.1  # 往下微移，讓格子不要貼到星期列

            # 平日 / 週末不同底色
            if day == 0:
                facecolor = "#FFFFFF"
            elif day_idx in (0, 6):
                # 週日 / 週六
                facecolor = "#FBFBFD"
            else:
                facecolor = "#FFFFFF"

            rect = patches.Rectangle(
                (x, y),
                1,
                1,
                fill=True,
                facecolor=facecolor,
                edgecolor="#DDDDDD",
                linewidth=0.8,
            )
            ax.add_patch(rect)

            if day == 0:
                continue

            # 日期
            ax.text(
                x + 0.05,
                y + 0.95,
                str(day),
                ha="left",
                va="top",
                fontsize=13,
                fontweight="bold",
                color="#333333",
            )

            # 活動內容
            events = events_by_day.get(day, [])
            if not events:
                continue

            wrapped_lines = []
            for event in events:
                # 每個活動做換行，寬度可調整
                wrapped = textwrap.fill(event, width=17)
                wrapped_lines.append("• " + wrapped)

            cell_text = "\n".join(wrapped_lines)

            ax.text(
                x + 0.05,
                y + 0.80,
                cell_text,
                ha="left",
                va="top",
                fontsize=10.5,
                color="#333333",
            )

    fig.tight_layout()

    buf = BytesIO()
    fig.savefig(buf, dpi=300, format="png", bbox_inches="tight")
    buf.seek(0)
    plt.close(fig)
    return buf


# ============================================
# 4. Streamlit UI：側邊欄 + 主畫面
# ============================================

st.set_page_config(page_title="信義多功能教室行事曆產生器", layout="wide")

# 簡單一點的 CSS，讓標題看起來舒服一點
st.markdown(
    """
    <style>
    .main-header {
        font-size: 2.1rem;
        font-weight: 700;
        margin-bottom: 0.3rem;
    }
    .sub-header {
        font-size: 0.95rem;
        color: #888888;
        margin-bottom: 1.2rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="main-header">信義多功能教室行事曆產生器</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-header">上傳控管表 Excel，選擇月份工作表，一鍵產生可下載的月曆 PNG（含申請事由與申請單位）。</div>',
    unsafe_allow_html=True,
)

# ---- 側邊欄：上傳與選項 ----
with st.sidebar:
    st.header("設定")

    uploaded_file = st.file_uploader("請上傳 Excel 檔（.xlsx）", type=["xlsx"])

    sheet_names = []
    xls = None
    if uploaded_file is not None:
        try:
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names
            st.success(f"成功讀取，共有 {len(sheet_names)} 個月份表。")
        except Exception as e:
            st.error(f"讀取 Excel 失敗：{e}")

    target_sheet = None
    if sheet_names:
        target_sheet = st.selectbox("請選擇要產生日曆的月份工作表", sheet_names)

    generate_btn = st.button("產生行事曆", type="primary")

# ---- 主畫面：圖與下載 ----
if uploaded_file is None:
    st.info("請先在左側上傳 Excel 控管表。")
elif not sheet_names:
    st.error("目前無法取得工作表清單，請確認檔案格式是否正確。")
elif not target_sheet:
    st.info("請在左側選擇要產生的月份工作表。")
elif generate_btn:
    try:
        df = pd.read_excel(uploaded_file, sheet_name=target_sheet, header=1)

        # 推出年份與月份
        tmp = df[df["日期"].notna()].copy()
        if tmp.empty:
            st.error("這個工作表裡找不到有效的『日期』欄位資料。")
        else:
            first_date = pd.to_datetime(tmp["日期"].iloc[0])
            year = first_date.year
            month = first_date.month

            events = build_events_dict(df)
            title = f"{year}年{month:02d}月 多功能教室使用情形"

            png_buf = draw_month_calendar(year, month, events, title)

            st.success(f"已產生 {year} 年 {month} 月的使用情形月曆。")
            st.image(png_buf, use_container_width=True)

            st.download_button(
                label="下載這個月曆 PNG 檔",
                data=png_buf,
                file_name=f"calendar_{target_sheet}.png",
                mime="image/png",
            )

    except Exception as e:
        st.error(f"產生行事曆時發生錯誤：{e}")
else:
    st.info("在左側選擇工作表並按下「產生行事曆」。")
