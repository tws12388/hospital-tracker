# -*- coding: utf-8 -*-
"""
醫務企劃室工作進度管制網頁
============================
技術架構：Streamlit + pandas + openpyxl
資料來源：同資料夾內的 Excel 檔案
          「2026-08-25--醫務企劃室待辦事項--辦理情形.xlsx」
          （自動處理副檔名大小寫，例如 .xlsx / .XLSX / .Xlsx 皆可）

Excel 實際結構：
    檔案內有三個工作表(sheet)，分別對應三組：
        醫勤組 / 行政組 / 通資組（畫面上分頁名稱顯示為「通資電組」）
    （若日後工作表名稱有異動，請修改下方 SHEET_NAME_MAP）

每個工作表欄位：
    案號 / 申請單位 / 工作項目 / 負責參謀 / 作業進度 / 預計完成日 / 實際完成日 / 目前狀態
"""

import glob
import os
from datetime import datetime, date, timedelta
from io import BytesIO

import pandas as pd
import streamlit as st

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ============================================================
# 中文字型設定（內附「文泉驛正黑」TrueType 字型，確保任何電腦開啟 PDF 都能正常顯示中文）
# 註：系統常見的 Noto Sans CJK 屬於 OpenType/CFF 外框格式，reportlab 無法內嵌，
#     故改用 TrueType(glyf) 外框的文泉驛正黑，繁體中文涵蓋完整。
# ============================================================
APP_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_PATH = os.path.join(APP_DIR, "fonts", "WenQuanYiZenHei.ttf")
FONT_NAME = "CJKFont"
FONT_NAME_BOLD = "CJKFont-Bold"  # 同一套字型檔（來源無獨立粗體），以顏色/底色/字級呈現視覺層級


def register_pdf_fonts():
    registered = pdfmetrics.getRegisteredFontNames()
    if FONT_NAME not in registered:
        pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH))
    if FONT_NAME_BOLD not in registered:
        pdfmetrics.registerFont(TTFont(FONT_NAME_BOLD, FONT_PATH))

# ============================================================
# 基本設定
# ============================================================
st.set_page_config(
    page_title="醫務企劃室工作進度管制",
    page_icon="🏥",
    layout="wide",
)

BASE_FILENAME = "2026-08-25--醫務企劃室待辦事項--辦理情形"

# 三組對應的工作表名稱：key 為畫面顯示的組名，value 為 Excel 實際工作表名稱
SHEET_NAME_MAP = {
    "醫勤組": "醫勤組",
    "行政組": "行政組",
    "通資電組": "通資組",   # Excel 內實際工作表名稱為「通資組」
}
GROUP_NAMES = list(SHEET_NAME_MAP.keys())

# 欄位名稱（依實際 Excel 欄位）
COL_CASE_NO = "案號"
COL_APPLY_UNIT = "申請單位"
COL_TASK = "工作項目"
COL_OWNER = "負責參謀"
COL_PROGRESS = "作業進度"
COL_DUE_DATE = "預計完成日"
COL_ACTUAL_DATE = "實際完成日"
COL_STATUS = "目前狀態"

# 完整欄位順序（存回 Excel 時依此順序輸出，與原始檔案一致）
REQUIRED_COLUMNS = [
    COL_CASE_NO, COL_APPLY_UNIT, COL_TASK, COL_OWNER,
    COL_PROGRESS, COL_DUE_DATE, COL_ACTUAL_DATE, COL_STATUS,
]

DONE_KEYWORDS = ["已完成", "完工", "結案", "已結案", "完成"]

# 實際使用的狀態值（可視需要增加）
STATUS_OPTIONS = ["辦理中", "已完成", "已逾期", "暫緩", "其他"]


# ============================================================
# 工具函式：檔案尋找（自動處理副檔名大小寫）
# ============================================================
def find_excel_file(base_dir="."):
    """在資料夾內尋找符合基本檔名、副檔名大小寫不拘的 Excel 檔案"""
    candidates = []
    for ext in ("xlsx", "xls", "XLSX", "XLS", "Xlsx", "Xls"):
        candidates.extend(glob.glob(os.path.join(base_dir, f"{BASE_FILENAME}.{ext}")))

    if candidates:
        return candidates[0]

    # 若完全找不到指定檔名，退而求其次：資料夾內任何 .xlsx/.xls 檔（不分大小寫）
    all_files = glob.glob(os.path.join(base_dir, "*"))
    fallback = [
        f for f in all_files
        if os.path.splitext(f)[1].lower() in (".xlsx", ".xls")
    ]
    return fallback[0] if fallback else None


# ============================================================
# 資料讀取 / 寫入
# ============================================================
@st.cache_data(show_spinner=False)
def load_data(file_path, mtime):
    """讀取三個分組工作表，回傳 dict[組名] = DataFrame，並統一補齊欄位"""
    data = {}
    xls = pd.ExcelFile(file_path)
    sheet_names_in_file = xls.sheet_names

    for group, sheet_name in SHEET_NAME_MAP.items():
        actual_sheet = None
        if sheet_name in sheet_names_in_file:
            actual_sheet = sheet_name
        else:
            for s in sheet_names_in_file:
                if s.strip() == sheet_name.strip():
                    actual_sheet = s
                    break

        if actual_sheet is None:
            data[group] = pd.DataFrame(columns=REQUIRED_COLUMNS)
            continue

        df = pd.read_excel(file_path, sheet_name=actual_sheet)

        for col in REQUIRED_COLUMNS:
            if col not in df.columns:
                df[col] = None

        df[COL_DUE_DATE] = pd.to_datetime(df[COL_DUE_DATE], errors="coerce")
        df[COL_ACTUAL_DATE] = pd.to_datetime(df[COL_ACTUAL_DATE], errors="coerce")
        df[COL_STATUS] = df[COL_STATUS].fillna("").astype(str)
        df[COL_PROGRESS] = df[COL_PROGRESS].fillna("").astype(str)
        df[COL_OWNER] = df[COL_OWNER].fillna("").astype(str).replace("nan", "")
        df[COL_APPLY_UNIT] = df[COL_APPLY_UNIT].fillna("").astype(str).replace("nan", "")
        df[COL_CASE_NO] = df[COL_CASE_NO].astype(str)

        df = df.reset_index(drop=True)
        data[group] = df

    return data, sheet_names_in_file


def save_group_data(file_path, group, df):
    """把單一組別的 DataFrame 寫回原 Excel 檔案（只覆寫該分頁，其餘分頁保留）"""
    sheet_name = SHEET_NAME_MAP[group]
    try:
        with pd.ExcelWriter(
            file_path, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        return True, None
    except Exception as e:
        return False, str(e)


def next_case_no(group):
    """為新案件自動產生該組別的下一個案號（沿用既有數字序號規則，若無法判斷則從1開始）"""
    df = st.session_state["all_data"][group]
    if df.empty:
        return "1"
    nums = pd.to_numeric(df[COL_CASE_NO], errors="coerce").dropna()
    if nums.empty:
        return "1"
    return str(int(nums.max()) + 1)


# ============================================================
# 紅黃綠燈判斷邏輯
# ============================================================
def is_done(status: str) -> bool:
    status = (status or "").strip()
    return any(k in status for k in DONE_KEYWORDS)


def get_light(row) -> str:
    """回傳 'red' / 'yellow' / 'green' / 'gray'（gray = 未完工但日期未知或尚早）"""
    status = row[COL_STATUS]
    due = row[COL_DUE_DATE]

    if is_done(status):
        return "green"

    if pd.isna(due):
        return "gray"

    today = pd.Timestamp(date.today())
    due_date = pd.Timestamp(due).normalize()
    days_left = (due_date - today).days

    if days_left < 0:
        return "red"
    elif days_left <= 2:
        return "yellow"
    else:
        return "gray"


LIGHT_EMOJI = {"red": "🔴", "yellow": "🟡", "green": "🟢", "gray": "⚪"}
LIGHT_LABEL = {"red": "逾期", "yellow": "即將到期", "green": "已完工", "gray": "進行中"}


def add_light_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if df.empty:
        df["燈號"] = []
        df["燈號說明"] = []
        return df
    df["燈號"] = df.apply(get_light, axis=1)
    df["燈號說明"] = df["燈號"].map(lambda x: f"{LIGHT_EMOJI[x]} {LIGHT_LABEL[x]}")
    return df


def style_row_by_light(row):
    colors_map = {
        "red": "background-color:#ffe3e3; color:#c81e1e;",
        "yellow": "background-color:#fff9db; color:#e67700;",
        "green": "background-color:#ebfbee; color:#2b8a3e;",
        "gray": "",
    }
    style = colors_map.get(row.get("燈號", "gray"), "")
    return [style] * len(row)


def get_this_week_range(ref_date=None):
    """回傳本週（週一~週日）的起訖日期"""
    ref_date = ref_date or date.today()
    monday = ref_date - timedelta(days=ref_date.weekday())
    sunday = monday + timedelta(days=6)
    return monday, sunday


# ============================================================
# 產生「主管週報 PDF」
# ============================================================
def generate_weekly_report_pdf(all_data: dict) -> bytes:
    register_pdf_fonts()

    week_start, week_end = get_this_week_range()
    today_str = date.today().strftime("%Y-%m-%d")
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")

    # ---- 彙整資料 ----
    all_df = pd.concat(
        [df.assign(**{"組別": g}) for g, df in all_data.items()],
        ignore_index=True,
    ) if all_data else pd.DataFrame(columns=REQUIRED_COLUMNS + ["組別"])
    all_df_light = add_light_columns(all_df)

    total_count = len(all_df_light)
    done_count = int((all_df_light["燈號"] == "green").sum()) if total_count else 0
    red_count = int((all_df_light["燈號"] == "red").sum()) if total_count else 0
    yellow_count = int((all_df_light["燈號"] == "yellow").sum()) if total_count else 0
    completion_rate = (done_count / total_count * 100) if total_count else 0.0

    group_rows = []
    for g in GROUP_NAMES:
        df = all_data.get(g, pd.DataFrame(columns=REQUIRED_COLUMNS))
        df_l = add_light_columns(df)
        n = len(df_l)
        d = int((df_l["燈號"] == "green").sum()) if n else 0
        r = int((df_l["燈號"] == "red").sum()) if n else 0
        y = int((df_l["燈號"] == "yellow").sum()) if n else 0
        rate = round(d / n * 100, 1) if n else 0.0
        group_rows.append({"組別": g, "總案件": n, "已完工": d, "完工率": rate, "紅燈": r, "黃燈": y})

    # 落後組別判斷：完工率最低，若同分則以紅燈數最多者優先
    ranked = sorted(group_rows, key=lambda x: (x["完工率"], -x["紅燈"]))
    worst = ranked[0] if ranked else None
    if worst and (worst["紅燈"] > 0 or worst["完工率"] < 50) and any(g["總案件"] > 0 for g in group_rows):
        insight_text = (
            f"【提醒】本週落後組別為「{worst['組別']}」，完工率 {worst['完工率']}%，"
            f"目前有 {worst['紅燈']} 件逾期案件，建議優先追蹤與協調資源。"
        )
    else:
        insight_text = "【良好】各組進度大致穩定，目前無明顯落後組別。"

    overdue_df = all_df_light[all_df_light["燈號"] == "red"].copy()
    if not overdue_df.empty:
        overdue_df["逾期天數"] = (pd.Timestamp(date.today()) - overdue_df[COL_DUE_DATE]).dt.days
        overdue_df = overdue_df.sort_values("逾期天數", ascending=False)

    # ---- 樣式 ----
    title_style = ParagraphStyle(
        "TitleTC", fontName=FONT_NAME_BOLD, fontSize=20, leading=26, alignment=TA_CENTER,
        textColor=colors.HexColor("#12293f"), spaceAfter=8,
    )
    subtitle_style = ParagraphStyle(
        "SubtitleTC", fontName=FONT_NAME, fontSize=10.5, leading=14, alignment=TA_CENTER,
        textColor=colors.HexColor("#5b6b7a"), spaceAfter=2,
    )
    section_style = ParagraphStyle(
        "SectionTC", fontName=FONT_NAME_BOLD, fontSize=13, alignment=TA_LEFT,
        textColor=colors.white, spaceBefore=2, spaceAfter=2, leftIndent=6,
    )
    body_style = ParagraphStyle(
        "BodyTC", fontName=FONT_NAME, fontSize=9, alignment=TA_LEFT, leading=13,
    )
    insight_warn_style = ParagraphStyle(
        "InsightWarnTC", fontName=FONT_NAME, fontSize=10.5, alignment=TA_LEFT,
        leading=15, textColor=colors.HexColor("#8a4b00"),
        backColor=colors.HexColor("#fff6e0"), borderPadding=8,
    )
    insight_good_style = ParagraphStyle(
        "InsightGoodTC", fontName=FONT_NAME, fontSize=10.5, alignment=TA_LEFT,
        leading=15, textColor=colors.HexColor("#1a6b34"),
        backColor=colors.HexColor("#e8f8ee"), borderPadding=8,
    )
    kpi_label_style = ParagraphStyle(
        "KpiLabel", fontName=FONT_NAME, fontSize=9.5, alignment=TA_CENTER, textColor=colors.white,
    )
    kpi_value_style = ParagraphStyle(
        "KpiValue", fontName=FONT_NAME_BOLD, fontSize=20, alignment=TA_CENTER, textColor=colors.white,
    )

    def section_bar(text):
        t = Table([[Paragraph(text, section_style)]], colWidths=[doc_width])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#12293f")),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        return t

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        topMargin=16 * mm, bottomMargin=14 * mm, leftMargin=16 * mm, rightMargin=16 * mm,
        title="醫務企劃室主管週報",
    )
    doc_width = doc.width

    elements = []
    elements.append(Paragraph("醫務企劃室　主管週報", title_style))
    elements.append(Spacer(1, 4))
    elements.append(Paragraph(
        f"本週期間：{week_start.strftime('%Y-%m-%d')} ～ {week_end.strftime('%Y-%m-%d')}　"
        f"｜　報表產出時間：{generated_at}",
        subtitle_style,
    ))
    elements.append(Spacer(1, 10))

    # ---- KPI 卡片列 ----
    kpi_defs = [
        ("總案件數", f"{total_count} 件", "#1c3d5a"),
        ("已完工數", f"{done_count} 件", "#2b8a3e"),
        ("整體完工率", f"{completion_rate:.1f}%", "#0b6e99"),
        ("逾期紅燈案件", f"{red_count} 件", "#c81e1e"),
    ]
    kpi_cell_w = doc_width / 4
    label_row = [Paragraph(l, kpi_label_style) for l, _, _ in kpi_defs]
    value_row = [Paragraph(v, kpi_value_style) for _, v, _ in kpi_defs]
    kpi_table = Table([label_row, value_row], colWidths=[kpi_cell_w] * 4)
    kpi_style_cmds = [
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 1), (-1, 1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 2),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LINEAFTER", (0, 0), (-2, -1), 1, colors.white),
    ]
    for i, (_, _, hexcolor) in enumerate(kpi_defs):
        kpi_style_cmds.append(("BACKGROUND", (i, 0), (i, -1), colors.HexColor(hexcolor)))
    kpi_table.setStyle(TableStyle(kpi_style_cmds))
    elements.append(kpi_table)
    elements.append(Spacer(1, 14))

    # ---- 三組落後分析 ----
    elements.append(section_bar("三組完工率與落後分析"))
    elements.append(Spacer(1, 6))

    group_header = ["組別", "總案件", "已完工", "完工率", "逾期", "即將到期"]
    group_table_data = [group_header]
    for g in group_rows:
        group_table_data.append([
            g["組別"], str(g["總案件"]), str(g["已完工"]),
            f"{g['完工率']}%", str(g["紅燈"]), str(g["黃燈"]),
        ])
    g_col_w = [doc_width * w for w in (0.22, 0.16, 0.16, 0.16, 0.15, 0.15)]
    group_table = Table(group_table_data, colWidths=g_col_w, repeatRows=1)
    g_style_cmds = [
        ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
        ("FONTNAME", (0, 0), (-1, 0), FONT_NAME_BOLD),
        ("FONTSIZE", (0, 0), (-1, -1), 9.5),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e8edf2")),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e0")),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]
    for row_i, g in enumerate(group_rows, start=1):
        if worst and g["組別"] == worst["組別"] and insight_text.startswith("【提醒】"):
            g_style_cmds.append(("BACKGROUND", (0, row_i), (-1, row_i), colors.HexColor("#ffe3e3")))
    group_table.setStyle(TableStyle(g_style_cmds))
    elements.append(group_table)
    elements.append(Spacer(1, 8))
    insight_final_style = insight_warn_style if insight_text.startswith("【提醒】") else insight_good_style
    elements.append(Paragraph(insight_text, insight_final_style))
    elements.append(Spacer(1, 14))

    # ---- 紅燈逾期名單 ----
    elements.append(section_bar(f"紅燈逾期案件優先清單（共 {len(overdue_df)} 件）"))
    elements.append(Spacer(1, 6))

    if overdue_df.empty:
        elements.append(Paragraph("目前沒有逾期案件，全室進度良好！", body_style))
    else:
        od_header = ["組別", "案號", "工作項目", "負責參謀", "預計完成日", "逾期天數"]
        od_data = [od_header]
        for _, r in overdue_df.iterrows():
            od_data.append([
                r["組別"], r[COL_CASE_NO],
                Paragraph(str(r[COL_TASK]), body_style),
                r[COL_OWNER] if r[COL_OWNER] else "—",
                r[COL_DUE_DATE].strftime("%Y-%m-%d"),
                f"{int(r['逾期天數'])} 天",
            ])
        od_col_w = [doc_width * w for w in (0.13, 0.09, 0.36, 0.14, 0.15, 0.13)]
        od_table = Table(od_data, colWidths=od_col_w, repeatRows=1)
        od_table.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
            ("FONTNAME", (0, 0), (-1, 0), FONT_NAME_BOLD),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#c81e1e")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("ALIGN", (0, 0), (1, -1), "CENTER"),
            ("ALIGN", (3, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e0")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#fff5f5")]),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        elements.append(od_table)

    def footer(canvas, doc_):
        canvas.saveState()
        canvas.setFont(FONT_NAME, 8)
        canvas.setFillColor(colors.HexColor("#9aa5b1"))
        canvas.drawString(16 * mm, 10 * mm, f"醫務企劃室工作進度管制系統　自動產出　{today_str}")
        canvas.drawRightString(A4[0] - 16 * mm, 10 * mm, f"第 {doc_.page} 頁")
        canvas.restoreState()

    doc.build(elements, onFirstPage=footer, onLaterPages=footer)
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
# 側邊欄：資料來源設定 / 重新整理
# ============================================================
st.sidebar.title("⚙️ 資料來源設定")

search_dir = st.sidebar.text_input("Excel 檔案所在資料夾", value=".")
file_path = find_excel_file(search_dir)

if file_path is None:
    st.sidebar.error("⚠️ 找不到符合檔名的 Excel 檔案")
    st.error(
        f"❌ 在資料夾「{os.path.abspath(search_dir)}」中找不到檔案：\n\n"
        f"**{BASE_FILENAME}.xlsx**（或大小寫變體，如 .XLSX）\n\n"
        "請確認：\n"
        "1. Excel 檔案與 app.py 放在同一個資料夾（或在左側輸入正確路徑）\n"
        "2. 檔名是否正確（可容許副檔名大小寫不同，但主檔名須一致）"
    )
    st.stop()
else:
    st.sidebar.success(f"✅ 已讀取檔案：\n{os.path.basename(file_path)}")

if st.sidebar.button("🔄 重新整理資料（重新讀取 Excel）"):
    st.cache_data.clear()
    st.session_state.pop("edit_target", None)
    st.session_state.pop("weekly_pdf_bytes", None)
    st.rerun()

st.sidebar.caption(f"最後讀取時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

mtime = os.path.getmtime(file_path)
all_data, sheet_names_in_file = load_data(file_path, mtime)
st.session_state["all_data"] = all_data  # 供 next_case_no() 等輔助函式讀取

missing_sheets = [g for g, s in SHEET_NAME_MAP.items() if s not in sheet_names_in_file]
if missing_sheets:
    st.sidebar.warning(
        f"⚠️ 在 Excel 中找不到工作表：{', '.join(missing_sheets)}\n\n"
        f"目前 Excel 內的工作表為：{', '.join(sheet_names_in_file)}\n\n"
        "請至 app.py 頂端 SHEET_NAME_MAP 修改對應名稱。"
    )

if "edit_target" not in st.session_state:
    st.session_state["edit_target"] = None  # (group, case_no) 或 None

if "weekly_pdf_bytes" not in st.session_state:
    st.session_state["weekly_pdf_bytes"] = None


def persist_and_refresh(group, success_msg):
    ok, err = save_group_data(file_path, group, all_data[group][REQUIRED_COLUMNS])
    if ok:
        st.success(success_msg)
        st.cache_data.clear()
        st.session_state["edit_target"] = None
        st.rerun()
    else:
        st.error(f"❌ 儲存失敗：{err}")


# ============================================================
# 共用元件：逐筆案件列表（含「✏️ 編輯狀態」按鈕）
# ============================================================
def render_case_list(df: pd.DataFrame, group_key: str):
    df = add_light_columns(df)

    col1, col2, col3 = st.columns([1, 1, 1])
    with col1:
        status_filter = st.multiselect(
            "依狀態篩選",
            options=sorted([s for s in df[COL_STATUS].dropna().unique().tolist() if s]) if not df.empty else [],
            key=f"status_filter_{group_key}",
        )
    with col2:
        owner_filter = st.multiselect(
            "依負責參謀篩選",
            options=sorted([o for o in df[COL_OWNER].dropna().unique().tolist() if o]) if not df.empty else [],
            key=f"owner_filter_{group_key}",
        )
    with col3:
        light_filter = st.multiselect(
            "依燈號篩選",
            options=["🔴 逾期", "🟡 即將到期", "🟢 已完工", "⚪ 進行中"],
            key=f"light_filter_{group_key}",
        )

    filtered = df.copy()
    if status_filter:
        filtered = filtered[filtered[COL_STATUS].isin(status_filter)]
    if owner_filter:
        filtered = filtered[filtered[COL_OWNER].isin(owner_filter)]
    if light_filter:
        light_map_rev = {
            "🔴 逾期": "red", "🟡 即將到期": "yellow",
            "🟢 已完工": "green", "⚪ 進行中": "gray",
        }
        selected_lights = [light_map_rev[x] for x in light_filter]
        filtered = filtered[filtered["燈號"].isin(selected_lights)]

    st.markdown(f"共 **{len(filtered)}** 筆案件（總數 {len(df)} 筆）")

    if filtered.empty:
        st.info("目前沒有符合篩選條件的案件。")
        return

    # 表頭
    h = st.columns([0.7, 1.3, 2.6, 1.2, 1.2, 1.2, 1.1])
    headers = ["燈號", "案號", "工作項目", "申請單位", "負責參謀", "預計完成日", "操作"]
    for c, txt in zip(h, headers):
        c.markdown(f"**{txt}**")

    for _, row in filtered.iterrows():
        case_no = row[COL_CASE_NO]
        light = row["燈號"]
        due_str = row[COL_DUE_DATE].strftime("%Y-%m-%d") if pd.notna(row[COL_DUE_DATE]) else "未填"

        row_bg = {
            "red": "#ffe3e3", "yellow": "#fff9db", "green": "#ebfbee", "gray": "transparent",
        }[light]

        with st.container(border=True):
            c = st.columns([0.7, 1.3, 2.6, 1.2, 1.2, 1.2, 1.1])
            c[0].markdown(f"### {LIGHT_EMOJI[light]}")
            c[1].markdown(f"**{case_no}**")
            c[2].markdown(row[COL_TASK])
            c[3].markdown(row[COL_APPLY_UNIT] if row[COL_APPLY_UNIT] else "—")
            c[4].markdown(row[COL_OWNER] if row[COL_OWNER] else "—")
            c[5].markdown(due_str)
            btn_label = "🔽 收合" if st.session_state["edit_target"] == (group_key, case_no) else "✏️ 編輯狀態"
            if c[6].button(btn_label, key=f"editbtn_{group_key}_{case_no}"):
                if st.session_state["edit_target"] == (group_key, case_no):
                    st.session_state["edit_target"] = None
                else:
                    st.session_state["edit_target"] = (group_key, case_no)
                st.rerun()

            if row[COL_PROGRESS]:
                st.caption(f"📋 目前進度：{row[COL_PROGRESS][:120]}{'…' if len(row[COL_PROGRESS]) > 120 else ''}")

            # 內嵌編輯面板
            if st.session_state["edit_target"] == (group_key, case_no):
                st.markdown("---")
                st.markdown(f"**編輯案件：{case_no}｜{row[COL_TASK]}**")

                full_df = all_data[group_key]
                row_idx = full_df.index[full_df[COL_CASE_NO] == case_no][0]
                current_row = full_df.loc[row_idx]

                e1, e2 = st.columns(2)
                with e1:
                    new_status = st.selectbox(
                        "狀態",
                        options=STATUS_OPTIONS,
                        index=STATUS_OPTIONS.index(current_row[COL_STATUS])
                        if current_row[COL_STATUS] in STATUS_OPTIONS else 0,
                        key=f"edit_status_{group_key}_{case_no}",
                    )
                with e2:
                    current_due = current_row[COL_DUE_DATE]
                    default_due = current_due.date() if pd.notna(current_due) else date.today()
                    new_due = st.date_input(
                        "預計完成日", value=default_due, key=f"edit_due_{group_key}_{case_no}"
                    )

                e3, e4 = st.columns(2)
                with e3:
                    mark_done_today = st.checkbox(
                        "設定「實際完成日」為今天",
                        value=(new_status == "已完成" and pd.isna(current_row[COL_ACTUAL_DATE])),
                        key=f"edit_markdone_{group_key}_{case_no}",
                    )
                with e4:
                    current_actual = current_row[COL_ACTUAL_DATE]
                    st.caption(
                        "實際完成日：" + (current_actual.strftime("%Y-%m-%d") if pd.notna(current_actual) else "尚未完成")
                    )

                new_progress = st.text_area(
                    "作業進度 / 備註",
                    value=current_row[COL_PROGRESS],
                    key=f"edit_progress_{group_key}_{case_no}",
                    height=100,
                )

                s1, s2 = st.columns([1, 1])
                with s1:
                    if st.button("💾 儲存並存回 Excel", key=f"save_{group_key}_{case_no}", type="primary"):
                        all_data[group_key].loc[row_idx, COL_STATUS] = new_status
                        all_data[group_key].loc[row_idx, COL_DUE_DATE] = pd.Timestamp(new_due)
                        all_data[group_key].loc[row_idx, COL_PROGRESS] = new_progress
                        if mark_done_today:
                            all_data[group_key].loc[row_idx, COL_ACTUAL_DATE] = pd.Timestamp(date.today())
                        elif new_status != "已完成":
                            all_data[group_key].loc[row_idx, COL_ACTUAL_DATE] = pd.NaT
                        persist_and_refresh(group_key, f"✅ 案號 {case_no} 已更新並存回 Excel！")
                with s2:
                    if st.button("取消", key=f"cancel_{group_key}_{case_no}"):
                        st.session_state["edit_target"] = None
                        st.rerun()


# ============================================================
# 共用元件：新增案件表單（可選擇組別）
# ============================================================
def render_add_case_form(default_group=None, key_prefix="global"):
    st.markdown("### ➕ 新增案件")
    with st.form(key=f"add_form_{key_prefix}", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            group_options = GROUP_NAMES
            default_idx = group_options.index(default_group) if default_group in group_options else 0
            sel_group = st.selectbox("組別", options=group_options, index=default_idx, key=f"add_group_{key_prefix}")
            new_case_no = st.text_input("案號（留白將自動編號）", key=f"add_caseno_{key_prefix}")
            new_apply_unit = st.text_input("申請單位", key=f"add_unit_{key_prefix}")
            new_task = st.text_input("工作項目", key=f"add_task_{key_prefix}")
        with c2:
            new_owner = st.text_input("承辦參謀（負責參謀）", key=f"add_owner_{key_prefix}")
            new_due_date = st.date_input("預計完成日", value=date.today(), key=f"add_due_{key_prefix}")
            new_status_add = st.selectbox("目前狀態", options=STATUS_OPTIONS, key=f"add_status_{key_prefix}")
            new_progress_add = st.text_area("作業進度 / 備註", height=80, key=f"add_progress_{key_prefix}")

        submitted = st.form_submit_button("➕ 新增案件並存入 Excel")
        if submitted:
            if not new_task:
                st.warning("⚠️ 請至少填寫「工作項目」。")
                return
            case_no_final = new_case_no.strip() if new_case_no.strip() else next_case_no(sel_group)

            if case_no_final in all_data[sel_group][COL_CASE_NO].values:
                st.warning(f"⚠️ 案號「{case_no_final}」在「{sel_group}」中已存在，請改用其他案號。")
                return

            new_row = {
                COL_CASE_NO: case_no_final,
                COL_APPLY_UNIT: new_apply_unit,
                COL_TASK: new_task,
                COL_OWNER: new_owner,
                COL_PROGRESS: new_progress_add,
                COL_DUE_DATE: pd.Timestamp(new_due_date),
                COL_ACTUAL_DATE: pd.Timestamp(date.today()) if new_status_add == "已完成" else pd.NaT,
                COL_STATUS: new_status_add,
            }
            all_data[sel_group] = pd.concat(
                [all_data[sel_group], pd.DataFrame([new_row])], ignore_index=True
            )
            persist_and_refresh(sel_group, f"✅ 新案件「{case_no_final}」已新增到「{sel_group}」並存回 Excel！")


# ============================================================
# 頁首（含右上角「匯出主管週報 PDF」按鈕）
# ============================================================
header_col1, header_col2 = st.columns([3, 1])
with header_col1:
    st.title("🏥 醫務企劃室工作進度管制")
    st.caption(f"資料來源：{os.path.basename(file_path)}　｜　今天日期：{date.today().strftime('%Y-%m-%d')}")
with header_col2:
    st.write("")  # 垂直對齊用的間距
    if st.button("📄 匯出主管週報 PDF", use_container_width=True, type="primary"):
        with st.spinner("正在產出 PDF 報告..."):
            st.session_state["weekly_pdf_bytes"] = generate_weekly_report_pdf(all_data)
    if st.session_state.get("weekly_pdf_bytes"):
        st.download_button(
            "⬇️ 下載週報 PDF",
            data=st.session_state["weekly_pdf_bytes"],
            file_name=f"醫務企劃室主管週報_{date.today().strftime('%Y%m%d')}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

tab_overview, tab_medical, tab_admin, tab_it = st.tabs(
    ["📊 全室總覽", "🩺 醫勤組", "🗂️ 行政組", "💻 通資電組"]
)

# ============================================================
# Tab 1：全室總覽
# ============================================================
with tab_overview:
    all_df = pd.concat(
        [df.assign(**{"組別": g}) for g, df in all_data.items()],
        ignore_index=True,
    ) if all_data else pd.DataFrame()

    all_df_light = add_light_columns(all_df)

    total_count = len(all_df_light)
    done_count = int((all_df_light["燈號"] == "green").sum()) if total_count else 0
    red_count = int((all_df_light["燈號"] == "red").sum()) if total_count else 0
    completion_rate = (done_count / total_count * 100) if total_count else 0

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("總案件數", f"{total_count} 件")
    k2.metric("已完工數", f"{done_count} 件")
    k3.metric("整體完工率", f"{completion_rate:.1f}%")
    k4.metric("🔴 逾期紅燈案件數", f"{red_count} 件", delta=None)

    st.divider()
    st.subheader("⚠️ 逾期案件優先清單")

    overdue_df = all_df_light[all_df_light["燈號"] == "red"].copy()
    if overdue_df.empty:
        st.success("🎉 目前沒有逾期案件！")
    else:
        overdue_df = overdue_df.sort_values(COL_DUE_DATE)
        overdue_df["逾期天數"] = (pd.Timestamp(date.today()) - overdue_df[COL_DUE_DATE]).dt.days
        show_cols = ["組別", COL_CASE_NO, COL_APPLY_UNIT, COL_TASK, COL_OWNER, COL_DUE_DATE, COL_STATUS, "逾期天數"]
        show_overdue = overdue_df[show_cols].copy()
        show_overdue[COL_DUE_DATE] = show_overdue[COL_DUE_DATE].dt.strftime("%Y-%m-%d")
        st.dataframe(
            show_overdue.style.apply(lambda r: ["background-color:#ffe3e3; color:#c81e1e;"] * len(r), axis=1),
            use_container_width=True,
            hide_index=True,
        )
        st.caption("💡 提示：可至各組分頁點選案件旁的「✏️ 編輯狀態」直接更新進度。")

    st.divider()
    st.subheader("各組完工概況")
    group_summary = []
    for g, df in all_data.items():
        df_l = add_light_columns(df)
        n = len(df_l)
        d = int((df_l["燈號"] == "green").sum()) if n else 0
        r = int((df_l["燈號"] == "red").sum()) if n else 0
        y = int((df_l["燈號"] == "yellow").sum()) if n else 0
        group_summary.append({
            "組別": g, "總案件數": n, "已完工": d,
            "🔴逾期": r, "🟡即將到期": y,
            "完工率(%)": round(d / n * 100, 1) if n else 0.0,
        })
    st.dataframe(pd.DataFrame(group_summary), use_container_width=True, hide_index=True)

    st.divider()
    render_add_case_form(default_group=None, key_prefix="overview")


# ============================================================
# Tab 2-4：各組分頁
# ============================================================
def render_group_tab(group_key: str):
    df = all_data.get(group_key, pd.DataFrame(columns=REQUIRED_COLUMNS))
    st.subheader(f"{group_key} 案件清單")
    render_case_list(df, group_key)
    st.divider()
    render_add_case_form(default_group=group_key, key_prefix=group_key)


with tab_medical:
    render_group_tab("醫勤組")

with tab_admin:
    render_group_tab("行政組")

with tab_it:
    render_group_tab("通資電組")
