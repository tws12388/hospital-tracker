# -*- coding: utf-8 -*-
"""
醫務企劃室工作進度管制網頁（雲端版）
============================
技術架構：Streamlit + pandas + streamlit-gsheets（Google Sheets 雲端資料庫）+ reportlab

資料來源：Google試算表，需在 .streamlit/secrets.toml 設定好 [connections.gsheets]
          （服務帳號金鑰 + 試算表網址），詳見「使用說明.md」。
          試算表內須包含三個工作表(sheet)：醫勤組 / 行政組 / 通資組
          （畫面上分頁名稱顯示為「通資電組」，對應到 Google Sheet 的「通資組」分頁）

每個工作表欄位：
    案號 / 申請單位 / 工作項目 / 負責參謀 / 作業進度 / 預計完成日 / 實際完成日 / 目前狀態 / 存證照片

本次修訂重點：
    1. 修復 StreamlitDuplicateElementKey：讀取資料時先過濾「案號」為空/nan 的無效資料列；
       所有逐筆動態元件 key 一律加上該筆資料在 DataFrame 中的原始索引 idx，確保全域唯一。
    2. 資料庫改為 Google Sheets（streamlit-gsheets／GSheetsConnection），三組各自獨立工作表，
       全室總覽自動合併三組資料；新增/編輯會精準寫回對應組別的工作表。
    3. 保留紅黃綠燈邏輯、手機卡片模式、st.camera_input 拍照存證、
       頂部「匯出 3 組別 Excel 報表」與「產製週報 PDF」功能。
"""

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

try:
    from streamlit_gsheets import GSheetsConnection
    GSHEETS_LIB_AVAILABLE = True
except ImportError:
    GSheetsConnection = None
    GSHEETS_LIB_AVAILABLE = False


# ============================================================
# 基本設定
# ============================================================
st.set_page_config(
    page_title="醫務企劃室工作進度管制",
    page_icon="🏥",
    layout="wide",
)

APP_DIR = os.path.dirname(os.path.abspath(__file__))

# 三組對應的 Google Sheet 工作表名稱：key 為畫面顯示的組名，value 為試算表內實際分頁名稱
SHEET_NAME_MAP = {
    "醫勤組": "醫勤組",
    "行政組": "行政組",
    "通資電組": "通資組",   # 試算表內實際分頁名稱為「通資組」
}
GROUP_NAMES = list(SHEET_NAME_MAP.keys())

# 欄位名稱
COL_CASE_NO = "案號"
COL_APPLY_UNIT = "申請單位"
COL_TASK = "工作項目"
COL_OWNER = "負責參謀"
COL_PROGRESS = "作業進度"
COL_DUE_DATE = "預計完成日"
COL_ACTUAL_DATE = "實際完成日"
COL_STATUS = "目前狀態"
COL_PHOTO_NOTE = "存證照片"

REQUIRED_COLUMNS = [
    COL_CASE_NO, COL_APPLY_UNIT, COL_TASK, COL_OWNER,
    COL_PROGRESS, COL_DUE_DATE, COL_ACTUAL_DATE, COL_STATUS, COL_PHOTO_NOTE,
]

DONE_KEYWORDS = ["已完成", "完工", "結案", "已結案", "完成"]
STATUS_OPTIONS = ["辦理中", "已完成", "已逾期", "暫緩", "其他"]

# 案號視為「無效/空白」的字樣（不分大小寫），讀取時會被過濾掉
INVALID_CASE_NO_TOKENS = {"", "nan", "none", "nat", "null", "na"}

# 本地暫存拍照存證的資料夾（雲端環境重啟後可能不保留，僅作單次工作階段的存證用途）
PHOTO_DIR = os.path.join(APP_DIR, "case_photos")


# ============================================================
# 中文字型設定（PDF 用）
# 優先使用 Windows 內建「微軟正黑體」，找不到才依序退回標楷體、內附文泉驛正黑，
# 最終仍找不到則安全退回 reportlab 內建字型，確保絕不讓程式崩潰。
# ============================================================
FONT_PATH = os.path.join(APP_DIR, "fonts", "WenQuanYiZenHei.ttf")
FONT_NAME = "CJKFont"
FONT_NAME_BOLD = "CJKFont-Bold"

FONT_CANDIDATES = [
    {
        "label": "Windows 微軟正黑體",
        "regular": ["C:/Windows/Fonts/msjh.ttc", "C:/Windows/Fonts/msjh.ttf"],
        "bold": ["C:/Windows/Fonts/msjhbd.ttc", "C:/Windows/Fonts/msjhbd.ttf"],
    },
    {
        "label": "Windows 標楷體",
        "regular": ["C:/Windows/Fonts/kaiu.ttf"],
        "bold": ["C:/Windows/Fonts/kaiu.ttf"],
    },
    {
        "label": "程式內附文泉驛正黑",
        "regular": [FONT_PATH],
        "bold": [FONT_PATH],
    },
]


def register_pdf_fonts():
    """依優先順序註冊可顯示中文的 PDF 字型；全部失敗則安全退回內建字型，絕不崩潰。"""
    cached = st.session_state.get("_pdf_font_status")
    if cached is not None:
        return cached

    global FONT_NAME, FONT_NAME_BOLD

    for candidate in FONT_CANDIDATES:
        reg_path = next((p for p in candidate["regular"] if os.path.exists(p)), None)
        if not reg_path:
            continue
        try:
            pdfmetrics.registerFont(TTFont(FONT_NAME, reg_path))
        except Exception:
            continue

        bold_path = next((p for p in candidate["bold"] if os.path.exists(p)), None)
        try:
            if bold_path:
                pdfmetrics.registerFont(TTFont(FONT_NAME_BOLD, bold_path))
            else:
                raise FileNotFoundError("無獨立粗體字型檔")
        except Exception:
            pdfmetrics.registerFont(TTFont(FONT_NAME_BOLD, reg_path))

        status = (True, candidate["label"])
        st.session_state["_pdf_font_status"] = status
        return status

    FONT_NAME = "Helvetica"
    FONT_NAME_BOLD = "Helvetica-Bold"
    status = (False, "找不到可用的中文字型（已嘗試微軟正黑體、標楷體、內附字型），PDF 中的中文可能無法正確顯示")
    st.session_state["_pdf_font_status"] = status
    return status


# ============================================================
# Google Sheets 連線與資料讀寫
# ============================================================
@st.cache_resource(show_spinner=False)
def get_gsheets_connection():
    """建立並快取 GSheetsConnection；若套件未安裝或設定有誤，交由呼叫端捕捉例外。"""
    return st.connection("gsheets", type=GSheetsConnection)


def clean_group_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    統一補齊欄位、修正型別，並【過濾掉案號為空白/nan/無效值的資料列】，
    避免這些髒資料在畫面上產生重複或空白的元件 key。
    """
    df = df.copy() if df is not None else pd.DataFrame()

    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            df[col] = None

    df[COL_DUE_DATE] = pd.to_datetime(df[COL_DUE_DATE], errors="coerce")
    df[COL_ACTUAL_DATE] = pd.to_datetime(df[COL_ACTUAL_DATE], errors="coerce")
    df[COL_STATUS] = df[COL_STATUS].fillna("").astype(str).replace("nan", "")
    df[COL_PROGRESS] = df[COL_PROGRESS].fillna("").astype(str).replace("nan", "")
    df[COL_OWNER] = df[COL_OWNER].fillna("").astype(str).replace("nan", "")
    df[COL_APPLY_UNIT] = df[COL_APPLY_UNIT].fillna("").astype(str).replace("nan", "")
    df[COL_PHOTO_NOTE] = df[COL_PHOTO_NOTE].fillna("").astype(str).replace("nan", "")

    # ---- 核心修復：過濾案號為空白、nan 或無效值的資料列 ----
    df[COL_CASE_NO] = df[COL_CASE_NO].astype(str).str.strip()
    df = df[~df[COL_CASE_NO].str.lower().isin(INVALID_CASE_NO_TOKENS)]

    df = df.reset_index(drop=True)
    return df[REQUIRED_COLUMNS]


def load_all_data(conn):
    """從 Google Sheets 讀取三組工作表，回傳 (data: dict[組名]=DataFrame, errors: dict[組名]=錯誤訊息)"""
    data, errors = {}, {}
    for group, worksheet_name in SHEET_NAME_MAP.items():
        try:
            raw_df = conn.read(worksheet=worksheet_name, ttl=60)
            data[group] = clean_group_df(raw_df)
        except Exception as e:
            data[group] = pd.DataFrame(columns=REQUIRED_COLUMNS)
            errors[group] = str(e)
    return data, errors


def save_group_data(conn, group, df):
    """把單一組別的 DataFrame 寫回 Google Sheets 對應分頁（只覆寫該分頁，其餘分頁不受影響）"""
    worksheet_name = SHEET_NAME_MAP[group]
    try:
        out_df = df[REQUIRED_COLUMNS].copy()
        out_df[COL_DUE_DATE] = out_df[COL_DUE_DATE].apply(
            lambda d: d.strftime("%Y-%m-%d") if pd.notna(d) else ""
        )
        out_df[COL_ACTUAL_DATE] = out_df[COL_ACTUAL_DATE].apply(
            lambda d: d.strftime("%Y-%m-%d") if pd.notna(d) else ""
        )
        conn.update(worksheet=worksheet_name, data=out_df)
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
    """回傳 'red'(逾期) / 'yellow'(即將到期) / 'green'(已完工) / 'gray'(進行中)"""
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


def get_this_week_range(ref_date=None):
    ref_date = ref_date or date.today()
    monday = ref_date - timedelta(days=ref_date.weekday())
    sunday = monday + timedelta(days=6)
    return monday, sunday


# ============================================================
# 匯出：3 組別 Excel 報表
# ============================================================
def generate_excel_export(all_data: dict) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for group, worksheet_name in SHEET_NAME_MAP.items():
            df = all_data.get(group, pd.DataFrame(columns=REQUIRED_COLUMNS))
            export_df = df[REQUIRED_COLUMNS].copy() if not df.empty else pd.DataFrame(columns=REQUIRED_COLUMNS)
            if not export_df.empty:
                export_df[COL_DUE_DATE] = pd.to_datetime(export_df[COL_DUE_DATE], errors="coerce").dt.strftime("%Y-%m-%d")
                export_df[COL_ACTUAL_DATE] = pd.to_datetime(export_df[COL_ACTUAL_DATE], errors="coerce").dt.strftime("%Y-%m-%d")
            export_df.to_excel(writer, sheet_name=worksheet_name, index=False)
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
# 產生「主管週報 PDF」
# ============================================================
def generate_weekly_report_pdf(all_data: dict):
    """回傳 (pdf_bytes, font_ok, font_label)。font_ok=False 代表退回無中文字形的安全字型。"""
    font_ok, font_label = register_pdf_fonts()

    week_start, week_end = get_this_week_range()
    today_str = date.today().strftime("%Y-%m-%d")
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")

    all_df = pd.concat(
        [df.assign(**{"組別": g}) for g, df in all_data.items()],
        ignore_index=True,
    ) if all_data else pd.DataFrame(columns=REQUIRED_COLUMNS + ["組別"])
    all_df_light = add_light_columns(all_df)

    total_count = len(all_df_light)
    done_count = int((all_df_light["燈號"] == "green").sum()) if total_count else 0
    red_count = int((all_df_light["燈號"] == "red").sum()) if total_count else 0
    ongoing_count = total_count - done_count - red_count
    completion_rate = (done_count / total_count * 100) if total_count else 0.0

    group_rows = []
    for g in GROUP_NAMES:
        df = all_data.get(g, pd.DataFrame(columns=REQUIRED_COLUMNS))
        df_l = add_light_columns(df)
        n = len(df_l)
        d = int((df_l["燈號"] == "green").sum()) if n else 0
        r = int((df_l["燈號"] == "red").sum()) if n else 0
        o = n - d - r
        rate = round(d / n * 100, 1) if n else 0.0
        group_rows.append({"組別": g, "總件數": n, "已結案": d, "進行中": o, "逾期": r, "完工率": rate})

    ranked = sorted(group_rows, key=lambda x: (x["完工率"], -x["逾期"]))
    worst = ranked[0] if ranked else None
    if worst and (worst["逾期"] > 0 or worst["完工率"] < 50) and any(g["總件數"] > 0 for g in group_rows):
        insight_text = (
            f"【提醒】本週推動落後組別為「{worst['組別']}」，完工率 {worst['完工率']}%，"
            f"目前有 {worst['逾期']} 件時效落後案件，建議優先追蹤與協調資源。"
        )
    else:
        insight_text = "【良好】各組推動進度大致穩定，目前無明顯落後組別。"

    overdue_df = all_df_light[all_df_light["燈號"] == "red"].copy()
    if not overdue_df.empty:
        overdue_df["逾期天數"] = (pd.Timestamp(date.today()) - overdue_df[COL_DUE_DATE]).dt.days
        overdue_df = overdue_df.sort_values("逾期天數", ascending=False)

    org_style = ParagraphStyle(
        "OrgTC", fontName=FONT_NAME, fontSize=10.5, alignment=TA_CENTER,
        textColor=colors.HexColor("#5b6b7a"),
    )
    title_style = ParagraphStyle(
        "TitleTC", fontName=FONT_NAME_BOLD, fontSize=21, leading=27, alignment=TA_CENTER,
        textColor=colors.HexColor("#12293f"), spaceAfter=6,
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
        "KpiValue", fontName=FONT_NAME_BOLD, fontSize=19, alignment=TA_CENTER, textColor=colors.white,
    )
    sig_label_style = ParagraphStyle(
        "SigLabel", fontName=FONT_NAME_BOLD, fontSize=10.5, alignment=TA_CENTER,
        textColor=colors.HexColor("#12293f"),
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
        title="醫務企劃室工作進度管制主管週報",
    )
    doc_width = doc.width

    elements = []
    elements.append(Paragraph("醫務企劃室", org_style))
    elements.append(Spacer(1, 2))
    elements.append(Paragraph("工作進度管制　主管週報", title_style))
    top_rule = Table([[""]], colWidths=[doc_width], rowHeights=[1.6])
    top_rule.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#12293f"))]))
    elements.append(top_rule)
    elements.append(Spacer(1, 6))
    elements.append(Paragraph(
        f"本週期間：{week_start.strftime('%Y-%m-%d')} ～ {week_end.strftime('%Y-%m-%d')}　"
        f"｜　報表產出時間：{generated_at}",
        subtitle_style,
    ))
    elements.append(Spacer(1, 10))

    elements.append(section_bar("全室列管綜效"))
    elements.append(Spacer(1, 6))
    kpi_defs = [
        ("總數", f"{total_count} 件", "#1c3d5a"),
        ("已結案", f"{done_count} 件", "#2b8a3e"),
        ("進行中", f"{ongoing_count} 件", "#6c757d"),
        ("逾期數", f"{red_count} 件", "#c81e1e"),
        ("完工率", f"{completion_rate:.1f}%", "#0b6e99"),
    ]
    kpi_cell_w = doc_width / 5
    label_row = [Paragraph(l, kpi_label_style) for l, _, _ in kpi_defs]
    value_row = [Paragraph(v, kpi_value_style) for _, v, _ in kpi_defs]
    kpi_table = Table([label_row, value_row], colWidths=[kpi_cell_w] * 5)
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

    elements.append(section_bar("醫勤組／行政組／通資電組　三組推動分析"))
    elements.append(Spacer(1, 6))
    group_header = ["組別", "總件數", "已結案", "進行中", "逾期", "完工率"]
    group_table_data = [group_header]
    for g in group_rows:
        group_table_data.append([
            g["組別"], str(g["總件數"]), str(g["已結案"]),
            str(g["進行中"]), str(g["逾期"]), f"{g['完工率']}%",
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

    elements.append(section_bar(f"紅燈時效落後案件明細清單（共 {len(overdue_df)} 件）"))
    elements.append(Spacer(1, 6))
    if overdue_df.empty:
        elements.append(Paragraph("目前無時效落後案件，全室進度良好。", body_style))
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
            ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#ffe3e3")),
            ("TEXTCOLOR", (0, 1), (-1, -1), colors.HexColor("#c81e1e")),
            ("FONTNAME", (5, 1), (5, -1), FONT_NAME_BOLD),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        elements.append(od_table)

    elements.append(Spacer(1, 22))
    elements.append(section_bar("核簽"))
    elements.append(Spacer(1, 6))
    sig_header = [
        Paragraph("承　辦", sig_label_style), Paragraph("組　長", sig_label_style), Paragraph("主　任", sig_label_style),
    ]
    sig_table = Table([sig_header, ["", "", ""]], colWidths=[doc_width / 3] * 3, rowHeights=[20, 52])
    sig_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e8edf2")),
        ("GRID", (0, 0), (-1, -1), 0.75, colors.HexColor("#8a94a3")),
        ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, 0), 4),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 4),
    ]))
    elements.append(sig_table)

    def footer(canvas, doc_):
        canvas.saveState()
        canvas.setFont(FONT_NAME, 8)
        canvas.setFillColor(colors.HexColor("#9aa5b1"))
        canvas.drawString(16 * mm, 10 * mm, f"醫務企劃室工作進度管制系統　自動產出　{today_str}")
        canvas.drawRightString(A4[0] - 16 * mm, 10 * mm, f"第 {doc_.page} 頁")
        canvas.restoreState()

    doc.build(elements, onFirstPage=footer, onLaterPages=footer)
    buffer.seek(0)
    return buffer.getvalue(), font_ok, font_label


# ============================================================
# 側邊欄：連線狀態 / 手機卡片模式 / 重新整理
# ============================================================
st.sidebar.title("⚙️ 系統設定")

mobile_mode = st.sidebar.toggle("📱 手機卡片模式", value=False, key="mobile_mode_toggle")
st.sidebar.caption("在手機瀏覽器開啟時，建議開啟卡片模式以獲得較佳的觸控體驗。")

if st.sidebar.button("🔄 重新整理資料（重新讀取 Google Sheet）"):
    st.cache_data.clear()
    st.session_state.pop("edit_target", None)
    st.session_state.pop("_pdf_font_status", None)
    st.rerun()

st.sidebar.caption(f"最後讀取時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

# ---- 建立 Google Sheets 連線（含安全容錯，設定有誤也不會讓整個系統崩潰）----
gsheets_ready = False
conn = None
connection_error = None

if not GSHEETS_LIB_AVAILABLE:
    connection_error = (
        "尚未安裝 streamlit-gsheets 套件。請在終端機執行：\n\n"
        "    pip install streamlit-gsheets\n\n"
        "安裝後請重新啟動程式。"
    )
else:
    try:
        conn = get_gsheets_connection()
        gsheets_ready = True
    except Exception as e:
        connection_error = (
            f"Google Sheets 連線建立失敗：{e}\n\n"
            "請確認 .streamlit/secrets.toml 是否已正確設定 [connections.gsheets]，"
            "詳見「使用說明.md」。"
        )

if not gsheets_ready:
    st.sidebar.error("⚠️ Google Sheets 尚未連線成功")
    st.error(f"❌ 無法連線至 Google Sheets 雲端資料庫：\n\n{connection_error}")
    # 安全容錯：連線失敗時仍以空白資料結構繼續往下執行，避免整個系統直接崩潰白屏
    all_data = {g: pd.DataFrame(columns=REQUIRED_COLUMNS) for g in GROUP_NAMES}
    load_errors = {}
else:
    st.sidebar.success("✅ 已連線至 Google Sheets")
    all_data, load_errors = load_all_data(conn)
    st.session_state["all_data"] = all_data
    for g, err in load_errors.items():
        st.sidebar.warning(f"⚠️ 讀取「{g}」工作表時發生問題：{err}")

if "edit_target" not in st.session_state:
    st.session_state["edit_target"] = None  # (group, case_no, idx) 或 None


def persist_and_refresh(group, success_msg):
    if not gsheets_ready:
        st.error("❌ 尚未連線至 Google Sheets，無法儲存。")
        return
    ok, err = save_group_data(conn, group, all_data[group][REQUIRED_COLUMNS])
    if ok:
        st.success(success_msg)
        st.cache_data.clear()
        st.session_state["edit_target"] = None
        st.rerun()
    else:
        st.error(f"❌ 儲存失敗：{err}")


# ============================================================
# 共用元件：逐筆案件列表（桌面表格模式 / 手機卡片模式），含「編輯狀態」按鈕
# 註：所有動態元件 key 皆以「組別 + 案號 + 原始資料列索引 idx」組成，確保全域唯一，
#     徹底修復 StreamlitDuplicateElementKey。
# ============================================================
def render_case_list(df: pd.DataFrame, group_key: str, mobile: bool = False):
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

    if not mobile:
        h = st.columns([0.7, 1.3, 2.6, 1.2, 1.2, 1.2, 1.1])
        headers = ["燈號", "案號", "工作項目", "申請單位", "負責參謀", "預計完成日", "操作"]
        for c, txt in zip(h, headers):
            c.markdown(f"**{txt}**")

    # 用 filtered 的原始索引 idx（對應 all_data[group_key] 中的資料列位置）作為唯一鍵的一部分，
    # 並直接用它定位要編輯的資料列，比用案號比對更穩定（案號重複或空白也不會出錯）。
    for idx, row in filtered.iterrows():
        case_no = row[COL_CASE_NO]
        light = row["燈號"]
        due_str = row[COL_DUE_DATE].strftime("%Y-%m-%d") if pd.notna(row[COL_DUE_DATE]) else "未填"
        edit_key = (group_key, case_no, idx)
        is_editing = st.session_state["edit_target"] == edit_key

        with st.container(border=True):
            if mobile:
                # ---- 手機卡片模式：直向堆疊，較大觸控區塊 ----
                st.markdown(f"### {LIGHT_EMOJI[light]} {case_no}　{row[COL_TASK]}")
                st.caption(
                    f"申請單位：{row[COL_APPLY_UNIT] or '—'}　"
                    f"｜　負責參謀：{row[COL_OWNER] or '—'}　"
                    f"｜　預計完成日：{due_str}"
                )
                btn_label = "🔽 收合" if is_editing else "✏️ 編輯狀態"
                if st.button(btn_label, key=f"editbtn_{group_key}_{case_no}_{idx}", use_container_width=True):
                    st.session_state["edit_target"] = None if is_editing else edit_key
                    st.rerun()
            else:
                # ---- 桌面表格模式 ----
                c = st.columns([0.7, 1.3, 2.6, 1.2, 1.2, 1.2, 1.1])
                c[0].markdown(f"### {LIGHT_EMOJI[light]}")
                c[1].markdown(f"**{case_no}**")
                c[2].markdown(row[COL_TASK])
                c[3].markdown(row[COL_APPLY_UNIT] if row[COL_APPLY_UNIT] else "—")
                c[4].markdown(row[COL_OWNER] if row[COL_OWNER] else "—")
                c[5].markdown(due_str)
                btn_label = "🔽 收合" if is_editing else "✏️ 編輯狀態"
                if c[6].button(btn_label, key=f"editbtn_{group_key}_{case_no}_{idx}"):
                    st.session_state["edit_target"] = None if is_editing else edit_key
                    st.rerun()

            if row[COL_PROGRESS]:
                st.caption(f"📋 目前進度：{row[COL_PROGRESS][:120]}{'…' if len(row[COL_PROGRESS]) > 120 else ''}")
            if row[COL_PHOTO_NOTE]:
                st.caption(f"📷 存證紀錄：{row[COL_PHOTO_NOTE]}")

            # ---- 內嵌編輯面板 ----
            if is_editing:
                st.markdown("---")
                st.markdown(f"**編輯案件：{case_no}｜{row[COL_TASK]}**")

                full_df = all_data[group_key]
                # 直接用 idx（原始資料列索引）定位，不再用案號比對，避免案號重複/空白造成錯誤
                row_idx = idx
                current_row = full_df.loc[row_idx]

                e1, e2 = st.columns(2)
                with e1:
                    new_status = st.selectbox(
                        "狀態",
                        options=STATUS_OPTIONS,
                        index=STATUS_OPTIONS.index(current_row[COL_STATUS])
                        if current_row[COL_STATUS] in STATUS_OPTIONS else 0,
                        key=f"edit_status_{group_key}_{case_no}_{idx}",
                    )
                with e2:
                    current_due = current_row[COL_DUE_DATE]
                    default_due = current_due.date() if pd.notna(current_due) else date.today()
                    new_due = st.date_input(
                        "預計完成日", value=default_due, key=f"edit_due_{group_key}_{case_no}_{idx}"
                    )

                e3, e4 = st.columns(2)
                with e3:
                    mark_done_today = st.checkbox(
                        "設定「實際完成日」為今天",
                        value=(new_status == "已完成" and pd.isna(current_row[COL_ACTUAL_DATE])),
                        key=f"edit_markdone_{group_key}_{case_no}_{idx}",
                    )
                with e4:
                    current_actual = current_row[COL_ACTUAL_DATE]
                    st.caption(
                        "實際完成日：" + (current_actual.strftime("%Y-%m-%d") if pd.notna(current_actual) else "尚未完成")
                    )

                new_progress = st.text_area(
                    "作業進度 / 備註",
                    value=current_row[COL_PROGRESS],
                    key=f"edit_progress_{group_key}_{case_no}_{idx}",
                    height=100,
                )

                # ---- 拍照存證 ----
                photo = st.camera_input(
                    "📷 拍照存證（設備稽核 / 驗收照片，可略過）",
                    key=f"camera_{group_key}_{case_no}_{idx}",
                )
                photo_note_update = None
                if photo is not None:
                    st.image(photo, caption="拍攝預覽", use_container_width=True)
                    ts = datetime.now().strftime("%Y-%m-%d %H:%M")
                    photo_note_update = f"已拍照存證：{ts}"
                    try:
                        os.makedirs(PHOTO_DIR, exist_ok=True)
                        safe_case_no = "".join(c for c in str(case_no) if c.isalnum()) or "case"
                        fname = f"{group_key}_{safe_case_no}_{datetime.now().strftime('%Y%m%d%H%M%S')}.jpg"
                        with open(os.path.join(PHOTO_DIR, fname), "wb") as f:
                            f.write(photo.getbuffer())
                    except Exception:
                        # 本地暫存失敗（例如雲端環境檔案系統唯讀）不影響其餘功能，僅存文字紀錄
                        pass

                s1, s2 = st.columns([1, 1])
                with s1:
                    if st.button(
                        "💾 儲存並存回 Google Sheet",
                        key=f"save_{group_key}_{case_no}_{idx}",
                        type="primary",
                    ):
                        all_data[group_key].loc[row_idx, COL_STATUS] = new_status
                        all_data[group_key].loc[row_idx, COL_DUE_DATE] = pd.Timestamp(new_due)
                        all_data[group_key].loc[row_idx, COL_PROGRESS] = new_progress
                        if photo_note_update:
                            all_data[group_key].loc[row_idx, COL_PHOTO_NOTE] = photo_note_update
                        if mark_done_today:
                            all_data[group_key].loc[row_idx, COL_ACTUAL_DATE] = pd.Timestamp(date.today())
                        elif new_status != "已完成":
                            all_data[group_key].loc[row_idx, COL_ACTUAL_DATE] = pd.NaT
                        persist_and_refresh(group_key, f"✅ 案號 {case_no} 已更新並存回 Google Sheet！")
                with s2:
                    if st.button("取消", key=f"cancel_{group_key}_{case_no}_{idx}"):
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

        submitted = st.form_submit_button("➕ 新增案件並存回 Google Sheet")
        if submitted:
            if not new_task:
                st.warning("⚠️ 請至少填寫「工作項目」。")
                return
            case_no_final = new_case_no.strip() if new_case_no.strip() else next_case_no(sel_group)

            if case_no_final.lower() in INVALID_CASE_NO_TOKENS:
                st.warning("⚠️ 案號不可為空白，請重新輸入。")
                return
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
                COL_PHOTO_NOTE: "",
            }
            all_data[sel_group] = pd.concat(
                [all_data[sel_group], pd.DataFrame([new_row])], ignore_index=True
            )
            persist_and_refresh(sel_group, f"✅ 新案件「{case_no_final}」已新增到「{sel_group}」並存回 Google Sheet！")


# ============================================================
# 頁首：標題 + 頂部「匯出3組別Excel報表」與「產製週報PDF」
# ============================================================
st.title("🏥 醫務企劃室工作進度管制")
st.caption(f"資料來源：Google Sheets 雲端資料庫　｜　今天日期：{date.today().strftime('%Y-%m-%d')}")

export_col1, export_col2 = st.columns(2)
with export_col1:
    excel_bytes = generate_excel_export(all_data)
    st.download_button(
        "📥 匯出 3 組別 Excel 報表",
        data=excel_bytes,
        file_name=f"醫務企劃室工作進度_{date.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with export_col2:
    try:
        pdf_bytes, pdf_font_ok, pdf_font_label = generate_weekly_report_pdf(all_data)
    except Exception as e:
        pdf_bytes = None
        pdf_font_ok = False
        pdf_font_label = f"PDF 模組載入異常（但不影響系統運作）: {e}"
    st.download_button(
        "📄 產製週報 PDF",
        data=pdf_bytes,
        file_name=f"醫務企劃室主管週報_{date.today().strftime('%Y%m%d')}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )
    if not pdf_font_ok:
        st.caption(f"⚠️ {pdf_font_label}")

tab_overview, tab_medical, tab_admin, tab_it = st.tabs(
    ["📊 全室總覽", "🩺 醫勤組", "🗂️ 行政組", "💻 通資電組"]
)

# ============================================================
# Tab 1：全室總覽
# ============================================================
with tab_overview:
    st.subheader("📊 全室總覽")

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
    render_case_list(df, group_key, mobile=mobile_mode)
    st.divider()
    render_add_case_form(default_group=group_key, key_prefix=group_key)


with tab_medical:
    render_group_tab("醫勤組")

with tab_admin:
    render_group_tab("行政組")

with tab_it:
    render_group_tab("通資電組")
