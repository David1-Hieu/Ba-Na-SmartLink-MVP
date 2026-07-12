from __future__ import annotations

import os
from pathlib import Path
import pandas as pd
import streamlit as st
import altair as alt

from database import (
    DB_PATH,
    clear_validation_logs,
    fetch_reports,
    fetch_tasks,
    fetch_validation_logs,
    fetch_villages,
    init_db,
    list_periods,
    seed_demo_data,
    upsert_reports,
    insert_validation_logs,
)
from report_generator import build_missing_fields_dataframe, build_summary_dataframe, generate_excel_report, generate_word_report
from schema import FIELD_LABELS, INDICATOR_FIELDS
from validators import get_raw_preview, has_blockers, parse_excel_to_reports, validate_reports, normalize_text

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "templates" / "mau_bao_cao_xa_gui_thon.xlsx"
DRIVE_SAMPLE_DIR = BASE_DIR / "sample_data" / "drive_imported"

st.set_page_config(page_title="Cổng dữ liệu UBND xã Bà Nà", layout="wide")

# Unified Premium Dark Theme CSS
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

/* Global Font and Light Background Override */
html, body, [class*="css"], .stApp {
    font-family: 'Inter', sans-serif;
    background-color: #f1f5f9 !important; /* Light Gray page background */
    color: #1e293b !important; /* Dark Slate text color */
}

/* Ensure all markdown paragraphs, lists and headings use dark color on light bg */
div[data-testid="stMarkdownContainer"] p,
div[data-testid="stMarkdownContainer"] span,
div[data-testid="stMarkdownContainer"] li {
    color: #334155 !important;
}

h1, h2, h3, h4, h5, h6 {
    color: #0f172a !important;
}

.block-container {
    padding-top: 1.5rem;
    padding-bottom: 2rem;
}

/* Sidebar Styling - Solid white background with gray border */
[data-testid="stSidebar"] {
    background-color: #ffffff !important;
    border-right: 1px solid #cbd5e1 !important;
}

[data-testid="stSidebar"] .stText, 
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] h2 {
    color: #334155 !important;
}

/* Sidebar navigation radio items */
div[data-testid="stSidebarUserContent"] .stRadio > div {
    gap: 6px;
}

div[data-testid="stSidebarUserContent"] .stRadio label {
    padding: 10px 14px;
    border-radius: 10px;
    transition: all 0.2s ease;
    font-weight: 500;
    cursor: pointer;
    border: 1px solid transparent;
    color: #475569 !important;
}

div[data-testid="stSidebarUserContent"] .stRadio label:hover {
    background-color: #f1f5f9 !important;
    color: #b91c1c !important; /* Crimson hover text */
}

/* Light Theme Card containers */
div[data-testid="stVerticalBlock"] > div[style*="border: 1px solid"] {
    border-radius: 12px !important;
    border: 1px solid #e2e8f0 !important;
    border-left: 5px solid #b91c1c !important; /* Crimson left border indicator */
    background-color: #ffffff !important; /* White card background */
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.04) !important;
    padding: 1.5rem !important;
}

/* Input fields (selectboxes, inputs, buttons) styled for light mode */
.stSelectbox div[data-baseweb="select"] {
    background-color: #ffffff !important;
    border: 1px solid #cbd5e1 !important;
    color: #1e293b !important;
    border-radius: 8px !important;
}

.stSelectbox div[data-baseweb="select"] div {
    color: #1e293b !important;
}

div[role="listbox"] {
    background-color: #ffffff !important;
    color: #1e293b !important;
}

.stButton>button {
    border-radius: 8px !important;
    font-weight: 600 !important;
    transition: all 0.2s ease !important;
    background-color: #ffffff !important;
    border: 1px solid #cbd5e1 !important;
    color: #1e293b !important;
}

.stButton>button:hover {
    background-color: #f8fafc !important;
    border-color: #b91c1c !important; /* Crimson border on hover */
    color: #b91c1c !important;
}

.stButton>button[kind="primary"] {
    background: linear-gradient(135deg, #b91c1c 0%, #7f1d1d 100%) !important; /* Crimson button gradient */
    color: white !important;
    border: 1px solid #d97706 !important; /* Gold border */
}

.stButton>button[kind="primary"]:hover {
    box-shadow: 0 4px 12px rgba(185, 28, 28, 0.2) !important;
    transform: translateY(-1px);
}

div[data-testid="stFileUploaderDropzone"] {
    background-color: #f8fafc !important;
    border: 1px dashed #cbd5e1 !important;
    border-radius: 10px !important;
}

/* Style the horizontal radio group container (the navbar wrapper) */
div[data-testid="stRadio"] > div[role="radiogroup"] {
    display: flex;
    flex-direction: row;
    flex-wrap: wrap;
    gap: 8px;
    background-color: #ffffff !important;
    padding: 10px !important;
    border-radius: 12px !important;
    border: 1px solid #cbd5e1 !important;
    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.02) !important;
    margin-bottom: 25px !important;
    margin-top: -10px !important; /* Pull closer to the header */
}

/* Style each horizontal option wrapper */
div[data-testid="stRadio"] > div[role="radiogroup"] > div[data-testid="stRadioOption"] {
    background: transparent !important;
    margin: 0 !important;
}

/* Style each horizontal option label */
div[data-testid="stRadio"] > div[role="radiogroup"] > div[data-testid="stRadioOption"] label {
    padding: 10px 18px !important;
    border-radius: 8px !important;
    background-color: #f1f5f9 !important; /* Light gray for inactive tabs */
    color: #475569 !important; /* Gray text */
    font-weight: 700 !important;
    border: 1px solid #cbd5e1 !important;
    transition: all 0.2s ease-in-out !important;
    cursor: pointer !important;
}

/* Horizontal option label hover state */
div[data-testid="stRadio"] > div[role="radiogroup"] > div[data-testid="stRadioOption"] label:hover {
    background-color: #fee2e2 !important; /* Light crimson on hover */
    color: #b91c1c !important; /* Crimson red text */
    border-color: #fca5a5 !important;
}

/* Active selected horizontal option label state */
div[data-testid="stRadio"] > div[role="radiogroup"] > div[data-testid="stRadioOption"]:has(input:checked) label {
    color: #ffffff !important; /* White text */
    background-color: #b91c1c !important; /* Solid Crimson background */
    border-color: #991b1b !important;
    box-shadow: 0 4px 12px rgba(185, 28, 28, 0.25) !important;
}

/* KPI Card Style Rules (Light Mode) */
.kpi-card {
    background-color: #ffffff !important;
    border: 1px solid #e2e8f0 !important;
    border-top: 3px solid #b91c1c !important; /* Crimson top border */
    border-radius: 12px !important;
    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.03) !important;
    color: #1e293b !important;
    padding: 1rem 1.25rem;
    display: flex;
    align-items: center;
    gap: 0;
    margin-bottom: 12px;
    height: 90px !important;
    box-sizing: border-box;
}
.kpi-label {
    color: #64748b;
    font-size: 0.72rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    line-height: 1;
}
.kpi-value {
    color: #0f172a;
    font-size: 1.35rem;
    font-weight: 800;
    line-height: 1.2;
    margin-top: 3px;
}

.small-note {
    font-size: 0.9rem;
    color: #64748b;
    margin-bottom: 1rem;
}

/* Remove layout space occupied by the chat bubble wrapper in document flow */
div[data-testid="element-container"]:has(a[href="?chat=toggle"]) {
    position: fixed !important;
    bottom: 0 !important;
    right: 0 !important;
    width: 0 !important;
    height: 0 !important;
    overflow: visible !important;
    z-index: 999999 !important;
}

/* Remove layout space occupied by the chat window wrapper in document flow */
div[data-testid="stVerticalBlockBorderWrapper"]:has(div[key="floating_chat"]) {
    position: fixed !important;
    bottom: 90px !important;
    right: 20px !important;
    width: 380px !important;
    height: 520px !important;
    z-index: 999999 !important;
}

/* Floating Chat Window CSS key overrides (Small widget layout) */
div[key="floating_chat"] {
    width: 100% !important;
    height: 100% !important;
    background-color: #ffffff !important; /* White background */
    border: 2px solid #b91c1c !important; /* Crimson border */
    border-radius: 12px !important;
    box-shadow: 0 10px 25px rgba(0,0,0,0.15) !important;
    padding: 0 !important;
    display: flex !important;
    flex-direction: column !important;
    overflow-y: auto !important;
}

/* Floating chat header styling */
.floating-chat-header {
    background: linear-gradient(135deg, #b91c1c 0%, #7f1d1d 100%) !important; /* Crimson header background */
    color: #ffffff !important;
    padding: 12px 15px !important;
    font-weight: 700 !important;
    display: flex !important;
    justify-content: space-between !important;
    align-items: center !important;
    border-top-left-radius: 10px !important;
    border-top-right-radius: 10px !important;
    margin-bottom: 15px !important;
}
.floating-chat-header span {
    color: #ffffff !important;
}

/* Scrollbar for chat container */
div[key="floating_chat"]::-webkit-scrollbar {
    width: 6px;
}
div[key="floating_chat"]::-webkit-scrollbar-track {
    background: #ffffff;
}
div[key="floating_chat"]::-webkit-scrollbar-thumb {
    background: #cbd5e1;
    border-radius: 3px;
}

/* Header-specific color overrides.
   These rules are intentionally placed after the global h1/span rules so
   Streamlit's Markdown styles cannot force the portal header back to black. */
.portal-agency-label,
.portal-main-title,
.portal-subtitle,
.portal-separator,
.portal-current-date {
    color: #f59e0b !important;
    -webkit-text-fill-color: #f59e0b !important;
    opacity: 1 !important;
}

.portal-main-title * ,
.portal-subtitle * ,
.portal-separator * ,
.portal-current-date * {
    color: #f59e0b !important;
    -webkit-text-fill-color: #f59e0b !important;
}
</style>

"""
st.markdown(CSS, unsafe_allow_html=True)

st.markdown(
    """
    <style>
    /* Floating AI button rendered by Streamlit, independent of page flow. */
    .st-key-open_ai_chat {
        position: fixed !important;
        right: 24px !important;
        bottom: 22px !important;
        width: 128px !important;
        height: 48px !important;
        z-index: 1000000 !important;
    }

    .st-key-open_ai_chat > div,
    .st-key-open_ai_chat [data-testid="stButton"] {
        width: 128px !important;
        height: 48px !important;
    }

    .st-key-open_ai_chat button {
    width: 128px !important;
    min-width: 128px !important;
    height: 48px !important;
    min-height: 48px !important;
    padding: 0 18px !important;
    border-radius: 999px !important;
    border: 3px solid #f59e0b !important;
    background: linear-gradient(
        135deg,
        #b91c1c 0%,
        #7f1d1d 100%
    ) !important;

    color: #ffffff !important;
    -webkit-text-fill-color: #ffffff !important;

    font-size: 15px !important;
    font-weight: 700 !important;
    line-height: 1.2 !important;
    box-shadow: 0 8px 24px rgba(15, 23, 42, 0.32) !important;
}

    .st-key-open_ai_chat button:hover {
        border-color: #fbbf24 !important;
        background: linear-gradient(135deg, #991b1b 0%, #6f1717 100%) !important;
        color: #ffffff !important;
        transform: translateY(-2px) !important;
    }

    /*
      st.dialog is a real modal/fragment. These rules move the small modal
      to the bottom-right so it behaves like a website chat widget.
    */
    div[data-testid="stDialog"] {
        align-items: flex-end !important;
        justify-content: flex-end !important;
        padding: 0 24px 94px 0 !important;
    }

    div[data-testid="stDialog"] > div[role="dialog"],
    div[data-testid="stDialog"] div[role="dialog"] {
        width: min(410px, calc(100vw - 28px)) !important;
        max-width: 410px !important;
        max-height: min(690px, calc(100vh - 120px)) !important;
        margin: 0 !important;
        border: 1px solid #d97706 !important;
        border-radius: 16px !important;
        box-shadow: 0 22px 60px rgba(15, 23, 42, 0.30) !important;
        overflow: hidden !important;
    }

    div[data-testid="stDialog"] [data-testid="stDialogContent"] {
        background: #f8fafc !important;
        padding: 0.85rem 1rem 1rem 1rem !important;
        overflow-y: auto !important;
    }

    div[data-testid="stDialog"] [data-testid="stDialogHeader"] {
        background: linear-gradient(135deg, #b91c1c 0%, #7f1d1d 100%) !important;
        border-bottom: 2px solid #d97706 !important;
        padding: 0.8rem 1rem !important;
    }

    div[data-testid="stDialog"] [data-testid="stDialogHeader"] *,
    div[data-testid="stDialog"] [data-testid="stDialogHeader"] button {
        color: #ffffff !important;
    }

    /* Keep the conversation compact inside the modal. */
    .st-key-ai_chat_history {
        border: 1px solid #e2e8f0 !important;
        border-radius: 12px !important;
        background: #ffffff !important;
        padding: 0.35rem !important;
    }

    @media (max-width: 640px) {
        .st-key-open_ai_chat {
            right: 14px !important;
            bottom: 14px !important;
        }

        div[data-testid="stDialog"] {
            padding: 0 10px 84px 10px !important;
        }

        div[data-testid="stDialog"] > div[role="dialog"],
        div[data-testid="stDialog"] div[role="dialog"] {
            width: calc(100vw - 20px) !important;
            max-width: none !important;
            max-height: calc(100vh - 100px) !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)



@st.cache_resource
def bootstrap() -> bool:
    init_db()
    seed_demo_data()
    return True


bootstrap()


def header() -> None:
    import datetime
    now = datetime.datetime.now()
    weekdays_vi = ["Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu", "Thứ Bảy", "Chủ Nhật"]
    weekday = weekdays_vi[now.weekday()]
    current_date_vietnamese = f"{weekday}, ngày {now.strftime('%d/%m/%Y')}"
    
    st.markdown(
        f"""
        <div style="
            background: linear-gradient(135deg, #991b1b 0%, #7f1d1d 100%);
            border-bottom: 3px solid #d97706;
            padding: 1.5rem 2rem;
            margin-bottom: 2rem;
            border-radius: 12px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            position: relative;
            overflow: hidden;
        ">
            <!-- Decorative flag stripe -->
            <div style="position: absolute; top: 0; left: 0; right: 0; height: 4px; background: linear-gradient(90deg, #f59e0b 0%, #b91c1c 50%, #f59e0b 100%);"></div>
            <div style="display: flex; align-items: center; gap: 20px; flex-wrap: wrap;">
                <!-- Golden Crest Emblem -->
                <div style="
                    width: 60px;
                    height: 60px;
                    border-radius: 50%;
                    background: radial-gradient(circle, #f59e0b 20%, #b91c1c 80%);
                    border: 2px solid #f59e0b;
                    box-shadow: 0 0 10px rgba(245, 158, 11, 0.5);
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    flex-shrink: 0;
                ">
                    <!-- Golden Star SVG -->
                    <svg xmlns="https://www.istockphoto.com/vi/b%E1%BB%A9c-%E1%BA%A3nh/vietnam-flag" width="36" height="36" viewBox="0 0 24 24" fill="#f59e0b" stroke="#d97706" stroke-width="1"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/></svg>
                </div>
                <div style="flex-grow: 1;">
                    <div class="portal-agency-label" style="font-size: 0.85rem; font-weight: 700; text-transform: uppercase; letter-spacing: 1.5px; margin-bottom: 2px;">
                        ỦY BAN NHÂN DÂN XÃ BÀ NÀ
                    </div>
                    <h1 class="portal-main-title" style="margin: 0; font-size: 1.6rem; font-weight: 800; text-transform: uppercase; letter-spacing: 0.5px; line-height: 1.2;">
                        CỔNG TỔNG HỢP SỐ LIỆU & CHỈ TIÊU KINH TẾ - XÃ HỘI CÁC THÔN
                    </h1>
                    <div style="display: flex; align-items: center; gap: 15px; margin-top: 8px; flex-wrap: wrap;">
                        <span class="portal-subtitle" style="font-size: 0.85rem; font-weight: 600;">
                            Hệ thống liên kết dữ liệu thông minh (SmartLink Portal)
                        </span>
                        <span class="portal-separator" style="font-size: 0.85rem;">|</span>
                        <span class="portal-current-date" style="font-size: 0.85rem; font-weight: 600; display: flex; align-items: center; gap: 6px;">
                            Hôm nay: {current_date_vietnamese}
                        </span>
                    </div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def metric_card(label: str, value: str) -> str:
    return f"""
    <div class="kpi-card">
        <div>
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
        </div>
    </div>
    """


def configure_dark_chart(chart: alt.Chart) -> alt.Chart:
    # Force transparent backgrounds, clean light-themed gridlines and dark text colors
    return chart.configure_view(
        stroke=None
    ).configure_axis(
        grid=True,
        gridColor="#e2e8f0",
        labelColor="#475569",
        titleColor="#1e293b",
        domainColor="#cbd5e1",
        tickColor="#cbd5e1"
    ).configure_title(
        color="#1e293b"
    ).configure_legend(
        labelColor="#475569",
        titleColor="#1e293b"
    )


def period_selector(default: str = "2026-Q2") -> str:
    periods = list_periods()
    if default not in periods:
        periods = [default] + periods
    return st.selectbox("Chọn kỳ báo cáo cần xem xét", periods, index=periods.index(default) if default in periods else 0)



def render_missing_fields_table(period: str, show_title: bool = True) -> pd.DataFrame:
    """Render the follow-up table for missing/incomplete inputs."""
    missing_df = build_missing_fields_dataframe(period)
    if show_title:
        st.markdown("### Bảng dữ liệu thiếu cần bổ sung")
    if missing_df.empty:
        st.success("Không phát hiện dữ liệu thiếu cần bổ sung trong kỳ này.")
    else:
        blocker_count = int((missing_df["Mức độ"] == "BLOCKER").sum())
        warning_count = int((missing_df["Mức độ"] == "WARNING").sum())
        st.caption(f"Tổng cộng {len(missing_df)} mục cần xử lý: {blocker_count} lỗi nghiêm trọng và {warning_count} cảnh báo.")
        display_df = missing_df.copy()
        display_df["Mức độ"] = display_df["Mức độ"].map({
            "BLOCKER": "BLOCKER",
            "WARNING": "WARNING",
        }).fillna(display_df["Mức độ"])
        with st.container(border=True):
            st.dataframe(display_df, use_container_width=True, hide_index=True)
    return missing_df


def page_overview() -> None:
    """Render the commune overview with an interactive CT01-CT14 dashboard."""
    period = period_selector()

    villages = pd.DataFrame(fetch_villages())
    reports = pd.DataFrame(fetch_reports(period))
    tasks = pd.DataFrame(fetch_tasks(period))

    total_villages = len(villages)
    submitted = reports["village_name"].nunique() if not reports.empty else 0
    not_submitted = max(total_villages - submitted, 0)
    late = (
        int(tasks["status"].astype(str).str.contains("Trễ|Chưa", regex=True).sum())
        if not tasks.empty and "status" in tasks.columns
        else 0
    )
    warnings = (
        int(tasks["status"].astype(str).str.contains("Cảnh báo").sum())
        if not tasks.empty and "status" in tasks.columns
        else 0
    )

    st.markdown("### Tổng quan số liệu toàn xã")
    st.caption(
        "Theo dõi đầy đủ 14 chỉ tiêu CT01-CT14 theo kỳ báo cáo, thôn và chỉ tiêu được chọn."
    )

    # Operational progress KPIs remain visible above the analytical dashboard.
    top1, top2, top3, top4 = st.columns(4)
    with top1:
        st.markdown(metric_card("Tổng số thôn", f"{total_villages}"), unsafe_allow_html=True)
    with top2:
        st.markdown(metric_card("Đã nộp", f"{submitted}"), unsafe_allow_html=True)
    with top3:
        st.markdown(metric_card("Chưa nộp", f"{not_submitted}"), unsafe_allow_html=True)
    with top4:
        st.markdown(metric_card("Cần xử lý", f"{late + warnings}"), unsafe_allow_html=True)

    # Show operational warnings before analytical charts.
    warning_tasks = (
        tasks[tasks["status"] == "Cảnh báo"]
        if not tasks.empty and "status" in tasks.columns
        else pd.DataFrame()
    )
    if not warning_tasks.empty:
        warning_list = []
        for _, row in warning_tasks.iterrows():
            village_name = str(row.get("village_name", "Không xác định"))
            village_report = (
                reports[reports["village_name"] == village_name]
                if not reports.empty and "village_name" in reports.columns
                else pd.DataFrame()
            )
            phone = ""
            if not village_report.empty and "phone" in village_report.columns:
                phone_value = village_report.iloc[0].get("phone")
                if pd.notna(phone_value) and str(phone_value).strip():
                    phone = f" (SĐT: {phone_value})"
            warning_list.append(f"**{village_name}**{phone}")

        st.warning(
            "**Cảnh báo chất lượng số liệu**: "
            f"Phát hiện **{len(warning_tasks)} thôn** có dữ liệu cần rà soát: "
            f"{', '.join(warning_list)}."
        )

    if reports.empty:
        st.info("Chưa có dữ liệu CT01-CT14 để xây dựng dashboard cho kỳ đã chọn.")
        return

    indicator_codes = {
        field: f"CT{index:02d}"
        for index, field in enumerate(INDICATOR_FIELDS, start=1)
    }
    indicator_labels = {
        field: f"{indicator_codes[field]} - {FIELD_LABELS.get(field, field)}"
        for field in INDICATOR_FIELDS
    }
    indicator_units = {
        "ct01_households": "hộ",
        "ct02_population": "người",
        "ct03_poor_households": "hộ",
        "ct04_near_poor_households": "hộ",
        "ct05_revolution_contributors": "người",
        "ct06_social_protection": "người",
        "ct07_children_under_16": "người",
        "ct08_special_children": "người",
        "ct09_cultural_households": "hộ",
        "ct10_working_age": "người",
        "ct11_health_insurance": "người",
        "ct12_digital_team_members": "người",
        "ct13_online_public_service_guided": "người",
        "ct14_domestic_violence_cases": "vụ",
    }
    indicator_groups = {
        "ct01_households": "Dân số và hộ dân",
        "ct02_population": "Dân số và hộ dân",
        "ct03_poor_households": "An sinh xã hội",
        "ct04_near_poor_households": "An sinh xã hội",
        "ct05_revolution_contributors": "An sinh xã hội",
        "ct06_social_protection": "An sinh xã hội",
        "ct07_children_under_16": "Trẻ em",
        "ct08_special_children": "Trẻ em",
        "ct09_cultural_households": "Văn hóa và xã hội",
        "ct10_working_age": "Lao động và y tế",
        "ct11_health_insurance": "Lao động và y tế",
        "ct12_digital_team_members": "Chuyển đổi số",
        "ct13_online_public_service_guided": "Chuyển đổi số",
        "ct14_domestic_violence_cases": "An toàn xã hội",
    }
    label_to_field = {label: field for field, label in indicator_labels.items()}

    # Guarantee a stable dashboard even when a legacy database is missing a column.
    dashboard_df = reports.copy()
    for field in INDICATOR_FIELDS:
        if field not in dashboard_df.columns:
            dashboard_df[field] = 0
        dashboard_df[field] = pd.to_numeric(
            dashboard_df[field], errors="coerce"
        ).fillna(0)

    if "village_name" not in dashboard_df.columns:
        st.error("Dữ liệu báo cáo không có trường tên thôn để thực hiện bộ lọc.")
        return

    village_options = sorted(
        dashboard_df["village_name"].dropna().astype(str).unique().tolist()
    )
    all_option = "Tất cả"
    village_filter_options = [all_option, *village_options]
    indicator_option_labels = list(indicator_labels.values())
    indicator_filter_options = [all_option, *indicator_option_labels]

    st.markdown("### Dashboard CT01-CT14")
    with st.container(border=True):
        st.markdown("##### Bộ lọc dữ liệu")
        filter_col1, filter_col2, filter_col3 = st.columns([1.35, 1.65, 0.65])

        with filter_col1:
            selected_village_choices = st.multiselect(
                "Lọc theo thôn",
                village_filter_options,
                default=[all_option],
                key=f"overview_villages_ct14_{period}",
                placeholder="Chọn Tất cả hoặc một hay nhiều thôn",
            )
            selected_villages = (
                village_options
                if all_option in selected_village_choices
                else selected_village_choices
            )

        with filter_col2:
            selected_indicator_choices = st.multiselect(
                "Chỉ tiêu hiển thị",
                indicator_filter_options,
                default=[all_option],
                key=f"overview_indicators_ct14_{period}",
                placeholder="Chọn Tất cả hoặc một hay nhiều chỉ tiêu",
            )
            selected_indicator_labels = (
                indicator_option_labels
                if all_option in selected_indicator_choices
                else selected_indicator_choices
            )

        with filter_col3:
            max_top_n = max(len(village_options), 1)
            default_top_n = min(10, max_top_n)
            top_n = st.number_input(
                "Số thôn so sánh",
                min_value=1,
                max_value=max_top_n,
                value=default_top_n,
                step=1,
                key=f"overview_top_n_ct14_{period}",
            )

    if not selected_villages:
        st.warning("Hãy chọn ít nhất một thôn để hiển thị dashboard.")
        return

    if not selected_indicator_labels:
        st.warning("Hãy chọn ít nhất một chỉ tiêu CT01-CT14.")
        return

    selected_fields = [label_to_field[label] for label in selected_indicator_labels]
    filtered_df = dashboard_df[
        dashboard_df["village_name"].astype(str).isin(selected_villages)
    ].copy()

    if filtered_df.empty:
        st.info("Không có dữ liệu phù hợp với bộ lọc hiện tại.")
        return

    # Indicator KPI cards recalculate immediately after every filter change.
    totals = {
        field: int(filtered_df[field].sum())
        for field in INDICATOR_FIELDS
    }
    st.markdown("#### Tổng hợp các chỉ tiêu đã chọn")
    for start_index in range(0, len(selected_fields), 3):
        row_fields = selected_fields[start_index:start_index + 3]
        row_columns = st.columns(len(row_fields))
        for column, field in zip(row_columns, row_fields):
            with column:
                value = f"{totals[field]:,}".replace(",", ".")
                unit = indicator_units.get(field, "")
                st.markdown(
                    metric_card(indicator_labels[field], f"{value} {unit}".strip()),
                    unsafe_allow_html=True,
                )

    # Chart 1: aggregate totals for all selected CT01-CT14 indicators.
    aggregate_df = pd.DataFrame(
        [
            {
                "Mã": indicator_codes[field],
                "Chỉ tiêu": indicator_labels[field],
                "Giá trị": int(filtered_df[field].sum()),
                "Nhóm": indicator_groups.get(field, "Khác"),
                "Đơn vị": indicator_units.get(field, ""),
            }
            for field in selected_fields
        ]
    ).sort_values("Giá trị", ascending=True)

    aggregate_bars = alt.Chart(aggregate_df).mark_bar(
        cornerRadiusTopRight=5,
        cornerRadiusBottomRight=5,
    ).encode(
        y=alt.Y(
            "Chỉ tiêu:N",
            sort=alt.SortField(field="Giá trị", order="descending"),
            title=None,
            axis=alt.Axis(labelLimit=310),
        ),
        x=alt.X("Giá trị:Q", title="Tổng giá trị"),
        color=alt.Color("Nhóm:N", title="Nhóm chỉ tiêu"),
        tooltip=[
            alt.Tooltip("Mã:N", title="Mã"),
            alt.Tooltip("Chỉ tiêu:N", title="Chỉ tiêu"),
            alt.Tooltip("Giá trị:Q", title="Tổng", format=",d"),
            alt.Tooltip("Đơn vị:N", title="Đơn vị"),
        ],
    )

    aggregate_labels = alt.Chart(aggregate_df).mark_text(
        align="left",
        baseline="middle",
        dx=5,
        fontSize=11,
        color="#334155",
    ).encode(
        y=alt.Y(
            "Chỉ tiêu:N",
            sort=alt.SortField(field="Giá trị", order="descending"),
        ),
        x=alt.X("Giá trị:Q"),
        text=alt.Text("Giá trị:Q", format=",d"),
    )

    aggregate_chart = configure_dark_chart(
        (aggregate_bars + aggregate_labels).properties(
            height=max(360, len(selected_fields) * 31),
            background="transparent",
        )
    )

    # Chart 2: compare one selected indicator across villages.
    comparison_label = st.selectbox(
        "Chỉ tiêu dùng để so sánh giữa các thôn",
        selected_indicator_labels,
        key=f"overview_compare_indicator_ct14_{period}",
    )
    comparison_field = label_to_field[comparison_label]
    comparison_df = filtered_df[["village_name", comparison_field]].copy()
    comparison_df = comparison_df.rename(
        columns={"village_name": "Thôn", comparison_field: "Giá trị"}
    )
    comparison_df = comparison_df.sort_values("Giá trị", ascending=False).head(int(top_n))

    comparison_chart = alt.Chart(comparison_df).mark_bar(
        cornerRadiusTopRight=5,
        cornerRadiusBottomRight=5,
    ).encode(
        y=alt.Y("Thôn:N", sort="-x", title=None, axis=alt.Axis(labelLimit=230)),
        x=alt.X("Giá trị:Q", title=comparison_label),
        tooltip=[
            alt.Tooltip("Thôn:N", title="Thôn"),
            alt.Tooltip("Giá trị:Q", title="Giá trị", format=",d"),
        ],
    ).properties(
        height=max(360, min(int(top_n) * 32, 650)),
        background="transparent",
    )
    comparison_chart = configure_dark_chart(comparison_chart)

    chart_left, chart_right = st.columns([1.15, 1])
    with chart_left:
        with st.container(border=True):
            st.markdown("##### Tổng số liệu CT01-CT14 theo bộ lọc")
            st.altair_chart(aggregate_chart, use_container_width=True, theme=None)

    with chart_right:
        with st.container(border=True):
            st.markdown("##### So sánh một chỉ tiêu giữa các thôn")
            st.altair_chart(comparison_chart, use_container_width=True, theme=None)

    # Chart 3: analyse one derived rate across villages.
    ratio_definitions = {
        "Tỷ lệ hộ nghèo trên tổng số hộ": (
            "ct03_poor_households",
            "ct01_households",
        ),
        "Tỷ lệ hộ cận nghèo trên tổng số hộ": (
            "ct04_near_poor_households",
            "ct01_households",
        ),
        "Tỷ lệ hộ đạt Gia đình văn hóa": (
            "ct09_cultural_households",
            "ct01_households",
        ),
        "Tỷ lệ trẻ em dưới 16 tuổi": (
            "ct07_children_under_16",
            "ct02_population",
        ),
        "Tỷ lệ người trong độ tuổi lao động": (
            "ct10_working_age",
            "ct02_population",
        ),
        "Tỷ lệ người tham gia BHYT": (
            "ct11_health_insurance",
            "ct02_population",
        ),
    }
    ratio_label = st.selectbox(
        "Tỷ lệ trọng yếu cần phân tích",
        list(ratio_definitions.keys()),
        key=f"overview_ratio_indicator_ct14_{period}",
    )
    numerator_field, denominator_field = ratio_definitions[ratio_label]
    ratio_df = filtered_df[
        ["village_name", numerator_field, denominator_field]
    ].copy()
    denominator = ratio_df[denominator_field].replace(0, pd.NA)
    ratio_df["Tỷ lệ"] = (
        ratio_df[numerator_field].div(denominator).mul(100).fillna(0)
    )
    ratio_df = ratio_df.rename(columns={"village_name": "Thôn"})
    ratio_df = ratio_df.sort_values("Tỷ lệ", ascending=False).head(int(top_n))

    ratio_chart = alt.Chart(ratio_df).mark_bar(
        cornerRadiusTopRight=5,
        cornerRadiusBottomRight=5,
    ).encode(
        y=alt.Y("Thôn:N", sort="-x", title=None, axis=alt.Axis(labelLimit=230)),
        x=alt.X("Tỷ lệ:Q", title=f"{ratio_label} (%)"),
        tooltip=[
            alt.Tooltip("Thôn:N", title="Thôn"),
            alt.Tooltip("Tỷ lệ:Q", title="Tỷ lệ", format=".2f"),
        ],
    ).properties(
        height=max(330, min(int(top_n) * 31, 620)),
        background="transparent",
    )
    ratio_chart = configure_dark_chart(ratio_chart)

    with st.container(border=True):
        st.markdown(f"##### {ratio_label} theo thôn")
        st.altair_chart(ratio_chart, use_container_width=True, theme=None)

    # Detailed CT01-CT14 table remains complete even when chart indicators are filtered.
    detail_df = filtered_df[["village_name", *INDICATOR_FIELDS]].copy()
    detail_df = detail_df.rename(
        columns={
            "village_name": "Thôn",
            **{
                field: indicator_labels[field]
                for field in INDICATOR_FIELDS
            },
        }
    )
    population_label = indicator_labels.get("ct02_population")
    if population_label in detail_df.columns:
        detail_df = detail_df.sort_values(population_label, ascending=False)

    column_config = {}
    for field in INDICATOR_FIELDS:
        label = indicator_labels[field]
        unit = indicator_units.get(field, "")
        suffix = f" {unit}" if unit else ""
        column_config[label] = st.column_config.NumberColumn(format=f"%d{suffix}")

    st.markdown("### Bảng dữ liệu CT01-CT14 sau lọc")
    st.caption(
        "Bảng hiển thị đầy đủ 14 chỉ tiêu; sử dụng thanh cuộn ngang để xem các cột phía sau."
    )
    with st.container(border=True):
        st.dataframe(
            detail_df,
            use_container_width=True,
            hide_index=True,
            column_config=column_config,
            height=460,
        )

def page_dashboard_charts() -> None:
    st.subheader("Bảng điều khiển số liệu")
    period = period_selector()

    summary = build_summary_dataframe(period)

    st.markdown("**Bảng trạng thái và chỉ tiêu báo cáo các thôn**")
    if summary.empty:
        st.info("Chưa có dữ liệu.")
        return

    progress_cols = [
        "village_name",
        "period",
        "phone",
        "submission_status",
        "due_at",
        "submitted_at",
        "days_late",
    ]
    progress_df = summary[
        [column for column in progress_cols if column in summary.columns]
    ].copy()

    # Filter by village name. normalize_text makes the search case-insensitive
    # and tolerant of Vietnamese accents, for example "dong son" matches
    # "Đông Sơn".
    with st.container(border=True):
        st.markdown("##### Tra cứu theo tên thôn")
        village_search = st.text_input(
            "Tên thôn cần tìm",
            placeholder="Ví dụ: Đông Sơn, Ninh An hoặc dong son...",
            key=f"dashboard_village_search_{period}",
        ).strip()

    total_rows_before_filter = len(progress_df)
    if village_search and "village_name" in progress_df.columns:
        search_normalized = normalize_text(village_search)
        village_normalized = progress_df["village_name"].fillna("").astype(str).map(normalize_text)
        progress_df = progress_df[
            village_normalized.str.contains(search_normalized, regex=False)
        ].copy()

    if "submission_status" in progress_df.columns:
        progress_df["submission_status"] = progress_df["submission_status"].map({
            "Đúng hạn": "Đúng hạn",
            "Đã nộp": "Đã nộp",
            "Trễ hạn": "Trễ hạn",
            "Chưa nộp": "Chưa nộp",
            "Hoàn thành": "Hoàn thành",
            "Cảnh báo": "Cảnh báo (SĐT sai)",
        }).fillna(progress_df["submission_status"])

    progress_labels = {
        "village_name": "Thôn",
        "period": "Kỳ báo cáo",
        "phone": "Số điện thoại",
        "submission_status": "Trạng thái",
        "due_at": "Hạn nộp",
        "submitted_at": "Thời điểm nộp",
        "days_late": "Số ngày trễ",
    }

    if village_search:
        st.caption(
            f"Kết quả tra cứu: {len(progress_df)}/{total_rows_before_filter} thôn phù hợp."
        )
    else:
        st.caption(f"Đang hiển thị toàn bộ {total_rows_before_filter} thôn.")

    with st.container(border=True):
        if progress_df.empty:
            st.info("Không tìm thấy thôn phù hợp với từ khóa đã nhập.")
        else:
            st.dataframe(
                progress_df.rename(columns=progress_labels),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Số ngày trễ": st.column_config.NumberColumn(format="%d ngày"),
                },
            )

    st.markdown("<div style='height: 18px;'></div>", unsafe_allow_html=True)
    render_missing_fields_table(period, show_title=True)

def page_upload() -> None:
    st.subheader("Tải lên báo cáo số liệu thôn")
    
    with st.container(border=True):
        st.markdown("### Bước 1: Chuẩn bị file báo cáo")
        left, right = st.columns([2, 1])
        with left:
            st.markdown(
                """
                Hệ thống hỗ trợ tự động chuẩn hóa 3 định dạng báo cáo Excel:
                1. **Phiếu báo cáo thôn**: Mẫu biểu chuẩn gồm các chỉ tiêu từ `CT01` đến `CT14`.
                2. **Báo cáo tổng hợp**: Bảng theo dõi tiến độ & số liệu tổng hợp các thôn từ Google Drive.
                3. **Bảng phẳng tự tạo**: Bảng dữ liệu thô có chứa cột Tên thôn và các mã chỉ tiêu tương ứng.
                """
            )
        with right:
            if TEMPLATE_PATH.exists():
                st.download_button(
                    " Tải mẫu Excel (Xã gửi Thôn)",
                    data=TEMPLATE_PATH.read_bytes(),
                    file_name=TEMPLATE_PATH.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            sample_files = sorted(DRIVE_SAMPLE_DIR.glob("*.xlsx"))
            st.caption(f" Đã cấu hình sẵn **{len(sample_files)}** file mẫu trong `sample_data/drive_imported` để phục vụ demo nhanh.")

    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
    
    with st.container(border=True):
        st.markdown("### Bước 2: Chọn file và tải lên hệ thống")
        uploaded = st.file_uploader("Kéo thả hoặc nhấn để chọn file Excel (.xlsx)", type=["xlsx"], label_visibility="collapsed")
        
    if uploaded is None:
        st.info("Gợi ý thử nghiệm: dùng `BC_T20_Thon_Dong_Son.xlsx` để xem cảnh báo sai định dạng điện thoại, hoặc file tổng hợp để import toàn bộ 22 thôn.")
        return

    try:
        raw_preview = get_raw_preview(uploaded)
        st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)
        st.subheader("Xem trước dữ liệu thô")
        with st.container(border=True):
            st.dataframe(raw_preview, use_container_width=True, hide_index=True)

        uploaded.seek(0)
        parsed = parse_excel_to_reports(uploaded, uploaded.name)
        clean, issues = validate_reports(parsed, allow_not_submitted_rows=True)
    except Exception as exc:
        st.error(f"Không xử lý được file: {exc}")
        return

    st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)
    st.subheader("Dữ liệu chuẩn hóa chuyển về kho dùng chung")

    metadata_cols = [
        "commune_name",
        "village_name",
        "period",
        "reporter_name",
        "reporter_title",
        "phone",
        "submission_status",
        "due_at",
        "submitted_at",
        "days_late",
    ]
    metadata_labels = {
        "commune_name": "Xã",
        "village_name": "Thôn",
        "period": "Kỳ báo cáo",
        "reporter_name": "Người lập báo cáo",
        "reporter_title": "Chức danh",
        "phone": "Số điện thoại",
        "submission_status": "Trạng thái",
        "due_at": "Hạn nộp",
        "submitted_at": "Thời điểm nộp",
        "days_late": "Số ngày trễ",
    }
    indicator_labels = {
        field: f"{field[:4].upper()} - {FIELD_LABELS[field]}"
        for field in INDICATOR_FIELDS
        if field in FIELD_LABELS
    }

    metadata_tab, indicator_tab = st.tabs([
        "Thông tin báo cáo",
        "Chỉ tiêu CT01-CT14",
    ])

    with metadata_tab:
        metadata_df = clean[
            [column for column in metadata_cols if column in clean.columns]
        ].rename(columns=metadata_labels)
        with st.container(border=True):
            st.dataframe(
                metadata_df,
                use_container_width=True,
                hide_index=True,
            )

    with indicator_tab:
        indicator_cols = ["village_name", "period", *INDICATOR_FIELDS]
        available_indicator_cols = [
            column for column in indicator_cols if column in clean.columns
        ]
        indicator_df = clean[available_indicator_cols].rename(
            columns={
                "village_name": "Thôn",
                "period": "Kỳ báo cáo",
                **indicator_labels,
            }
        )
        indicator_number_config = {
            indicator_labels[field]: st.column_config.NumberColumn(format="%d")
            for field in INDICATOR_FIELDS
            if field in clean.columns
        }
        st.caption(
            "Kiểm tra thanh cuộn ngang để xác nhận tất cả CT01-CT14 trước khi lưu."
        )
        with st.container(border=True):
            st.dataframe(
                indicator_df,
                use_container_width=True,
                hide_index=True,
                column_config=indicator_number_config,
                height=420,
            )

    st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)
    st.subheader("Kết quả kiểm định AI Validator")
    
    if issues:
        issues_df = pd.DataFrame(issues)
        issues_disp = issues_df.copy()
        issues_disp["severity"] = issues_disp["severity"].map({
            "BLOCKER": "BLOCKER (Lỗi)",
            "WARNING": "WARNING (Cảnh báo)"
        }).fillna(issues_disp["severity"])
        
        issues_disp = issues_disp.rename(columns={
            "village_name": "Thôn",
            "period": "Kỳ báo cáo",
            "severity": "Mức độ",
            "error_type": "Loại kiểm định",
            "field_name": "Trường dữ liệu",
            "message": "Nội dung phản hồi"
        })
        
        with st.container(border=True):
            st.dataframe(
                issues_disp[["Thôn", "Kỳ báo cáo", "Mức độ", "Loại kiểm định", "Trường dữ liệu", "Nội dung phản hồi"]], 
                use_container_width=True, 
                hide_index=True
            )
            
        blocker_count = int((issues_df["severity"] == "BLOCKER").sum())
        warning_count = int((issues_df["severity"] == "WARNING").sum())
        
        if blocker_count > 0:
            st.error(f"Phát hiện **{blocker_count}** lỗi nghiêm trọng (BLOCKER). Vui lòng điều chỉnh lại file Excel để được phép lưu vào hệ thống.")
        if warning_count > 0:
            st.warning(f"Phát hiện **{warning_count}** cảnh báo (WARNING). Bạn vẫn có thể lưu dữ liệu, cán bộ kiểm định sẽ rà soát lại sau.")
    else:
        st.success("Hoàn hảo! Không phát hiện lỗi logic hay lỗi định dạng dữ liệu nào.")

    if has_blockers(issues):
        st.button("Không thể lưu khi dữ liệu còn lỗi nghiêm trọng (BLOCKER)", disabled=True, use_container_width=True)
        return

    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
    if st.button("Lưu dữ liệu chuẩn hóa vào Kho dùng chung", type="primary", use_container_width=True):
        saved_ids = upsert_reports(clean.to_dict(orient="records"), source_file=uploaded.name)
        insert_validation_logs(issues, source_file=uploaded.name)
        st.success(f"Đã cập nhật thành công {len(saved_ids)} báo cáo thôn. Dữ liệu tiến độ đã được đồng bộ.")
        st.rerun()


def page_quality() -> None:
    st.subheader("Nhật ký kiểm soát chất lượng dữ liệu")
    period = period_selector()

    # New follow-up table: missing fields that need correction or supplementation.
    missing_df = render_missing_fields_table(period, show_title=True)
    st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)

    logs = pd.DataFrame(fetch_validation_logs(period))
    if logs.empty:
        if missing_df.empty:
            st.success("Không có cảnh báo dữ liệu nào được ghi nhận cho kỳ báo cáo này.")
        else:
            st.info("Bảng phía trên đang được suy luận trực tiếp từ dữ liệu tổng hợp/trạng thái nộp. Chưa có log AI Validator chi tiết trong database.")
        return

    blocker_count = int((logs["severity"] == "BLOCKER").sum())
    warning_count = int((logs["severity"] == "WARNING").sum())

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(metric_card("Tổng cảnh báo", f"{len(logs)}"), unsafe_allow_html=True)
    with col2:
        st.markdown(metric_card("Lỗi BLOCKER", f"{blocker_count}"), unsafe_allow_html=True)
    with col3:
        st.markdown(metric_card("Cảnh báo WARNING", f"{warning_count}"), unsafe_allow_html=True)

    st.divider()

    with st.container(border=True):
        st.markdown("##### Lọc thông tin & Thao tác nhanh")
        c1, c2 = st.columns([3, 1])
        with c1:
            severity_options = sorted(logs["severity"].dropna().unique())
            selected = st.multiselect("Lọc mức độ nghiêm trọng", severity_options, default=severity_options)
        with c2:
            st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
            if st.button("Xóa sạch log cảnh báo", type="secondary", use_container_width=True):
                clear_validation_logs(period)
                st.success("Đã làm sạch lịch sử cảnh báo.")
                st.rerun()

    filtered = logs[logs["severity"].isin(selected)] if selected else logs

    st.markdown("### Log kiểm định chi tiết")
    if filtered.empty:
        st.info("Không có dữ liệu cảnh báo phù hợp với bộ lọc hiện tại.")
    else:
        filtered_disp = filtered.copy()
        filtered_disp["severity"] = filtered_disp["severity"].map({
            "BLOCKER": "BLOCKER",
            "WARNING": "WARNING"
        }).fillna(filtered_disp["severity"])

        filtered_disp = filtered_disp.rename(columns={
            "village_name": "Thôn",
            "severity": "Mức độ",
            "error_type": "Loại lỗi",
            "field_name": "Trường dữ liệu",
            "message": "Nội dung chi tiết",
            "source_file": "Nguồn file Excel"
        })

        with st.container(border=True):
            st.dataframe(
                filtered_disp[["Thôn", "Mức độ", "Loại lỗi", "Trường dữ liệu", "Nội dung chi tiết", "Nguồn file Excel"]],
                use_container_width=True,
                hide_index=True
            )

def page_tasks() -> None:
    st.subheader("Tiến độ nộp báo cáo & nhắc việc")
    period = period_selector()
    tasks = pd.DataFrame(fetch_tasks(period))
    reports = pd.DataFrame(fetch_reports(period))
    if not tasks.empty and not reports.empty:
        tasks = tasks.merge(reports[["village_name", "phone"]], on="village_name", how="left")
    
    if tasks.empty:
        st.info("Chưa có nhiệm vụ báo cáo nào được ghi nhận cho kỳ này.")
        return

    total_tasks = len(tasks)
    completed = int((tasks["status"] == "Hoàn thành").sum())
    pending = int((tasks["status"] == "Chưa nộp").sum())
    late = int((tasks["status"] == "Trễ hạn").sum())
    warnings = int((tasks["status"] == "Cảnh báo").sum())
    
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.markdown(metric_card("Tổng số nhiệm vụ", f"{total_tasks}"), unsafe_allow_html=True)
    with col2:
        st.markdown(metric_card("Đã hoàn thành", f"{completed}"), unsafe_allow_html=True)
    with col3:
        st.markdown(metric_card("Chưa nộp", f"{pending}"), unsafe_allow_html=True)
    with col4:
        st.markdown(metric_card("Nộp trễ hạn", f"{late}"), unsafe_allow_html=True)
    with col5:
        st.markdown(metric_card("Cảnh báo số liệu", f"{warnings}"), unsafe_allow_html=True)

    st.divider()

    st.subheader("Bảng chi tiết tiến độ các thôn")
    tasks_disp = tasks.copy()
    tasks_disp["status"] = tasks_disp["status"].map({
        "Hoàn thành": "Hoàn thành",
        "Chưa nộp": "Chưa nộp",
        "Trễ hạn": "Trễ hạn",
        "Cảnh báo": "Cảnh báo"
    }).fillna(tasks_disp["status"])
    
    tasks_disp = tasks_disp.rename(columns={
        "village_name": "Thôn",
        "task_name": "Nhiệm vụ",
        "period": "Kỳ",
        "due_at": "Hạn nộp",
        "submitted_at": "Thời điểm nộp",
        "status": "Trạng thái",
        "days_late": "Trễ (ngày)",
        "reminder_status": "Nhắc việc",
    })
    
    with st.container(border=True):
        st.dataframe(
            tasks_disp[["Thôn", "Nhiệm vụ", "Kỳ", "Hạn nộp", "Thời điểm nộp", "Trạng thái", "Trễ (ngày)", "Nhắc việc"]],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Trễ (ngày)": st.column_config.NumberColumn(format="%d ngày")
            }
        )

    st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)
    st.subheader("Tin nhắn nhắc việc tự động (Mô phỏng Zalo/SMS OA)")
    
    pending_tasks = tasks[tasks["status"].astype(str).isin(["Chưa nộp", "Trễ hạn", "Cảnh báo"])]
    if pending_tasks.empty:
        st.success("Tuyệt vời! Tất cả các thôn đã nộp báo cáo đúng hạn. Không cần nhắc việc.")
    else:
        for _, row in pending_tasks.iterrows():
            village = row["village_name"]
            task = row["task_name"]
            due = row["due_at"]
            status = row["status"]
            days = row.get("days_late", 0)
            
            # Themed light reminder notifications
            if status == "Chưa nộp":
                st.markdown(
                    f"""
                    <div style="background-color: #fff7ed; border: 1px solid #ffedd5; border-left: 4px solid #f97316; padding: 15px; border-radius: 10px; margin-bottom: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
                        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                            <span style="font-weight: 700; color: #c2410c;">NHẮC NHỞ: QUÁ HẠN NỘP BÁO CÁO</span>
                            <span style="background-color: #ffedd5; color: #c2410c; font-size: 0.75rem; font-weight: bold; padding: 2px 8px; border-radius: 9999px;">ZALO OA THÔN</span>
                        </div>
                        <div style="color: #475569; font-size: 0.95rem; line-height: 1.5;">
                            Kính gửi Đồng chí Trưởng thôn <b>{village}</b>,<br>
                            Nhiệm vụ "<i>{task}</i>" kỳ báo cáo <b>{period}</b> hiện đã quá hạn nộp (Hạn cuối: {due}). Đề nghị đồng chí khẩn trương tổng hợp số liệu và nộp báo cáo Excel lên hệ thống xã.
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
            elif status == "Cảnh báo":
                phone_val = row.get("phone") or "không rõ"
                st.markdown(
                    f"""
                    <div style="background-color: #fefce8; border: 1px solid #fef08a; border-left: 4px solid #eab308; padding: 15px; border-radius: 10px; margin-bottom: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
                        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                            <span style="font-weight: 700; color: #854d0e;">CẢNH BÁO: SAI ĐỊNH DẠNG SỐ ĐIỆN THOẠI</span>
                            <span style="background-color: #fef08a; color: #854d0e; font-size: 0.75rem; font-weight: bold; padding: 2px 8px; border-radius: 9999px;">ZALO OA THÔN</span>
                        </div>
                        <div style="color: #475569; font-size: 0.95rem; line-height: 1.5;">
                            Kính gửi Đồng chí Trưởng thôn <b>{village}</b>,<br>
                            Báo cáo "<i>{task}</i>" kỳ báo cáo <b>{period}</b> đã được nộp nhưng hệ thống phát hiện <b>Số điện thoại người lập ({phone_val}) không đúng định dạng</b> (yêu cầu là 10 chữ số bắt đầu bằng số 0). Đề nghị đồng chí cập nhật lại số điện thoại chính xác để hệ thống hoàn tất đồng bộ.
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
            else:
                st.markdown(
                    f"""
                    <div style="background-color: #fef2f2; border: 1px solid #fee2e2; border-left: 4px solid #ef4444; padding: 15px; border-radius: 10px; margin-bottom: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
                        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                            <span style="font-weight: 700; color: #b91c1c;">BÁO CÁO TRỄ HẠN</span>
                            <span style="background-color: #fee2e2; color: #b91c1c; font-size: 0.75rem; font-weight: bold; padding: 2px 8px; border-radius: 9999px;">HỆ THỐNG XÃ</span>
                        </div>
                        <div style="color: #475569; font-size: 0.95rem; line-height: 1.5;">
                            Thông báo gửi Cán bộ kiểm soát,<br>
                            Đơn vị <b>{village}</b> đã hoàn tất nộp muộn trễ <b>{days} ngày</b>. Hệ thống ghi nhận trạng thái nộp trễ. Đề xuất rà soát lý do chậm trễ của đơn vị.
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )


def page_export() -> None:
    st.subheader("Trung tâm Kết xuất Báo cáo tự động")
    period = period_selector()
    export_dir = BASE_DIR / "exports"
    export_dir.mkdir(exist_ok=True)
    
    st.markdown(
        "Hệ thống tự động liên kết các nguồn số liệu sạch đã được chuẩn hóa để xuất các mẫu văn bản báo cáo hành chính chuẩn mực:"
    )
    
    col1, col2 = st.columns(2)
    with col1:
        with st.container(border=True):
            st.markdown(
                """
                ### Báo cáo Số liệu Excel (.xlsx)
                Báo cáo tổng quan chi tiết và phân rã các chỉ tiêu theo hàng dọc của từng thôn.
                *   **Sheet Số liệu**: Tổng hợp 14 chỉ tiêu CT01–CT14 toàn bộ các thôn.
                *   **Sheet Tiến độ**: Thống kê thời gian nộp và trễ hạn.
                *   **Định dạng**: Đã đóng băng tiêu đề, tự động co giãn cột chuẩn đẹp.
                """
            )
            if st.button("Tạo báo cáo Excel tổng hợp", type="primary", use_container_width=True):
                path = export_dir / f"bao_cao_tong_hop_{period}.xlsx"
                with st.spinner("Đang xử lý xuất dữ liệu..."):
                    generate_excel_report(period, path)
                st.success("Khởi tạo báo cáo Excel thành công!")
                st.download_button(
                    " Tải xuống file Excel (.xlsx)",
                    data=path.read_bytes(),
                    file_name=path.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
    with col2:
        with st.container(border=True):
            st.markdown(
                """
                ### Văn bản tổng hợp Word (.docx)
                Văn bản báo cáo hành chính tổng hợp tình hình kinh tế - xã hội phục vụ công tác họp giao ban.
                *   **Nội dung**: Soạn sẵn lời mở đầu, ngày tháng tổng hợp và các bảng biểu.
                *   **Bảng số liệu**: Hộ nghèo, BHYT, nhân khẩu chi tiết.
                *   **Định dạng**: Đã thiết kế tiêu chuẩn Arial văn phòng hành chính.
                """
            )
            if st.button("Tạo báo cáo Word hành chính", type="primary", use_container_width=True):
                path = export_dir / f"bao_cao_tong_hop_{period}.docx"
                with st.spinner("Đang sinh văn bản hành chính..."):
                    generate_word_report(period, path)
                st.success("Khởi tạo văn bản Word thành công!")
                st.download_button(
                    " Tải xuống file Word (.docx)",
                    data=path.read_bytes(),
                    file_name=path.name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )


def page_assistant() -> None:
    st.subheader("Trợ lý AI Bản demo")
    st.caption("AI Agent truy vấn cơ sở dữ liệu có cấu trúc cấp thôn/xã. Bản nâng cấp tiếp theo sẽ tích hợp mô hình ngôn ngữ lớn LLM + RAG.")
    
    period = period_selector()
    reports = pd.DataFrame(fetch_reports(period))
    villages = pd.DataFrame(fetch_villages())
    tasks = pd.DataFrame(fetch_tasks(period))
    logs = pd.DataFrame(fetch_validation_logs(period))

    if "chat_history" not in st.session_state:
        st.session_state.chat_history = [
            {
                "role": "assistant",
                "content": "Xin chào! Tôi là Trợ lý AI Bà Nà. Tôi có thể giúp gì cho đồng chí? Dưới đây là một số câu hỏi gợi ý đồng chí có thể hỏi tôi về tình hình số liệu các thôn:",
            }
        ]

    st.markdown("<p style='font-size: 0.85rem; font-weight: bold; color: #64748b;'> CÂU HỎI TRUY VẤN NHANH:</p>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    suggestions = [
        "Thôn nào chưa nộp báo cáo?",
        "Thôn nào nộp trễ?",
        "Tổng số hộ nghèo là bao nhiêu?",
        "Tổng số người tham gia BHYT là bao nhiêu?",
        "Có cảnh báo dữ liệu nào không?",
        "Tóm tắt tình hình kỳ này"
    ]
    for idx, sug in enumerate(suggestions):
        col_selector = c1 if idx < 2 else (c2 if idx < 4 else c3)
        if col_selector.button(sug, key=f"sug_btn_{idx}", use_container_width=True):
            st.session_state.pending_user_message = sug
            st.rerun()

    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

    for msg in st.session_state.chat_history:
        role_label = "Trợ lý AI" if msg["role"] == "assistant" else "Người dùng"
        with st.container(border=True):
            st.caption(role_label)
            st.markdown(msg["content"])
            if "df" in msg:
                st.dataframe(msg["df"], use_container_width=True, hide_index=True)

    user_query = st.chat_input("Nhập thắc mắc hoặc câu hỏi về số liệu tại đây...")

    if st.session_state.get("pending_user_message"):
        user_query = st.session_state.pending_user_message
        del st.session_state.pending_user_message

    if user_query:
        st.session_state.chat_history.append({"role": "user", "content": user_query})
        
        q_norm = normalize_text(user_query)
        response_content = ""
        response_df = None
        
        if "chua_nop" in q_norm or ("chua" in q_norm and "nop" in q_norm):
            missing = tasks[tasks["status"] == "Chưa nộp"] if not tasks.empty else pd.DataFrame()
            if missing.empty:
                response_content = "Tất cả các thôn đã nộp báo cáo đầy đủ trong kỳ này."
            else:
                response_content = f"Có **{len(missing)} thôn chưa nộp** báo cáo: {', '.join(missing['village_name'].tolist())}."
        elif "tre" in q_norm or "nop_tre" in q_norm:
            late = tasks[tasks["status"] == "Trễ hạn"] if not tasks.empty else pd.DataFrame()
            if late.empty:
                response_content = "Không có thôn nào nộp báo cáo trễ hạn kỳ này."
            else:
                response_content = f"Có **{len(late)} thôn nộp trễ** hạn: {', '.join(late['village_name'].tolist())}."
        elif "ho_ngheo" in q_norm or "ngheo" in q_norm:
            total = int(reports.get("ct03_poor_households", pd.Series(dtype=float)).sum()) if not reports.empty else 0
            response_content = f"Tổng số hộ nghèo ghi nhận trên toàn xã trong kỳ {period} là: **{total} hộ**."
        elif "bhyt" in q_norm or "bao_hiem" in q_norm or "y_te" in q_norm:
            total = int(reports.get("ct11_health_insurance", pd.Series(dtype=float)).sum()) if not reports.empty else 0
            response_content = f"Tổng số người tham gia BHYT ghi nhận toàn xã kỳ {period} là: **{total} người**."
        elif "du_lieu_thieu" in q_norm or "thieu" in q_norm or "bo_sung" in q_norm:
            missing_df = build_missing_fields_dataframe(period)
            if missing_df.empty:
                response_content = "Không phát hiện dữ liệu thiếu cần bổ sung trong kỳ này."
            else:
                response_content = f"Hệ thống phát hiện **{len(missing_df)} mục dữ liệu thiếu/cần sửa** trong kỳ {period}. Dưới đây là bảng chi tiết:"
                response_df = missing_df
        elif "canh_bao" in q_norm or "loi" in q_norm:
            if logs.empty:
                response_content = "Không ghi nhận bất kỳ lỗi hay cảnh báo dữ liệu nào kỳ này."
            else:
                response_content = f"Hệ thống phát hiện **{len(logs)} cảnh báo dữ liệu** kỳ {period}. Dưới đây là bảng chi tiết:"
                response_df = logs[["village_name", "severity", "error_type", "field_name", "message"]].rename(columns={
                    "village_name": "Thôn",
                    "severity": "Mức độ",
                    "error_type": "Loại lỗi",
                    "field_name": "Trường",
                    "message": "Nội dung cảnh báo"
                })
        else:
            total_villages = len(villages)
            submitted = reports["village_name"].nunique() if not reports.empty else 0
            total_population = int(reports.get("ct02_population", pd.Series(dtype=float)).sum()) if not reports.empty else 0
            total_poor = int(reports.get("ct03_poor_households", pd.Series(dtype=float)).sum()) if not reports.empty else 0
            missing = len(tasks[tasks["status"] == "Chưa nộp"]) if not tasks.empty else 0
            late = len(tasks[tasks["status"] == "Trễ hạn"]) if not tasks.empty else 0
            response_content = (
                f"**Tóm tắt tình hình số liệu kỳ {period}:**\n\n"
                f"- Số thôn đã nộp: **{submitted}/{total_villages}** thôn.\n"
                f"- Số nhân khẩu: **{total_population:,}** người.\n"
                f"- Tổng số hộ nghèo: **{total_poor}** hộ.\n"
                f"- Chưa nộp: **{missing}** thôn | Nộp trễ: **{late}** thôn.\n"
                f"- Cảnh báo dữ liệu: **{len(logs)}** cảnh báo được tìm thấy."
                .replace(",", ".")
            )
            
        ai_msg = {"role": "assistant", "content": response_content}
        if response_df is not None:
            ai_msg["df"] = response_df
        st.session_state.chat_history.append(ai_msg)
        st.rerun()


def page_roadmap() -> None:
    st.subheader("Lộ trình phát triển hệ thống (Roadmap)")
    
    st.markdown(
        """
        <div style="position: relative; margin-left: 20px; border-left: 2px dashed #3b82f6; padding-left: 30px;">
            <div style="margin-bottom: 25px; position: relative;">
                <div style="position: absolute; left: -41px; top: 2px; width: 20px; height: 20px; border-radius: 50%; background: #10b981; border: 4px solid #f1f5f9;"></div>
                <div style="background: #ffffff; padding: 20px; border-radius: 12px; border: 1px solid #cbd5e1; border-left: 5px solid #10b981; box-shadow: 0 4px 12px rgba(0,0,0,0.05);">
                    <span style="background-color: #d1fae5; color: #065f46; font-size: 0.75rem; font-weight: bold; padding: 4px 10px; border-radius: 9999px; text-transform: uppercase;">Giai đoạn 1 (Hiện tại - MVP)</span>
                    <h4 style="margin: 10px 0 5px 0; color: #0f172a !important; font-weight: 700;">Hoàn thiện luồng nộp và kiểm soát chéo (Excel-first)</h4>
                    <p style="margin: 0; color: #475569; font-size: 0.95rem; line-height: 1.6;">
                        • Thiết lập biểu mẫu chỉ tiêu chuẩn mực <b>CT01–CT14</b> của xã gửi cho trưởng thôn.<br>
                        • Đồng bộ và kế thừa nhanh dữ liệu từ Google Drive của phòng Văn hóa - Xã hội.<br>
                        • AI tự động kiểm định chất lượng chéo số liệu (phát hiện số âm, lỗi nhân khẩu/hộ, trẻ em dưới 16 tuổi vượt quá nhân khẩu, v.v.).<br>
                        • Quản lý tiến độ nộp và nhắc nhở gửi trễ tự động, kết xuất báo cáo chuẩn Excel/Word phục vụ giao ban.
                    </p>
                </div>
            </div>
            <div style="margin-bottom: 25px; position: relative;">
                <div style="position: absolute; left: -41px; top: 2px; width: 20px; height: 20px; border-radius: 50%; background: #3b82f6; border: 4px solid #f1f5f9;"></div>
                <div style="background: #ffffff; padding: 20px; border-radius: 12px; border: 1px solid #cbd5e1; border-left: 5px solid #3b82f6; box-shadow: 0 4px 12px rgba(0,0,0,0.05);">
                    <span style="background-color: #dbeafe; color: #1e40af; font-size: 0.75rem; font-weight: bold; padding: 4px 10px; border-radius: 9999px; text-transform: uppercase;">Giai đoạn 2 (Tiếp theo)</span>
                    <h4 style="margin: 10px 0 5px 0; color: #0f172a !important; font-weight: 700;">Tích hợp AI tạo sinh LLM, RAG và Phân quyền sử dụng</h4>
                    <p style="margin: 0; color: #475569; font-size: 0.95rem; line-height: 1.6;">
                        • <b>Trợ lý ảo thông minh (LLM + RAG)</b>: Cho phép cán bộ truy vấn số liệu tự nhiên bằng ngôn ngữ nói, tự động soạn thảo văn bản giải trình hoặc tóm tắt số liệu phức tạp.<br>
                        • <b>Công nghệ OCR quét ảnh</b>: Tự động số hóa báo cáo từ ảnh chụp bản viết tay của trưởng thôn.<br>
                        • <b>Tích hợp tin nhắn Zalo OA Webhook</b>: Trực tiếp nhận file báo cáo Excel thông qua tài khoản Zalo OA của trưởng thôn.<br>
                        • Nâng cấp hệ quản trị lên <b>PostgreSQL</b> và xây dựng ứng dụng web phân quyền chi tiết (Lãnh đạo xã, Cán bộ phòng ban, Trưởng/phó thôn).
                    </p>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


with st.sidebar:
    st.markdown(
        """
        <div style="margin-bottom: 1.5rem; text-align: center; border-bottom: 2px solid #991b1b; padding-bottom: 15px;">
            <h2 style="font-weight: 800; font-size: 1.2rem; color: #b91c1c; margin: 0; text-transform: uppercase; letter-spacing: 0.5px;">
                UBND XÃ BÀ NÀ
            </h2>
            <span style="color: #475569; font-size: 0.75rem; font-weight: 600; display: block; margin-top: 4px;">SmartLink Portal • v1.0</span>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    if st.button("Khởi tạo lại Database từ Drive", use_container_width=True):
        if DB_PATH.exists():
            DB_PATH.unlink()
        seed_demo_data(force=True)
        st.success("Đã thiết lập lại cơ sở dữ liệu về mặc định!")
        st.rerun()
        
    st.caption("Cơ chế hoạt động: Nhập Excel thôn/Drive → Auto-Validator chéo → Lưu SQLite → Tổng hợp Dashboard & báo cáo hành chính.")

# ROOT-LEVEL HEADER AND HORIZONTAL NAVIGATION BAR
header()

menu_options = [
    "Tổng quan", 
    "Bảng điều khiển", 
    "Upload báo cáo", 
    "Kiểm tra dữ liệu", 
    "Nhiệm vụ & nhắc việc", 
    "Xuất báo cáo", 
    "Lộ trình"
]
selected = st.radio("Chọn chức năng hệ thống", menu_options, horizontal=True, label_visibility="collapsed")

# STRICT ROUTING LOGIC
if selected == "Tổng quan":
    page_overview()
elif selected == "Bảng điều khiển":
    page_dashboard_charts()
elif selected == "Upload báo cáo":
    page_upload()
elif selected == "Kiểm tra dữ liệu":
    page_quality()
elif selected == "Nhiệm vụ & nhắc việc":
    page_tasks()
elif selected == "Xuất báo cáo":
    page_export()
else:
    page_roadmap()

# FLOATING AI CHAT DIALOG
# The initial click opens a real Streamlit dialog. Widgets inside the dialog
# rerun only the dialog fragment instead of rebuilding and scrolling the page.


def _initial_chat_history() -> list[dict]:
    return [
        {
            "role": "assistant",
            "content": (
                "Xin chào! Tôi là Trợ lý AI Bà Nà. "
                "Đồng chí cần kiểm tra thông tin nào?"
            ),
        }
    ]


def _build_ai_response(period: str, user_query: str) -> tuple[str, pd.DataFrame | None]:
    """Create a safe, read-only answer from the structured local database."""
    reports = pd.DataFrame(fetch_reports(period))
    villages = pd.DataFrame(fetch_villages())
    tasks = pd.DataFrame(fetch_tasks(period))
    logs = pd.DataFrame(fetch_validation_logs(period))

    q_norm = normalize_text(user_query)
    response_df: pd.DataFrame | None = None

    if "chua_nop" in q_norm or ("chua" in q_norm and "nop" in q_norm):
        missing = tasks[tasks["status"] == "Chưa nộp"] if not tasks.empty else pd.DataFrame()
        if missing.empty:
            response = "Tất cả các thôn đã nộp báo cáo đầy đủ trong kỳ này."
        else:
            response = (
                f"Có **{len(missing)} thôn chưa nộp** báo cáo: "
                f"{', '.join(missing['village_name'].astype(str).tolist())}."
            )

    elif "tre" in q_norm or "nop_tre" in q_norm:
        late = tasks[tasks["status"] == "Trễ hạn"] if not tasks.empty else pd.DataFrame()
        if late.empty:
            response = "Không có thôn nào nộp báo cáo trễ hạn kỳ này."
        else:
            response = (
                f"Có **{len(late)} thôn nộp trễ**: "
                f"{', '.join(late['village_name'].astype(str).tolist())}."
            )

    elif "ho_ngheo" in q_norm or "ngheo" in q_norm:
        total = (
            int(reports.get("ct03_poor_households", pd.Series(dtype=float)).sum())
            if not reports.empty
            else 0
        )
        response = f"Tổng số hộ nghèo toàn xã trong kỳ **{period} là {total} hộ**."

    elif "bhyt" in q_norm or "bao_hiem" in q_norm or "y_te" in q_norm:
        total = (
            int(reports.get("ct11_health_insurance", pd.Series(dtype=float)).sum())
            if not reports.empty
            else 0
        )
        response = f"Tổng số người tham gia BHYT kỳ **{period} là {total} người**."

    elif "du_lieu_thieu" in q_norm or "bo_sung" in q_norm or (
        "thieu" in q_norm and "chua_nop" not in q_norm
    ):
        missing_df = build_missing_fields_dataframe(period)
        if missing_df.empty:
            response = "Không phát hiện dữ liệu thiếu cần bổ sung trong kỳ này."
        else:
            response = (
                f"Hệ thống phát hiện **{len(missing_df)} mục dữ liệu thiếu/cần sửa** "
                f"trong kỳ {period}."
            )
            response_df = missing_df.head(30)

    elif "canh_bao" in q_norm or "loi" in q_norm:
        if logs.empty:
            response = "Không ghi nhận lỗi hoặc cảnh báo dữ liệu nào trong kỳ này."
        else:
            response = f"Hệ thống ghi nhận **{len(logs)} cảnh báo dữ liệu** trong kỳ {period}."
            wanted = ["village_name", "severity", "field_name", "message"]
            available = [column for column in wanted if column in logs.columns]
            response_df = logs[available].rename(
                columns={
                    "village_name": "Thôn",
                    "severity": "Mức độ",
                    "field_name": "Trường",
                    "message": "Nội dung",
                }
            ).head(30)

    else:
        total_villages = len(villages)
        submitted = reports["village_name"].nunique() if not reports.empty else 0
        total_population = (
            int(reports.get("ct02_population", pd.Series(dtype=float)).sum())
            if not reports.empty
            else 0
        )
        total_poor = (
            int(reports.get("ct03_poor_households", pd.Series(dtype=float)).sum())
            if not reports.empty
            else 0
        )
        missing_count = (
            len(tasks[tasks["status"] == "Chưa nộp"]) if not tasks.empty else 0
        )
        late_count = (
            len(tasks[tasks["status"] == "Trễ hạn"]) if not tasks.empty else 0
        )
        response = (
            f"**Tóm tắt kỳ {period}:**\n\n"
            f"- Đã nộp: **{submitted}/{total_villages} thôn**.\n"
            f"- Tổng nhân khẩu: **{total_population:,} người**.\n"
            f"- Tổng số hộ nghèo: **{total_poor} hộ**.\n"
            f"- Chưa nộp: **{missing_count} thôn**; nộp trễ: **{late_count} thôn**.\n"
            f"- Cảnh báo dữ liệu: **{len(logs)}**."
        ).replace(",", ".")

    return response, response_df


@st.dialog(
    "Trợ lý AI Bản demo",
    width="small",
    dismissible=True,
    on_dismiss="ignore",
)
def open_ai_chat_dialog() -> None:
    """
    Compact assistant window.

    Because st.dialog inherits st.fragment behavior, interactions inside this
    function rerun only the dialog instead of rerunning the complete dashboard.
    """
    if "assistant_chat_history" not in st.session_state:
        st.session_state.assistant_chat_history = _initial_chat_history()

    periods = list_periods()
    if not periods:
        periods = ["2026-Q2"]

    chat_period = st.selectbox(
        "Chọn kỳ báo cáo",
        periods,
        key="assistant_chat_period",
    )

    st.caption("Câu hỏi truy vấn nhanh")
    quick_prompt = None
    col1, col2 = st.columns(2)

    with col1:
        if st.button(
            "Thôn nào chưa nộp?",
            key="assistant_missing_button",
            use_container_width=True,
        ):
            quick_prompt = "Thôn nào chưa nộp báo cáo?"

        if st.button(
            "Tổng số hộ nghèo?",
            key="assistant_poor_button",
            use_container_width=True,
        ):
            quick_prompt = "Tổng số hộ nghèo là bao nhiêu?"

    with col2:
        if st.button(
            "Thôn nào nộp trễ?",
            key="assistant_late_button",
            use_container_width=True,
        ):
            quick_prompt = "Thôn nào nộp trễ?"

        if st.button(
            "Có cảnh báo dữ liệu?",
            key="assistant_warning_button",
            use_container_width=True,
        ):
            quick_prompt = "Có cảnh báo dữ liệu nào không?"

    # Declare the conversation area before the input, then populate it after
    # processing the current prompt. This shows the response immediately and
    # avoids an additional st.rerun().
    history_box = st.container(height=285, key="ai_chat_history")

    typed_prompt = st.chat_input(
        "Nhập câu hỏi về số liệu...",
        key="assistant_chat_input",
    )
    prompt = typed_prompt or quick_prompt

    if prompt:
        st.session_state.assistant_chat_history.append(
            {"role": "user", "content": prompt}
        )
        response, response_df = _build_ai_response(chat_period, prompt)
        assistant_message = {"role": "assistant", "content": response}
        if response_df is not None:
            assistant_message["df"] = response_df
        st.session_state.assistant_chat_history.append(assistant_message)

    with history_box:
        for message in st.session_state.assistant_chat_history:
            role_label = "Trợ lý AI" if message["role"] == "assistant" else "Người dùng"
            with st.container(border=True):
                st.caption(role_label)
                st.markdown(message["content"])
                if message.get("df") is not None:
                    st.dataframe(
                        message["df"],
                        use_container_width=True,
                        hide_index=True,
                        height=180,
                    )

    left_action, right_action = st.columns([1, 1])
    with left_action:
        st.caption("Dữ liệu truy vấn ở chế độ chỉ đọc.")
    with right_action:
        if st.button(
            " Xóa hội thoại",
            key="assistant_clear_history",
            use_container_width=True,
        ):
            st.session_state.assistant_chat_history = _initial_chat_history()


# The button's first click opens the modal. Once open, all controls inside the
# modal rerun only the dialog fragment, so the underlying page stays in place.
if st.button(
    "Trợ lý AI",
    key="open_ai_chat",
    help="Mở Trợ lý AI Bà Nà",
):
    open_ai_chat_dialog()
