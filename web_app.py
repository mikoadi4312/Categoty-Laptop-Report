import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook


PROJECT_ROOT = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from config.database import get_connection  # noqa: E402
from etl.upload_sales import upload_sales  # noqa: E402
from etl.upload_stock import upload_stock  # noqa: E402
from reports.generate_laptop_report import generate_report  # noqa: E402


APP_DATA_DIR = PROJECT_ROOT / "app_data"
UPLOAD_DIR = APP_DATA_DIR / "uploads"
REPORT_DIR = APP_DATA_DIR / "reports"
ALLOWED_TYPES = ["xlsx", "xlsm", "xls", "csv"]


def setup_page() -> None:
    st.set_page_config(
        page_title="Laptop Daily Report",
        page_icon=None,
        layout="wide",
        initial_sidebar_state="collapsed",
    )
    st.markdown(
        """
        <style>
            :root {
                --navy: #0b2f66;
                --blue: #2f6fed;
                --ink: #000000;
                --muted: #475569;
                --line: #cfe0f5;
                --soft: #f4f9ff;
                --page: #eaf4ff;
                --card: #ffffff;
                --button: #0b2f66;
                --button-hover: #123f84;
                --accent-soft: #dcecff;
                --green: #0f8f5f;
            }

            .stApp {
                background:
                    radial-gradient(circle at top left, rgba(47, 111, 237, 0.14), transparent 34%),
                    linear-gradient(180deg, #eaf4ff 0%, #f5fbff 46%, #ffffff 100%);
                color: var(--ink);
                font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
            }

            .block-container {
                padding-top: 1rem;
                padding-bottom: 2.2rem;
                max-width: 1380px;
            }

            .topbar {
                display: flex;
                align-items: center;
                justify-content: space-between;
                margin-bottom: 18px;
                color: var(--navy);
            }

            .brand-lockup {
                display: flex;
                align-items: center;
                gap: 10px;
                font-weight: 850;
                font-size: 18px;
            }

            .brand-mark {
                width: 32px;
                height: 32px;
                border-radius: 7px;
                background: var(--button);
                color: #ffffff;
                display: inline-flex;
                align-items: center;
                justify-content: center;
                font-weight: 950;
                box-shadow: 0 6px 16px rgba(11, 47, 102, 0.18);
            }

            .topbar-status {
                color: var(--muted);
                font-size: 13px;
                font-weight: 650;
                background: rgba(255, 255, 255, 0.78);
                border: 1px solid var(--line);
                border-radius: 999px;
                padding: 8px 12px;
            }

            .hero {
                display: grid;
                grid-template-columns: minmax(0, 1.05fr) minmax(360px, 0.95fr);
                gap: 24px;
                align-items: stretch;
                margin-bottom: 22px;
            }

            .hero-copy {
                background: rgba(255, 255, 255, 0.9);
                border: 1px solid var(--line);
                border-radius: 8px;
                padding: 30px;
                box-shadow: 0 18px 42px rgba(11, 47, 102, 0.08);
            }

            .eyebrow {
                display: inline-flex;
                align-items: center;
                color: #000000;
                background: var(--accent-soft);
                border: 1px solid #b8d4f8;
                border-radius: 999px;
                padding: 7px 11px;
                font-size: 12px;
                line-height: 1;
                font-weight: 850;
                margin-bottom: 16px;
            }

            .hero-copy h1 {
                color: #000000;
                font-size: 42px;
                line-height: 1.07;
                margin: 0 0 14px 0;
                letter-spacing: 0;
                font-weight: 900;
            }

            .hero-copy p {
                margin: 0;
                color: var(--muted);
                font-size: 16px;
                line-height: 1.55;
                max-width: 620px;
            }

            .hero-actions {
                display: flex;
                gap: 12px;
                flex-wrap: wrap;
                margin-top: 22px;
            }

            .action-chip {
                border: 1px solid var(--line);
                background: #f4f9ff;
                border-radius: 7px;
                padding: 10px 13px;
                color: #000000;
                font-size: 13px;
                font-weight: 750;
            }

            .sheet-visual {
                background: rgba(255, 255, 255, 0.92);
                border-radius: 8px;
                padding: 18px;
                min-height: 282px;
                color: var(--ink);
                box-shadow: 0 18px 42px rgba(11, 47, 102, 0.08);
                border: 1px solid var(--line);
            }

            .sheet-toolbar {
                display: flex;
                align-items: center;
                justify-content: space-between;
                margin-bottom: 14px;
                color: var(--navy);
                font-size: 12px;
                font-weight: 750;
            }

            .toolbar-dots {
                display: flex;
                gap: 6px;
            }

            .toolbar-dot {
                width: 9px;
                height: 9px;
                border-radius: 50%;
                background: var(--blue);
                opacity: .92;
            }

            .sheet-grid {
                background: white;
                border-radius: 7px;
                overflow: hidden;
                border: 1px solid var(--line);
            }

            .sheet-row {
                display: grid;
                grid-template-columns: 1.1fr .75fr .75fr .75fr;
                min-height: 36px;
                border-bottom: 1px solid #e7edf4;
            }

            .sheet-row:last-child {
                border-bottom: none;
            }

            .sheet-cell {
                padding: 10px 12px;
                color: var(--ink);
                border-right: 1px solid #e7edf4;
                font-size: 12px;
                font-weight: 650;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
            }

            .sheet-cell:last-child {
                border-right: none;
            }

            .sheet-head .sheet-cell {
                background: #eff4ff;
                color: var(--navy);
                font-weight: 850;
            }

            .sheet-total .sheet-cell {
                background: #e6f1ff;
                color: #000000;
                font-weight: 900;
            }

            .sheet-accent {
                color: var(--green);
            }

            .section-title {
                font-size: 16px;
                font-weight: 850;
                color: #000000;
                margin: 12px 0 8px;
            }

            div[data-testid="stMetric"] {
                background: rgba(255, 255, 255, 0.9);
                border: 1px solid var(--line);
                border-radius: 8px;
                padding: 12px 14px;
                box-shadow: 0 10px 28px rgba(11, 47, 102, 0.06);
                min-height: 86px;
                overflow: hidden;
            }

            div[data-testid="stMetric"] label {
                color: var(--muted);
                font-weight: 750;
            }

            div[data-testid="stMetric"] [data-testid="stMetricValue"] {
                font-size: 24px;
                line-height: 1.15;
                color: #000000;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
            }

            div[data-testid="stMetric"] [data-testid="stMetricLabel"] {
                font-size: 12px;
                line-height: 1.2;
                white-space: nowrap;
            }

            div[data-testid="stFileUploader"] section {
                border: 1px dashed #9fc2ee;
                background: #ffffff;
                border-radius: 8px;
            }

            div[data-testid="stFileUploader"] * {
                color: var(--ink);
            }

            div[data-testid="stFileUploader"] section:hover {
                border-color: var(--blue);
                background: #f4f9ff;
            }

            div[data-testid="stTabs"] button {
                font-weight: 850;
                color: #000000 !important;
            }

            div[data-testid="stTabs"] button * {
                color: #000000;
            }

            div[data-testid="stTabs"] button[role="tab"],
            div[data-testid="stTabs"] button[role="tab"] p,
            div[data-testid="stTabs"] button[role="tab"] span,
            button[data-baseweb="tab"],
            button[data-baseweb="tab"] p,
            button[data-baseweb="tab"] span {
                color: #000000 !important;
            }

            div[data-testid="stTabs"] button[aria-selected="true"],
            div[data-testid="stTabs"] button[aria-selected="true"] p,
            div[data-testid="stTabs"] button[aria-selected="true"] span,
            button[data-baseweb="tab"][aria-selected="true"],
            button[data-baseweb="tab"][aria-selected="true"] p,
            button[data-baseweb="tab"][aria-selected="true"] span {
                color: #000000 !important;
            }

            div[data-testid="stTabs"] [data-baseweb="tab-highlight"] {
                background-color: var(--button) !important;
            }

            div[data-baseweb="input"],
            div[data-baseweb="select"],
            div[data-baseweb="popover"] div,
            div[data-testid="stDateInput"] div,
            div[data-testid="stTextInput"] div {
                background-color: #ffffff !important;
                color: #000000 !important;
                border-color: #9fc2ee !important;
            }

            input,
            textarea,
            input::placeholder,
            textarea::placeholder {
                color: #000000 !important;
                background-color: #ffffff !important;
            }

            div[data-testid="stTextInput"] input,
            div[data-testid="stDateInput"] input {
                color: #000000 !important;
                background-color: #ffffff !important;
                -webkit-text-fill-color: #000000 !important;
            }

            div[data-testid="stTextInput"] label,
            div[data-testid="stDateInput"] label,
            div[data-testid="stFileUploader"] label,
            div[data-testid="stFileUploader"] p,
            div[data-testid="stFileUploader"] span,
            div[data-testid="stFileUploader"] small {
                color: #000000 !important;
            }

            .stButton button, .stDownloadButton button {
                border-radius: 6px;
                min-height: 42px;
                font-weight: 850;
                border: 1px solid var(--button);
                background: #ffffff;
                color: #000000;
            }

            .stButton button[kind="primary"] {
                background: var(--button);
                border-color: var(--button);
                color: #ffffff;
                box-shadow: 0 10px 18px rgba(11, 47, 102, 0.18);
            }

            .stButton button[kind="primary"] *,
            .stDownloadButton button * {
                color: #ffffff !important;
            }

            .stButton button[kind="primary"]:hover,
            .stDownloadButton button:hover {
                background: var(--button-hover);
                border-color: var(--button-hover);
                color: #ffffff;
            }

            .stDownloadButton button {
                background: var(--button);
                border-color: var(--button);
                color: #ffffff;
                box-shadow: 0 10px 18px rgba(11, 47, 102, 0.18);
            }

            div[data-testid="stFileUploader"] button,
            div[data-testid="stFileUploader"] button[kind],
            button[data-testid="stBaseButton-secondary"] {
                background: var(--button) !important;
                border-color: var(--button) !important;
                color: #ffffff !important;
                border-radius: 6px !important;
                font-weight: 850 !important;
            }

            div[data-testid="stFileUploader"] button *,
            button[data-testid="stBaseButton-secondary"] * {
                color: #ffffff !important;
            }

            div[data-testid="stDataFrame"] {
                background: white;
                border-radius: 8px;
            }

            .status-card {
                background: rgba(255, 255, 255, 0.9);
                border: 1px solid var(--line);
                border-radius: 8px;
                padding: 14px 16px;
            }

            .panel-title {
                font-size: 18px;
                font-weight: 900;
                color: #000000;
                margin: 0 0 4px 0;
            }

            .panel-subtitle {
                color: var(--muted);
                font-size: 13px;
                margin-bottom: 14px;
            }

            div[data-testid="stAlert"] {
                border-radius: 8px;
            }

            @media (max-width: 860px) {
                .hero {
                    grid-template-columns: 1fr;
                }

                .hero-copy h1 {
                    font-size: 34px;
                }

                .sheet-visual {
                    min-height: 240px;
                }
            }
        </style>
        """,
        unsafe_allow_html=True,
    )


def init_dirs() -> None:
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    REPORT_DIR.mkdir(parents=True, exist_ok=True)


def clean_name(name: str) -> str:
    allowed = []
    for char in Path(name).name:
        allowed.append(char if char.isalnum() or char in "._- " else "_")
    cleaned = "".join(allowed).strip()
    return cleaned or "uploaded_file"


def save_uploaded_file(uploaded_file, prefix: str) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = UPLOAD_DIR / f"{prefix}_{timestamp}_{clean_name(uploaded_file.name)}"
    path.write_bytes(uploaded_file.getbuffer())
    return path


def db_status() -> tuple[bool, str]:
    try:
        conn = get_connection()
        try:
            with conn.cursor() as cur:
                cur.execute("SELECT 1")
                cur.fetchone()
            return True, "Connected"
        finally:
            conn.close()
    except Exception as exc:
        return False, str(exc)


def latest_sales_date() -> date:
    try:
        conn = get_connection()
        try:
            with conn.cursor() as cur:
                cur.execute("SELECT MAX(sale_date) FROM fact_sales")
                value = cur.fetchone()[0]
                return value or date.today()
        finally:
            conn.close()
    except Exception:
        return date.today()


def latest_stock_date() -> date | None:
    try:
        conn = get_connection()
        try:
            with conn.cursor() as cur:
                cur.execute("SELECT MAX(stock_date) FROM fact_stock")
                return cur.fetchone()[0]
        finally:
            conn.close()
    except Exception:
        return None


def default_report_date() -> date:
    stock_date = latest_stock_date()
    if stock_date:
        return stock_date
    sales_date = latest_sales_date()
    return sales_date + timedelta(days=1)


def app_header() -> None:
    st.markdown(
        """
        <div class="topbar">
            <div class="brand-lockup">
                <span class="brand-mark">L</span>
                <span>Category Laptop</span>
            </div>
            <div class="topbar-status">PostgreSQL based daily report</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def show_database_status() -> None:
    ok, message = db_status()
    latest_date = latest_sales_date() if ok else None
    stock_date = latest_stock_date() if ok else None
    col1, col2, col3, col4 = st.columns([1, 1, 1.25, 1.25])
    with col1:
        st.metric("Database", "Online" if ok else "Offline")
    with col2:
        st.metric("Category", "LAPTOP")
    with col3:
        st.metric("Latest Sales", latest_date.strftime("%Y-%m-%d") if latest_date else "-")
    with col4:
        st.metric("Latest Stock", stock_date.strftime("%Y-%m-%d") if stock_date else "-")
    if ok:
        st.success("PostgreSQL connection is ready.")
    else:
        st.error(f"Database connection failed: {message}")


def sales_tab(uploaded_by: str) -> None:
    st.markdown('<div class="section-title">Revenue Upload</div>', unsafe_allow_html=True)
    sales_file = st.file_uploader("REVENUEERA file", type=ALLOWED_TYPES, key="sales_file")

    upload_clicked = st.button("Upload Revenue", type="primary", use_container_width=True)

    if not sales_file and upload_clicked:
        st.warning("Choose a REVENUEERA file first.")
        return

    if sales_file and upload_clicked:
        path = save_uploaded_file(sales_file, "sales")
        with st.spinner("Uploading sales data..."):
            try:
                inserted, updated = upload_sales(path, uploaded_by)
                st.success(f"Sales uploaded. Inserted: {inserted:,}. Updated: {updated:,}.")
            except Exception as exc:
                st.error(f"Sales upload failed: {exc}")


def stock_tab(uploaded_by: str, stock_date: date) -> None:
    st.markdown('<div class="section-title">Stock Upload</div>', unsafe_allow_html=True)
    stock_file = st.file_uploader("GeneralInventory file", type=ALLOWED_TYPES, key="stock_file")

    upload_clicked = st.button("Upload Stock", type="primary", use_container_width=True)

    if not stock_file and upload_clicked:
        st.warning("Choose a GeneralInventory file first.")
        return

    if stock_file and upload_clicked:
        path = save_uploaded_file(stock_file, "stock")
        with st.spinner("Uploading stock data..."):
            try:
                inserted, updated = upload_stock(path, uploaded_by, stock_date)
                st.success(f"Stock uploaded. Inserted: {inserted:,}. Updated: {updated:,}.")
            except Exception as exc:
                st.error(f"Stock upload failed: {exc}")


def report_tab(report_date: date) -> None:
    st.markdown('<div class="section-title">Excel Report</div>', unsafe_allow_html=True)
    output_name = f"LAPTOP_REPORT_{report_date:%Y%m%d}.xlsx"
    output_path = REPORT_DIR / output_name

    if st.button("Generate Excel Report", type="primary", use_container_width=True):
        with st.spinner("Generating Excel report..."):
            try:
                generate_report(report_date, output_path)
                st.session_state["latest_report"] = str(output_path)
                st.success(f"Report generated: {output_name}")
            except Exception as exc:
                st.error(f"Report generation failed: {exc}")

    latest = st.session_state.get("latest_report")
    if latest and Path(latest).exists():
        path = Path(latest)
        st.download_button(
            label="Download Excel Report",
            data=path.read_bytes(),
            file_name=path.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.markdown('<div class="status-card">No report generated in this session.</div>', unsafe_allow_html=True)


def generated_report_files() -> list[Path]:
    if not REPORT_DIR.exists():
        return []
    return sorted(REPORT_DIR.glob("*.xlsx"), key=lambda path: path.stat().st_mtime, reverse=True)


def preview_report_data(path: Path, max_rows: int = 18) -> pd.DataFrame:
    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        ws = wb["LAPTOP REPORT"] if "LAPTOP REPORT" in wb.sheetnames else wb[wb.sheetnames[0]]
        headers = [ws.cell(3, col).value for col in range(1, ws.max_column + 1)]
        headers = [header if header else f"Column {idx}" for idx, header in enumerate(headers, start=1)]
        rows = []
        for row_idx in range(4, min(ws.max_row, max_rows + 3) + 1):
            row = [ws.cell(row_idx, col).value for col in range(1, len(headers) + 1)]
            if any(value is not None for value in row):
                rows.append(row)
        return pd.DataFrame(rows, columns=headers)
    finally:
        wb.close()


def generated_data_panel(report_date: date) -> None:
    st.markdown('<div class="section-title">Generated Data</div>', unsafe_allow_html=True)
    output_name = f"LAPTOP_REPORT_{report_date:%Y%m%d}.xlsx"
    output_path = REPORT_DIR / output_name

    with st.container(border=True):
        top_left, top_right = st.columns([1.3, 1])
        with top_left:
            st.markdown("**Laptop report output**")
            st.caption(f"Report date: {report_date:%Y-%m-%d} | Sales H-1: {(report_date - timedelta(days=1)):%Y-%m-%d}")
        with top_right:
            generate_clicked = st.button("Generate Excel", type="primary", use_container_width=True)

        if generate_clicked:
            with st.spinner("Generating Excel report..."):
                try:
                    generate_report(report_date, output_path)
                    st.session_state["latest_report"] = str(output_path)
                    st.success(f"Report generated: {output_name}")
                except Exception as exc:
                    st.error(f"Report generation failed: {exc}")

        files = generated_report_files()
        if not files:
            st.markdown(
                '<div class="status-card">Belum ada report yang dibuat. Upload file lalu klik Generate Excel.</div>',
                unsafe_allow_html=True,
            )
            return

        latest_session = st.session_state.get("latest_report")
        default_index = 0
        if latest_session:
            for idx, file_path in enumerate(files):
                if str(file_path) == latest_session:
                    default_index = idx
                    break

        selected = st.selectbox(
            "Generated reports",
            files,
            index=default_index,
            format_func=lambda path: f"{path.name} ({datetime.fromtimestamp(path.stat().st_mtime):%Y-%m-%d %H:%M})",
        )

        download_left, download_right = st.columns([1, 1])
        with download_left:
            st.download_button(
                label="Download Excel",
                data=selected.read_bytes(),
                file_name=selected.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with download_right:
            st.metric("File Size", f"{selected.stat().st_size / 1024:.1f} KB")

        try:
            preview_df = preview_report_data(selected)
            st.markdown("**Preview: LAPTOP REPORT**")
            st.dataframe(preview_df, use_container_width=True, hide_index=True)
        except Exception as exc:
            st.warning(f"Preview not available: {exc}")


def upload_workflow_panel(uploaded_by: str, target_date: date) -> None:
    with st.container(border=True):
        st.markdown('<div class="panel-title">Upload Data</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="panel-subtitle">Upload REVENUEERA and GeneralInventory in one place before generating the report.</div>',
            unsafe_allow_html=True,
        )
        sales_tab(uploaded_by)
        st.divider()
        stock_tab(uploaded_by, target_date)


def main() -> None:
    init_dirs()
    setup_page()
    app_header()

    st.write("")
    suggested_date = default_report_date()
    show_database_status()
    st.write("")

    left, right = st.columns([1, 1])
    with left:
        uploaded_by = st.text_input("Uploaded by", value="Yusuf")
    with right:
        target_date = st.date_input("Report / stock date", value=suggested_date, format="YYYY-MM-DD")
        st.caption(f"Revenue uses H-1 sales date: {(target_date - timedelta(days=1)):%Y-%m-%d}")

    upload_col, generated_col = st.columns([0.78, 1.62], gap="large")
    with upload_col:
        upload_workflow_panel(uploaded_by, target_date)
    with generated_col:
        generated_data_panel(target_date)


if __name__ == "__main__":
    main()
