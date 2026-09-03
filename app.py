"""
Streamlit UI for Financial Automation Report Generator.
Phase 1: Basic UI with file upload and report generation.
"""

import os
import tempfile

import streamlit as st
from pathlib import Path
from streamlit_backend import FileHandler, PipelineOrchestrator, StreamlitLogger, ExcelPreviewHandler
from streamlit_config import ConfigManager, AppConfig

# Page configuration
st.set_page_config(
    page_title="Financial Automation",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Global typography — Inter for everything
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    /* Pull the whole app up by reducing Streamlit's default top padding */
    .block-container {
        padding-top: 1.2rem !important;
    }

    html, body, [class*="css"], h2, h3, h4, button, input, label, p {
        font-family: 'Inter', sans-serif !important;
        font-size: 15px !important;
    }

    h1 {
        font-family: 'Inter', sans-serif !important;
    }


    /* Section headers */
    h2 { font-size: 1.3rem !important; }

    /* Sub-headers */
    h3 { font-size: 1.1rem !important; }

    /* Buttons */
    .stButton > button {
        font-size: 15px !important;
    }

    /* Softer red for primary buttons and the download button */
    .stButton > button[kind="primary"],
    [data-testid="stDownloadButton"] > button {
        background-color: #c0392b !important;
        border-color: #c0392b !important;
        color: #fff !important;
    }
    .stButton > button[kind="primary"]:hover,
    [data-testid="stDownloadButton"] > button:hover {
        background-color: #a93226 !important;
        border-color: #a93226 !important;
    }

    /* Muted success (green) alerts */
    [data-testid="stAlert"][kind="success"],
    div[data-baseweb="notification"][kind="positive"] {
        background-color: #dceedd !important;
        border-left-color: #5a8a5e !important;
        color: #2d4f30 !important;
    }
    [data-testid="stAlert"][kind="success"] svg,
    div[data-baseweb="notification"][kind="positive"] svg {
        fill: #5a8a5e !important;
    }

    /* Muted warning (yellow) alerts */
    [data-testid="stAlert"][kind="warning"],
    div[data-baseweb="notification"][kind="warning"] {
        background-color: #faf0d0 !important;
        border-left-color: #b89a2a !important;
        color: #5a4a10 !important;
    }
    [data-testid="stAlert"][kind="warning"] svg,
    div[data-baseweb="notification"][kind="warning"] svg {
        fill: #b89a2a !important;
    }
    /* Muted info (blue) alerts */
    [data-testid="stAlert"][kind="info"],
    div[data-baseweb="notification"][kind="info"] {
        background-color: #eef1f7 !important;
        border-left-color: #6b7fa0 !important;
        color: #2c3a52 !important;
    }
    [data-testid="stAlert"][kind="info"] svg,
    div[data-baseweb="notification"][kind="info"] svg {
        fill: #6b7fa0 !important;
    }

    /* File uploader — replace Streamlit's accent red with muted green */

    /* "Browse files" button — solid green to match the Add button */
    [data-testid="stFileUploaderDropzoneInput"] + div button,
    [data-testid="baseButton-secondary"] {
        background-color: #5a8a5e !important;
        border-color: #5a8a5e !important;
        color: #fff !important;
    }
    [data-testid="baseButton-secondary"]:hover {
        background-color: #3d6b41 !important;
        border-color: #3d6b41 !important;
        color: #fff !important;
    }

    /* Dropzone */
    [data-testid="stFileUploaderDropzone"] {
        border-color: #8a9bb5 !important;
    }
    [data-testid="stFileUploaderDropzone"]:hover {
        background-color: #eef1f7 !important;
        border-color: #6b7fa0 !important;
    }

    /* Uploaded file name row */
    [data-testid="stFileUploaderFile"] {
        border-color: #8a9bb5 !important;
        background-color: #eef1f7 !important;
    }

    /* Progress bar */
    [data-testid="stFileUploaderProgressBar"] > div {
        background-color: #5a8a5e !important;
    }

    /* Multiselect selected tag/chip */
    [data-testid="stMultiSelect"] span[data-baseweb="tag"] {
        background-color: #5a8a5e !important;
        border-color: #5a8a5e !important;
        color: #fff !important;
    }
    /* X (remove) icon inside the tag */
    [data-testid="stMultiSelect"] span[data-baseweb="tag"] span[role="presentation"] svg {
        fill: #fff !important;
    }
    /* Dropdown option highlight on hover */
    [data-testid="stMultiSelect"] li[aria-selected="true"],
    [data-testid="stMultiSelect"] li:hover {
        background-color: #dceedd !important;
        color: #2d4f30 !important;
    }

    /* Add-type toggle buttons */
    div.add-type-toggle .stButton > button {
        background-color: #fff !important;
        border: 1.5px solid #5a8a5e !important;
        color: #2d4f30 !important;
        font-weight: 500 !important;
        border-radius: 6px !important;
    }
    div.add-type-toggle .stButton > button:hover {
        background-color: #dceedd !important;
    }
    div.add-type-toggle .stButton > button[kind="primary"] {
        background-color: #5a8a5e !important;
        border-color: #5a8a5e !important;
        color: #fff !important;
    }

    /* "Add" identifier button — green, scoped to its wrapper div */
    div.add-btn .stButton > button {
        background-color: #5a8a5e !important;
        border-color: #5a8a5e !important;
        color: #fff !important;
    }
    div.add-btn .stButton > button:hover {
        background-color: #3d6b41 !important;
        border-color: #3d6b41 !important;
    }

    /* Any remaining accent/primary colour Streamlit injects via CSS vars */
    :root {
        --primary-color: #5a8a5e !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Initialize session state
if 'temp_dir' not in st.session_state:
    st.session_state.temp_dir = None
if 'output_path' not in st.session_state:
    st.session_state.output_path = None
if 'processing' not in st.session_state:
    st.session_state.processing = False
if 'extracted_cost_centers' not in st.session_state:
    st.session_state.extracted_cost_centers = []
if 'selected_cost_centers' not in st.session_state:
    st.session_state.selected_cost_centers = []
if 'template_type' not in st.session_state:
    st.session_state.template_type = 'opex'   # 'opex' or 'project'
# Always reload config from disk so stale session-state fields never persist
st.session_state.app_config = ConfigManager.load_config()


def render_preview_section(output_path: str):
    """Render Excel file preview section"""
    st.subheader("Preview Report")
    
    try:
        # Get sheet names
        sheet_names = ExcelPreviewHandler.get_sheet_names(output_path)
        
        if not sheet_names:
            st.warning("Could not load sheet names from output file")
            return
        
        # Sheet selector
        selected_sheet = st.selectbox(
            "Select Sheet to Preview",
            sheet_names,
            help="Choose which sheet to preview"
        )
        
        # Determine max rows (50 for first sheet, 10 for others)
        is_first_sheet = (selected_sheet == sheet_names[0])
        max_rows = 50 if is_first_sheet else 10
        
        # Get sheet info
        sheet_info = ExcelPreviewHandler.get_sheet_info(output_path, selected_sheet)
        
        # Display metadata
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Rows", f"{sheet_info['row_count']:,}")
        with col2:
            st.metric("Columns", sheet_info['column_count'])
        with col3:
            st.metric("Preview Rows", max_rows)
        
        # Load and display preview
        # For first sheet, use configured header row
        header_row = None
        if is_first_sheet:
            header_row = st.session_state.app_config.template.header_row
        
        with st.spinner(f"Loading preview of {selected_sheet}..."):
            df_preview = ExcelPreviewHandler.preview_sheet(
                output_path,
                selected_sheet,
                max_rows,
                header_row=header_row
            )
        
        if df_preview is not None:
            st.dataframe(
                df_preview,
                use_container_width=True,
                height=400
            )
            
            # Show preview info
            rows_shown = len(df_preview)
            total_rows = sheet_info['row_count']
            if rows_shown < total_rows:
                st.caption(
                    f"Showing first {rows_shown} rows of {total_rows:,} total rows"
                )
            else:
                st.caption(f"Showing all {rows_shown} rows")
        else:
            st.error("Could not load sheet preview")
            
    except Exception as e:
        st.error(f"Error displaying preview: {str(e)}")


def render_config_section():
    """Render configuration settings section"""
    
    config = st.session_state.app_config
    
    # Status indicator
    col_status, col_buttons = st.columns([3, 1])
    with col_status:
        if ConfigManager.is_using_custom_config():
            st.success("Using Custom Configuration")
        else:
            st.info("Using Default Configuration")
    
    with col_buttons:
        # Action buttons in a row
        btn_col1, btn_col2 = st.columns(2)
        
        with btn_col1:
            if st.button("Save Configuration", use_container_width=True, help="Save configuration to file", key="save_config_btn"):
                is_valid, errors = ConfigManager.validate_config(config)
                if is_valid:
                    if ConfigManager.save_config(config):
                        st.success("Config saved!")
                        st.rerun()
                    else:
                        st.error("Failed to save config")
                else:
                    st.error("Validation errors:")
                    for error in errors:
                        st.error(f"• {error}")
        
        with btn_col2:
            if st.button("Reset to Defaults", use_container_width=True, help="Reset to default configuration", key="reset_config_btn"):
                # Show confirmation in a separate area
                st.session_state.show_reset_confirm = True
            
            # Handle confirmation separately
            if st.session_state.get('show_reset_confirm', False):
                if st.checkbox("Confirm reset to defaults?", key="confirm_reset_main"):
                    # Delete custom config file first
                    ConfigManager.delete_custom_config()
                    # Then reset session state
                    st.session_state.app_config = ConfigManager.get_default_config()
                    st.session_state.show_reset_confirm = False
                    st.success("Reset to defaults!")
                    st.rerun()
    
    
    # Configuration sections in columns for better use of space
    col1, col2 = st.columns(2)
    
    with col1:
        # Template Settings
        with st.expander("Template Settings", expanded=False):
            config.template.header_row = st.number_input(
                "Header Row",
                value=config.template.header_row,
                min_value=1,
                max_value=1000,
                help="Row number where PO headers start in the template (1-based)"
            )
            
            po_col_input = st.text_input(
                "PO Column",
                value=config.template.po_col,
                help="Column letter containing PO numbers (e.g., 'B')"
            )
            config.template.po_col = po_col_input.upper() if po_col_input else config.template.po_col
            
            config.template.po_stop_marker = st.text_input(
                "PO Stop Marker",
                value=config.template.po_stop_marker,
                help="Text marker indicating end of PO section"
            )
            
            cc_col_input = st.text_input(
                "Cost Center Column",
                value=config.template.cost_center_col,
                help="Column letter containing cost center IDs (e.g., 'A')"
            )
            config.template.cost_center_col = cc_col_input.upper() if cc_col_input else config.template.cost_center_col
            
            config.template.cost_center_start_row = st.number_input(
                "Cost Center Start Row",
                value=config.template.cost_center_start_row,
                min_value=1,
                max_value=1000,
                help="Row number where cost centers start (1-based)"
            )
            
            use_cc_end_row = st.checkbox(
                "Set Cost Center End Row",
                value=config.template.cost_center_end_row is not None,
                help="Enable to stop reading cost centers at a specific row instead of the first blank cell"
            )
            if use_cc_end_row:
                config.template.cost_center_end_row = st.number_input(
                    "Cost Center End Row",
                    value=config.template.cost_center_end_row if config.template.cost_center_end_row is not None else config.template.cost_center_start_row,
                    min_value=1,
                    max_value=1000,
                    help="Row number where cost centers stop being read (1-based, inclusive)"
                )
            else:
                config.template.cost_center_end_row = None
        
        # Forecast Settings
        with st.expander("Forecast Settings", expanded=False):
            config.forecast_reader.po_col = st.text_input(
                "PO Column Name",
                value=config.forecast_reader.po_col,
                help="Column name in forecast files containing PO numbers",
                key="forecast_po_col"
            )
    
    with col2:
        # Transactional Settings
        with st.expander("Transactional Settings", expanded=False):
            config.transactional_detail_reader.required_cols = st.multiselect(
                "Required Columns",
                options=["PO Number", "Accounting Period", "GL Transaction Amount", "Type", "Cost Center*", "WBS Element"],
                default=config.transactional_detail_reader.required_cols,
                help="Columns that must exist in transactional file",
                key="trans_required_cols"
            )
            
            config.transactional_detail_reader.valid_types = st.multiselect(
                "Valid Transaction Types",
                options=["Actual", "Accrual", "Reversal", "Reclass", "ER", "Budget", "Forecast"],
                default=config.transactional_detail_reader.valid_types,
                help="Valid transaction types to process",
                key="trans_valid_types"
            )
            
            # Column Mappings (nested)
            with st.expander("Column Mappings", expanded=False):
                colmap = config.transactional_detail_reader.colmap
                
                colmap['po'] = st.text_input("PO Column", value=colmap['po'], help="Column name for PO numbers", key="trans_po_col")
                colmap['month'] = st.text_input("Period Column", value=colmap['month'], help="Column name for accounting period (YYYYMM format, e.g. 202601 = Jan)", key="trans_month_col")
                colmap['amount'] = st.text_input("Amount Column", value=colmap['amount'], help="Column name for transaction amount", key="trans_amount_col")
                colmap['classifier'] = st.text_input("Classifier Column", value=colmap['classifier'], help="Column name for transaction classifier", key="trans_classifier_col")
                colmap['cost_center'] = st.text_input("Cost Center Column", value=colmap['cost_center'], help="Column name for cost center", key="trans_cc_col")
                colmap['wbs'] = st.text_input("WBS Column", value=colmap['wbs'], help="Column name for WBS element", key="trans_wbs_col")
                colmap['type'] = st.text_input("Type Column", value=colmap['type'], help="Column name for transaction type", key="trans_type_col")
        
        # Writer Settings
        with st.expander("Writer Settings", expanded=False):
            config.template_writer.output_path = st.text_input(
                "Output Filename",
                value=config.template_writer.output_path,
                help="Name for generated output file (must end with .xlsx)",
                key="writer_output_path"
            )
            
            config.template_writer.overwrite = st.checkbox(
                "Overwrite Existing Files",
                value=config.template_writer.overwrite,
                help="Allow overwriting existing output files",
                key="writer_overwrite"
            )
            
            dec_col_input = st.text_input(
                "Dec Accrual Reversal Column",
                value=config.template_writer.dec_acc_reversal_col,
                help="Column letter for December accrual reversals (e.g., 'N')",
                key="writer_dec_col"
            )
            config.template_writer.dec_acc_reversal_col = dec_col_input.upper() if dec_col_input else config.template_writer.dec_acc_reversal_col
            
            # Source Columns (nested)
            with st.expander("Source Columns", expanded=False):
                st.write("**Forecast Source Columns:**")
                config.template_writer.forecast_source_cols = st.multiselect(
                    "forecast_cols_label",
                    options=[
                        "PO #",
                        "Jan 2026 - FTotal", "Feb 2026 - FTotal", "March 2026 - FTotal",
                        "April 2026 - FTotal", "May 2026 - FTotal", "June 2026 - FTotal",
                        "July 2026 - FTotal", "Aug 2026 - FTotal", "Sep 2026 - FTotal",
                        "Oct 2026 - FTotal", "Nov 2026 - FTotal", "Dec 2026 - FTotal"
                    ],
                    default=config.template_writer.forecast_source_cols,
                    label_visibility="collapsed",
                    help="Columns to copy from forecast files",
                    key="writer_forecast_cols"
                )
                
                st.write("**Transactional Source Columns:**")
                config.template_writer.transactional_source_cols = st.multiselect(
                    "trans_cols_label",
                    options=[
                        "PO Number", "Accounting Period", "AP Voucher Number",
                        "Vendor Name", "WBS Element", "GL Invoice Date",
                        "GL Posting Date", "GL Line Description", "Description",
                        "GL Transaction Amount", "GL BER Corp Amount",
                        "AP01", "AP02", "AP03", "Type"
                    ],
                    default=config.template_writer.transactional_source_cols,
                    label_visibility="collapsed",
                    help="Columns to copy from transactional file",
                    key="writer_trans_cols"
                )


# Header
st.markdown(
    """
    <h1 style='
        text-align: center;
        font-size: 1.4rem !important;
        font-weight: 700 !important;
        line-height: 1.2 !important;
        margin-bottom: 0.2rem;
        color: #1f2328;
    '>Financial Automation Report Generator</h1>
    """,
    unsafe_allow_html=True,
)

st.divider()

# ── Three-column uploader row ─────────────────────────────────────────────────
up_col1, up_col2, up_col3 = st.columns(3, gap="medium")

with up_col1:
    st.subheader("Template File")
    template_file = st.file_uploader(
        "template_label",
        type=['xlsx'],
        help="Upload the financial spreadsheet template (.xlsx format)",
        key="template_upload",
        label_visibility="collapsed"
    )

    # Detect template type and extract identifiers when template is uploaded
    if template_file is not None:
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                tmp_file.write(template_file.getvalue())
                tmp_path = tmp_file.name

            from src.utils import detect_template_type
            ttype = detect_template_type(tmp_path)
            st.session_state.template_type = ttype
            app_config = st.session_state.app_config

            if ttype == 'project':
                from src.project_template_reader import ProjectTemplateReader
                from src.utils import load_config as _load_proj_cfg
                _pcfg = _load_proj_cfg('configs/config_project.yaml')['template']
                temp_reader = ProjectTemplateReader(
                    file_path=tmp_path,
                    header_row=_pcfg['header_row'],
                    po_col=_pcfg['po_col'],
                    po_stop_marker=_pcfg.get('po_stop_marker', 'Previous Period Invoices'),
                    wbs_col=_pcfg.get('wbs_col', 'A'),
                    p3_id_col=_pcfg.get('p3_id_col', 'B'),
                    wbs_start_row=_pcfg.get('wbs_start_row', 2),
                )
                identifiers = list(temp_reader.p3_wbs_map.keys())
            else:
                from src.template_reader import TemplateReader
                temp_reader = TemplateReader(
                    file_path=tmp_path,
                    header_row=app_config.template.header_row,
                    po_col=app_config.template.po_col,
                    po_stop_marker=app_config.template.po_stop_marker,
                    cost_center_col=app_config.template.cost_center_col,
                    cost_center_start_row=app_config.template.cost_center_start_row,
                )
                identifiers = temp_reader.cost_centers

            prev = st.session_state.extracted_cost_centers
            if set(identifiers) != set(prev):
                st.session_state.extracted_cost_centers = identifiers
                st.session_state.selected_cost_centers = list(identifiers)

            os.unlink(tmp_path)

        except Exception as e:
            st.error(f"Error reading template: {str(e)}")

    if template_file is None:
        st.info("Template File Required")

with up_col2:
    st.subheader("Transactional Detail File")
    transactional_file = st.file_uploader(
        "transactional_label",
        type=['xlsx'],
        help="Upload the TIES transactional detail file (.xlsx format)",
        key="transactional_upload",
        label_visibility="collapsed"
    )
    if transactional_file:
        st.success("Transactional File Uploaded")
    else:
        st.info("Transactional File Required")

with up_col3:
    st.subheader("Forecast Files")
    forecast_files = st.file_uploader(
        "forecast_label",
        type=['xlsx'],
        accept_multiple_files=True,
        help="Upload one or more vendor forecast files (.xlsx format)",
        key="forecast_upload",
        label_visibility="collapsed"
    )
    if forecast_files:
        st.success(f"{len(forecast_files)} Forecast File(s) Uploaded")
    else:
        st.info("At Least One Forecast File Required")

# ── Template type banner (full width, below all three upload columns) ─────────
if template_file is not None and st.session_state.get('template_type'):
    is_proj = st.session_state.template_type == 'project'
    banner_color = "#b5651d" if is_proj else "#6b4fa0"
    banner_label = "📂 Project / P3 Template" if is_proj else "📋 OpEx Template"
    st.markdown(
        f'<div style="background:{banner_color};color:#fff;padding:10px 20px;'
        f'border-radius:6px;font-size:14px;font-weight:600;margin-top:4px;">'
        f'{banner_label}</div>',
        unsafe_allow_html=True,
    )

# ── Optional LE File ──────────────────────────────────────────────────────────
le_col, = st.columns([1])
with le_col:
    st.subheader("LE File (Optional)")
    st.caption("OpEx Only")
    le_file = st.file_uploader(
        "le_label",
        type=['xlsx'],
        help="Upload the ERP LE file (e.g. ERP LE3 2026 8-13-26.xlsx)",
        key="le_upload",
        label_visibility="collapsed"
    )
    if le_file:
        st.success(f"LE File Uploaded: {le_file.name}")
    else:
        st.info("No LE file uploaded — LE data will be skipped")

st.divider()

# ── Generate button ───────────────────────────────────────────────────────────
all_files_uploaded = all([template_file, forecast_files, transactional_file])

generate_button = st.button(
    "Generate Report",
    type="primary",
    disabled=not all_files_uploaded or st.session_state.processing,
    use_container_width=True
)

# Process pipeline when button is clicked
if generate_button:
    st.session_state.processing = True
    st.session_state.output_path = None
    
    # Create containers for progress and status
    progress_container = st.empty()
    status_container = st.empty()
    log_container = st.container()
    
    try:
        # Progress bar
        progress_bar = progress_container.progress(0)
        
        def update_progress(percentage):
            """Callback to update progress bar"""
            progress_bar.progress(percentage / 100.0)
        
        # Initialize logger
        logger = StreamlitLogger(status_container=status_container)
        
        # Save uploaded files
        logger.info("Saving uploaded files...")
        update_progress(5)
        temp_dir, file_paths = FileHandler.save_uploaded_files(
            template_file=template_file,
            forecast_files=forecast_files,
            transactional_file=transactional_file,
            le_file=le_file if le_file else None
        )
        st.session_state.temp_dir = temp_dir
        
        # Run pipeline with config from session state
        app_config = st.session_state.app_config
        config_dict = ConfigManager._config_to_dict(app_config)
        orchestrator = PipelineOrchestrator(config=config_dict, logger=logger, progress_callback=update_progress)
        output_path = orchestrator.run(file_paths, selected_cost_centers=st.session_state.selected_cost_centers)
        st.session_state.output_path = output_path
        
        # Clear progress bar after completion
        progress_container.empty()
        
        # Display execution summary
        execution_time = logger.get_execution_time()
        exception_summary = orchestrator.get_exception_summary()
        
        st.success(f"Report generated successfully in {execution_time:.2f} seconds!")
        
        # Show exception summary if available
        if exception_summary:
            with st.expander("Exception Summary", expanded=False):
                total_exceptions = exception_summary.get('total', 0)
                st.metric("Total Exceptions", total_exceptions)
                
                if total_exceptions > 0:
                    st.write("**Breakdown by Type:**")
                    counts = exception_summary.get('counts', {})
                    percentages = exception_summary.get('percentages', {})
                    for exc_type, count in counts.items():
                        percentage = percentages.get(exc_type, 0)
                        st.write(f"- {exc_type}: {count} ({percentage:.1f}%)")
        
        # Show detailed logs
        with st.expander("Detailed Logs", expanded=False):
            logs = logger.get_logs()
            for level, timestamp, message in logs:
                if level == 'ERROR':
                    st.error(f"[{timestamp}] {message}")
                elif level == 'WARNING':
                    st.warning(f"[{timestamp}] {message}")
                elif level == 'SUCCESS':
                    st.success(f"[{timestamp}] {message}")
                else:
                    st.info(f"[{timestamp}] {message}")
    
    except Exception as e:
        st.error(f"Error generating report: {str(e)}")
        st.exception(e)
    
    finally:
        st.session_state.processing = False

st.divider()

# Download Section
if st.session_state.output_path:
    st.header("Download Report")
    
    output_path = st.session_state.output_path
    output_filename = Path(output_path).name
    
    # Read file for download
    file_bytes = FileHandler.get_output_file_bytes(output_path)
    
    if file_bytes:
        st.download_button(
            label="Download Generated Report",
            data=file_bytes,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
        
        st.success(f"Report ready: {output_filename}")
        st.info("Click the button above to download your report.")
    else:
        st.error("Could not read output file for download.")
    
    st.divider()
    
    # Preview Section
    render_preview_section(output_path)

# Cleanup temp files on session end (best effort)
if st.session_state.temp_dir:
    # Note: Streamlit doesn't have a reliable session end callback
    # Temp files will be cleaned up by OS eventually
    pass

# Footer
st.divider()
st.markdown("""
<div style='text-align: center; color: #666; padding: 20px;'>
    <p>Financial Automation Report Generator v1.0</p>
</div>
""", unsafe_allow_html=True)

# Made with Bob
