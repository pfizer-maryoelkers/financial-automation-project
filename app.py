"""
Streamlit UI for Financial Automation Report Generator.
Phase 1: Basic UI with file upload and report generation.
"""

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
if 'app_config' not in st.session_state:
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
                options=["PO Number", "Month", "GL Transaction Amount", "Type", "Cost Center*", "WBS Element"],
                default=config.transactional_detail_reader.required_cols,
                help="Columns that must exist in transactional file",
                key="trans_required_cols"
            )
            
            config.transactional_detail_reader.valid_types = st.multiselect(
                "Valid Transaction Types",
                options=["Actual", "Accrual", "Reversal", "Budget", "Forecast"],
                default=config.transactional_detail_reader.valid_types,
                help="Valid transaction types to process",
                key="trans_valid_types"
            )
            
            # Column Mappings (nested)
            with st.expander("Column Mappings", expanded=False):
                colmap = config.transactional_detail_reader.colmap
                
                colmap['po'] = st.text_input("PO Column", value=colmap['po'], help="Column name for PO numbers", key="trans_po_col")
                colmap['month'] = st.text_input("Month Column", value=colmap['month'], help="Column name for month/period", key="trans_month_col")
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
                        "GL Transaction Amount", "GL BER Corp Amount", "Month",
                        "AP01", "AP02", "AP03", "Type"
                    ],
                    default=config.template_writer.transactional_source_cols,
                    label_visibility="collapsed",
                    help="Columns to copy from transactional file",
                    key="writer_trans_cols"
                )


# Header
st.title("Financial Automation Report Generator")
st.markdown("""
Upload your files and generate a comprehensive financial report with forecast data,
transactional details, and exception tracking.
""")

st.divider()

# File Upload Section
st.header("Upload Files")

# First row: Template File and Generate By Cost Center (TBD)
row1_col1, row1_col2 = st.columns(2)

with row1_col1:
    st.subheader("Template File")
    template_file = st.file_uploader(
        "template_label",
        type=['xlsx'],
        help="Upload the financial spreadsheet template (.xlsx format)",
        key="template_upload",
        label_visibility="collapsed"
    )
    
    # Extract cost centers when template is uploaded
    if template_file is not None:
        try:
            # Save template temporarily to extract cost centers
            import tempfile
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                tmp_file.write(template_file.getvalue())
                tmp_path = tmp_file.name
            
            # Extract cost centers using TemplateReader
            from src.template_reader import TemplateReader
            app_config = st.session_state.app_config
            temp_reader = TemplateReader(
                file_path=tmp_path,
                header_row=app_config.template.header_row,
                po_col=app_config.template.po_col,
                po_stop_marker=app_config.template.po_stop_marker,
                cost_center_col=app_config.template.cost_center_col,
                cost_center_start_row=app_config.template.cost_center_start_row
            )
            
            # Update session state with extracted cost centers
            st.session_state.extracted_cost_centers = temp_reader.cost_centers
            
            # Initialize selected cost centers if not already set
            if not st.session_state.selected_cost_centers:
                st.session_state.selected_cost_centers = temp_reader.cost_centers.copy()
            
            # Clean up temp file
            import os
            os.unlink(tmp_path)
            
            st.success(f"Extracted {len(temp_reader.cost_centers)} cost centers from template")
            
        except Exception as e:
            st.error(f"Error extracting cost centers: {str(e)}")

with row1_col2:
    st.subheader("Select Cost Centers")
    
    if st.session_state.extracted_cost_centers:
        # Display extracted cost centers
        st.write(f"**Found {len(st.session_state.extracted_cost_centers)} cost centers:**")
        
        # Multiselect for cost centers
        st.session_state.selected_cost_centers = st.multiselect(
            "Select cost centers to process",
            options=st.session_state.extracted_cost_centers,
            default=st.session_state.selected_cost_centers,
            help="Select which cost centers to include in the report generation"
        )
        
        # Option to add custom cost centers
        st.write("**Add Cost Center:**")
        col_input, col_button = st.columns([3, 1])
        with col_input:
            new_cc = st.text_input("Enter cost center ID", key="new_cost_center", label_visibility="collapsed", placeholder="Enter cost center ID")
        with col_button:
            if st.button("Add", use_container_width=True):
                if new_cc and new_cc not in st.session_state.extracted_cost_centers:
                    st.session_state.extracted_cost_centers.append(new_cc)
                    st.session_state.selected_cost_centers.append(new_cc)
                    st.success(f"Added: {new_cc}")
                    st.rerun()
                elif new_cc in st.session_state.extracted_cost_centers:
                    st.warning("Already exists")
                else:
                    st.warning("Enter a cost center ID")
        
        # Show selection summary
        if st.session_state.selected_cost_centers:
            st.info(f"Selected: {len(st.session_state.selected_cost_centers)} of {len(st.session_state.extracted_cost_centers)} cost centers")
        else:
            st.warning("No cost centers selected - report will include all cost centers")
    else:
        st.info("Upload a template file to extract cost centers")

# Second row: Transactional File and Forecast Files
row2_col1, row2_col2 = st.columns(2)

with row2_col1:
    st.subheader("Transactional Detail File")
    transactional_file = st.file_uploader(
        "transactional_label",
        type=['xlsx'],
        help="Upload the C-TIES transactional detail file (.xlsx format)",
        key="transactional_upload",
        label_visibility="collapsed"
    )

with row2_col2:
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
        st.info(f"{len(forecast_files)} forecast file(s) uploaded")

st.divider()

# File status indicators
st.subheader("Upload Status")
status_col1, status_col2, status_col3 = st.columns(3)

with status_col1:
    if template_file:
        st.success("Template file uploaded")
    else:
        st.warning("Template file required")

with status_col2:
    if transactional_file:
        st.success("Transactional file uploaded")
    else:
        st.warning("Transactional file required")

with status_col3:
    if forecast_files:
        st.success(f"{len(forecast_files)} forecast file(s) uploaded")
    else:
        st.warning("At least one forecast file required")

st.divider()

# Configuration Settings
st.subheader("Configuration Settings")
render_config_section()

st.divider()

# Generate Report Section
st.header("Generate Report")

# Check if all required files are uploaded
all_files_uploaded = all([template_file, forecast_files, transactional_file])

if not all_files_uploaded:
    st.info("Please upload all required files to enable report generation.")

# Generate button
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
            transactional_file=transactional_file
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
