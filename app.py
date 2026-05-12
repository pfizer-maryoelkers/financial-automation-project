"""
Streamlit UI for Financial Automation Report Generator.
Phase 1: Basic UI with file upload and report generation.
"""

import streamlit as st
from pathlib import Path
from streamlit_backend import FileHandler, PipelineOrchestrator, StreamlitLogger
from src.utils import load_config

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

# Load default configuration
@st.cache_resource
def get_default_config():
    """Load default configuration from config_base.yaml"""
    return load_config('configs/config_base.yaml')

config = get_default_config()

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
            temp_reader = TemplateReader(
                file_path=tmp_path,
                header_row=config['template']['header_row'],
                po_col=config['template']['po_col'],
                po_stop_marker=config['template']['po_stop_marker'],
                cost_center_col=config['template']['cost_center_col'],
                cost_center_start_row=config['template']['cost_center_start_row']
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
        
        # Run pipeline
        orchestrator = PipelineOrchestrator(config=config, logger=logger, progress_callback=update_progress)
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
    <p>Phase 1: Basic UI with File Upload and Report Generation</p>
</div>
""", unsafe_allow_html=True)

# Made with Bob
