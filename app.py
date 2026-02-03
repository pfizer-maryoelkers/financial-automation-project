import streamlit as st
import tempfile
import pandas as pd
import time

from main import (
    ForecastReader,
    InvoiceActualReader,
    AccrualReader,
    combine_data,
    FinancialTemplateV2Writer,
)

# ===============================================
# Session State
# ===============================================
if "final_output_path" not in st.session_state:
    st.session_state.final_output_path = None

def remove_po(po):
    st.session_state.po_list.remove(po)


# ===============================================
# UI
# ===============================================
st.title("Financial Forecast Report Generator")

# -------------------------------
# Upload Inputs
# -------------------------------
st.subheader("1. Upload Input Files")

forecast_file = st.file_uploader("Forecast File (.xlsx)", type=["xlsx"])
tdf_file = st.file_uploader("Transactional Detail File (.xlsx)", type=["xlsx"])

# -------------------------------
# Uploading Template
# -------------------------------
st.subheader("2. Upload Financial Template")
st.text("Upload template file with POs filled out in column B")

template_file = st.file_uploader("Template File (.xlsx)", type=["xlsx"])

# -------------------------------
# Generate Report
# -------------------------------
st.subheader("3. Generate Report")

if st.button("Generate Report"):
    if not forecast_file or not tdf_file:
        st.error("Please upload both required files.")
        st.stop()

    if not template_file:
        st.error("Please upload financial template.")
        st.stop()

    # Area where progress messages appear
    progress_area = st.empty()
    log = []

    def update_progress(msg):
        log.append(msg)
        formatted = "<br>".join(log)
        progress_area.markdown(formatted, unsafe_allow_html=True)
        time.sleep(0.1)

    with st.spinner("Processing..."):

        try:
            # ----------------------------
            # Temp file writes
            # ----------------------------
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_f:
                tmp_f.write(forecast_file.getvalue())
                forecast_path = tmp_f.name

            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_tdf:
                tmp_tdf.write(tdf_file.getvalue())
                tdf_path = tmp_tdf.name

            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_temp:
                tmp_temp.write(template_file.getvalue())
                template_path = tmp_temp.name

            #template_path = "data/Financial Spreadsheet Template v2.xlsx"

            # ----------------------------
            # ETL Steps (with updates)
            # ----------------------------
            update_progress("✓ Loading forecast data...")
            forecast_reader = ForecastReader(forecast_path)
            forecast_reader.load_forecast()
            forecast_data = forecast_reader.get_forecast_data()
            forecast_reader_data = forecast_reader.data

            update_progress("✓ Loading actuals...")
            actual_reader = InvoiceActualReader(tdf_path)
            actual_reader.load_transactional_detail()
            actual_data = actual_reader.get_transactional_data()
            actual_reader_data = actual_reader.data

            update_progress("✓ Loading accruals...")
            accrual_reader = AccrualReader(tdf_path)
            accrual_reader.load_transactional_detail()
            accrual_data = accrual_reader.get_transactional_data()
            accrual_reader_data = accrual_reader.data

            update_progress("✓ Combining data...")
            combined = combine_data(forecast_data, actual_data, accrual_data)

            update_progress("✓ Filtering POs...")
            writer = FinancialTemplateV2Writer(template_path)
            writer.parse_template()
            pos = writer.pos
            filtered = {po: combined[po] for po in pos if po in combined}

            update_progress("✓ Writing template...")
            writer.write_data(filtered)
            writer.write_forecast_audit(forecast_reader_data, pos)
            writer.write_transactions_audit(actual_reader_data, pos)

            update_progress("✓ Saving output file...")
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_out:
                writer.save(tmp_out.name)
                st.session_state.final_output_path = tmp_out.name

            update_progress("Report generation complete!")

        except Exception as e:
            update_progress(f"❌ Pipeline failed: {e}")
            st.error("Pipeline failed — see details above.")
            st.stop()


# -------------------------------
# Download Section
# -------------------------------
if st.session_state.final_output_path:
    with open(st.session_state.final_output_path, "rb") as f:
        st.download_button(
            label="Download report.xlsx",
            data=f.read(),
            file_name="report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
