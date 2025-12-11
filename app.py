import streamlit as st

st.title("Transactional Detail File Automation Tool")

st.write("Upload the two Excel files below to generate financial template.")

st.write('----------------')

# PO number
po_number = st.text_input("PO Number", placeholder="Enter the PO number...")
if po_number:
    st.write(f"PO Number: {po_number}")

# File upload section
forecast_file = st.file_uploader("Upload Vendor Forecast File (.xlsx)", type=["xlsx"])
actual_file = st.file_uploader("Upload C-Ties File (.xlsx)", type=["xlsx"])


# Placeholder section once files are uploaded
if forecast_file and actual_file:
    st.success("Files uploaded successfully!")
    # TODO: Add preview of files here

    if st.button("Generate Forecast Template"):
        st.info("Processing... demo only")
else:
    st.warning("Please upload both files to continue.")
