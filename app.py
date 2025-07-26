import streamlit as st
import pandas as pd
import tempfile
import os

st.set_page_config(page_title="Employee Setup Automation", layout="wide")
st.title("Employee Setup - SAP Primary Role Extractor")

uploaded_file = st.file_uploader("Upload the Excel file", type=["xls", "xlsx"])

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.xls'):
            raw0 = pd.read_excel(uploaded_file, header=None, dtype=str, engine='xlrd')
        else:
            raw0 = pd.read_excel(uploaded_file, header=None, dtype=str)
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        st.stop()

    # Strip leading/trailing whitespaces
    raw0 = raw0.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    # Identify the header row
    header_row = None
    for i, row in raw0.iterrows():
        if isinstance(row[0], str) and row[0].strip().lower() == 'employee name':
            header_row = i
            break

    if header_row is None:
        st.error("Header row with 'Employee Name' not found.")
        st.stop()

    headers = raw0.iloc[header_row].tolist()
    data = raw0.iloc[header_row + 1:].reset_index(drop=True)
    data.columns = headers

    # Clean the dataframe
    data = data[headers]
    data = data.dropna(how='all')

    # Filter rows where 'Task Description' is 'Emp Setup 08.1- SAP Primary Role'
    filtered_df = data[data['Task Description'] == 'Emp Setup 08.1- SAP Primary Role'].copy()

    if filtered_df.empty:
        st.warning("No matching rows with 'Emp Setup 08.1- SAP Primary Role' found.")
    else:
        # Select required columns if they exist
        columns_to_display = ['Employee Name', 'Person Number', 'Assignee Name',
                              'Task Description', 'Task Status', 'Start Date', 'Target End Date']

        available_columns = [col for col in columns_to_display if col in filtered_df.columns]
        output_df = filtered_df[available_columns]

        # Show the result
        st.subheader("Filtered Data")
        st.dataframe(output_df, use_container_width=True)

        # Save to Excel
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            with pd.ExcelWriter(tmp.name, engine='openpyxl') as writer:
                output_df.to_excel(writer, index=False, sheet_name='FilteredData')

                # Add autofilter
                worksheet = writer.sheets['FilteredData']
                worksheet.auto_filter.ref = worksheet.dimensions

            tmp_path = tmp.name

        with open(tmp_path, "rb") as f:
            st.download_button(
                label="Download Filtered Excel",
                data=f,
                file_name="filtered_employee_setup.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # Clean up temp file
        os.remove(tmp_path)
