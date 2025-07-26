import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Employee Setup Filter", layout="wide")

st.title("Employee Setup Automation")
st.markdown("Upload the Excel file containing employee setup tasks to extract only the relevant ones.")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xls", "xlsx"])

if uploaded_file is not None:
    try:
        # Determine engine based on file type
        if uploaded_file.name.endswith('.xls'):
            raw0 = pd.read_excel(uploaded_file, header=None, dtype=str, engine='xlrd')
        else:
            raw0 = pd.read_excel(uploaded_file, header=None, dtype=str)
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
    else:
        # Find the row containing 'Employee Name' to locate header
        header_row_idx = None
        for i, row in raw0.iterrows():
            if row.str.contains('Employee Name', na=False).any():
                header_row_idx = i
                break

        if header_row_idx is None:
            st.error("Header row with 'Employee Name' not found.")
        else:
            df = pd.read_excel(uploaded_file, header=header_row_idx, dtype=str)
            df.columns = df.columns.str.strip()

            if 'Task Name' not in df.columns:
                st.error("'Task Name' column not found in the file.")
            else:
                filtered_df = df[df['Task Name'].str.contains("Emp Setup 08.1- SAP Primary Role", na=False)]

                if filtered_df.empty:
                    st.warning("No matching rows found for 'Emp Setup 08.1- SAP Primary Role'.")
                else:
                    st.success(f"Found {len(filtered_df)} matching rows.")
                    st.dataframe(filtered_df, use_container_width=True)

                    # Convert to Excel in-memory
                    buffer = BytesIO()
                    filtered_df.to_excel(buffer, index=False, engine='openpyxl')
                    buffer.seek(0)

                    st.download_button(
                        label="📥 Download Filtered Excel",
                        data=buffer,
                        file_name="filtered_employee_setup.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
