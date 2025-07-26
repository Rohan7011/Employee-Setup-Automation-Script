# app.py

import pandas as pd
import os
from datetime import datetime
import numpy as np
import streamlit as st
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

st.title("Emp Setup Request Filter App")

uploaded_file = st.file_uploader("Upload your .xls or .xlsx file", type=["xls", "xlsx"])
if uploaded_file:
    file_extension = os.path.splitext(uploaded_file.name)[-1].lower()

    # Read uploaded file
    if file_extension == ".xls":
        raw0 = pd.read_excel(uploaded_file, header=None, dtype=str)
        output = BytesIO()
        raw0.to_excel(output, index=False, header=False, engine='openpyxl')
        output.seek(0)
        process_path = output
    else:
        raw0 = pd.read_excel(uploaded_file, header=None, dtype=str)
        process_path = uploaded_file

    raw = pd.read_excel(process_path, header=None, dtype=str)

    # Locate header rows
    row_main = row_sub = None
    for i in range(len(raw)):
        vals = raw.iloc[i].fillna('').astype(str).str.strip().tolist()
        if row_main is None and 'HD ID' in vals:
            row_main = i
        if row_sub is None and 'Task ID' in vals:
            row_sub = i
        if row_main is not None and row_sub is not None:
            break

    if row_main is None or row_sub is None:
        st.error("Could not find both 'HD ID' and 'Task ID' header rows.")
        st.stop()

    main_hdr = raw.iloc[row_main].fillna('').astype(str).str.strip().tolist()
    sub_hdr  = raw.iloc[row_sub].fillna('').astype(str).str.strip().tolist()
    flat_cols = [sub_hdr[c] if sub_hdr[c] else main_hdr[c] for c in range(raw.shape[1])]

    data = raw.iloc[row_sub+1 :].copy()
    data.columns = flat_cols

    if 'HD ID' in data.columns:
        data['HD ID'] = data['HD ID'].replace('HD ID', np.nan)
        data['HD ID'] = data['HD ID'].ffill()

    expected = ['HD ID', 'Task ID', 'Task Desc', 'Task Tech', 'Task Create', 'Task Status', 'Task Group']
    missing = [col for col in expected if col not in data.columns]
    if missing:
        st.error(f"Missing expected columns: {missing}")
        st.stop()

    clean = data.dropna(how='all')
    clean = clean[clean['Task Desc'] == "Emp Setup 08.1- SAP Primary Role"]
    clean = clean[expected]

    st.success("Filtered Data:")
    st.dataframe(clean)

    # Export to Excel with filters
    output_excel = BytesIO()
    clean.to_excel(output_excel, index=False)
    output_excel.seek(0)

    # Apply filter using openpyxl
    wb = load_workbook(output_excel)
    ws = wb.active
    last_col_letter = get_column_letter(ws.max_column)
    ws.auto_filter.ref = f"A1:{last_col_letter}1"
    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    st.download_button(
        label="Download Filtered Excel File",
        data=final_output,
        file_name=f"EmpSetup_Requests_{datetime.today():%Y-%m-%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
