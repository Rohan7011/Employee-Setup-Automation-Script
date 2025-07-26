# 🧾 SAP Primary Role Task Extractor

This Streamlit web app allows you to upload Excel files containing onboarding task data and automatically extract rows related to the **"Emp Setup 08.1 - SAP Primary Role"** task. The app cleans, processes, and outputs a formatted Excel report with filters applied.

---

## ✅ Features

- Accepts `.xls` and `.xlsx` Excel files
- Automatically converts `.xls` to `.xlsx` internally
- Detects multi-level headers and builds a flat header structure
- Extracts only the required task rows based on task description
- Saves a clean Excel file with only relevant columns
- Applies auto-filter to the final Excel sheet

---

## 📁 Expected Input Format

The uploaded Excel file should contain task tracking data, typically with two header rows:
- **Main Header Row** (e.g., with "HD ID")
- **Sub Header Row** (e.g., with "Task ID", "Task Desc", etc.)

---

## 🖥 How to Use

1. Clone or download this repository.
2. Install requirements using:
   ```bash
   pip install -r requirements.txt


## ✅ Run the App using

streamlit run app.py
