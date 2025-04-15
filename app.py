import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from io import BytesIO
from openpyxl.styles.alignment import Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="AI Review Workflow", layout="wide")
st.title("AI Review Workflow")

# Step 1: Upload Excel File
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
    sheet_names = xls.sheet_names

    if "Case Data" not in sheet_names:
        st.error("Sheet 'Case Data' not found in the uploaded file.")
        st.stop()

    df = pd.read_excel(xls, sheet_name="Case Data")
    df.columns = df.columns.map(str).str.strip().str.replace(r'\s+', '_', regex=True).str.lower()

    if 'completed' not in df.columns:
        df['completed'] = "no"
    else:
        df['completed'] = df['completed'].fillna("no").astype(str).str.strip().str.lower()

    case_indices = df.index.tolist()
    reviewed_cases = df[df['completed'] == "yes"]
    unreviewed_cases = df[df['completed'] == "no"]

    if "current_case_index" not in st.session_state:
        st.session_state["current_case_index"] = unreviewed_cases.index.min() if not unreviewed_cases.empty else None

    if st.session_state["current_case_index"] is None:
        st.success("You have completed all available cases!")
        st.stop()

    current_index = st.session_state["current_case_index"]
    case = df.loc[current_index]

    st.write(f"Progress: {len(reviewed_cases)}/{len(df)} cases completed")
    st.write(f"Case {current_index+1}/{len(df)}: {case.get('accession', '')}")

    tp_fp = st.radio("True Positive / False Positive", ["TP", "FP"], key=f"tp-fp_{current_index}")
    second_opinion = st.checkbox("Request Second Opinion", key=f"second-opinion_{current_index}")
    request_report = st.radio("Request Report", ["No", "Yes"], key=f"request-report_{current_index}")
    location_type = st.text_area("Location/Type", key=f"location-type_{current_index}")
    comments = st.text_area("Comments (Optional)", key=f"comment_{current_index}")

    col1, col2 = st.columns([1, 1])

    with col1:
        if st.button("Previous Case"):
            prev_index = case_indices.index(current_index) - 1
            if prev_index >= 0:
                st.session_state["current_case_index"] = case_indices[prev_index]
                st.rerun()

    with col2:
        if st.button("Submit & Next"):
            df.at[current_index, 'review_(tp/fp)'] = tp_fp
            df.at[current_index, '2nd_opinion_(y/n)'] = "Yes" if second_opinion else "No"
            df.at[current_index, 'request_report_(y/n)'] = request_report
            df.at[current_index, 'location/type'] = location_type.strip()
            df.at[current_index, 'comments'] = comments.strip()
            df.at[current_index, 'completed'] = "yes"

            next_index = unreviewed_cases.index[unreviewed_cases.index > current_index].min()
            st.session_state["current_case_index"] = next_index if not pd.isna(next_index) else None

            st.rerun()

    # Step 3: Download Button for Final Workbook
    st.markdown("---")
    st.subheader("Download Completed Workbook")

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Case Data")
        writer.book.create_sheet("index")
        index_sheet = writer.sheets["index"]
        index_sheet.append(["Sheet", "Last_Index"])
        index_sheet.append(["Case Data", st.session_state.get("current_case_index", len(df))])

        # Format Case Data columns
        ws = writer.sheets["Case Data"]
        for i, column in enumerate(df.columns, 1):
            ws.column_dimensions[get_column_letter(i)].width = 20
            for cell in ws[get_column_letter(i)]:
                cell.alignment = Alignment(horizontal='left')

    st.download_button(
        label="📥 Download Updated Excel",
        data=output.getvalue(),
        file_name="updated_review.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
