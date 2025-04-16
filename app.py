import streamlit as st
import pandas as pd
import webbrowser
from io import BytesIO
from datetime import datetime

st.set_page_config(layout="wide")
st.title("🧠 High Confidence Case Review")

st.markdown("---")

if "current_case_index" not in st.session_state:
    st.session_state.current_case_index = 0

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

df = None
all_sheets = {}

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    for sheet_name in xls.sheet_names:
        sheet_df = xls.parse(sheet_name)
        all_sheets[sheet_name] = sheet_df
        if df is None:
            df = sheet_df

    if df is not None:
        # Find the index of the first unreviewed case
        review_column = "review_(tp/fp)"
        if review_column in df.columns:
            unreviewed_cases = df[df[review_column].isna()]
            if not unreviewed_cases.empty:
                st.session_state.current_case_index = unreviewed_cases.index.min()
            else:
                st.session_state.current_case_index = len(df) - 1
        else:
            st.session_state.current_case_index = 0

        current_index = st.session_state.current_case_index

        reviewed_cases = df[df[review_column].notna()] if review_column in df.columns else pd.DataFrame()

        if len(reviewed_cases) == len(df):
            st.success("All cases have been completed.")

        case = df.loc[current_index]

        st.write(f"Progress: {len(reviewed_cases)}/{len(df)} cases completed")
        st.write(f"Case {current_index+1}/{len(df)}: {case.get('accession', '')}")

        col_open, col_text = st.columns([1, 2])
        with col_open:
            studio_url = case.get("studio_link", "")
            if studio_url:
                st.markdown(f"<a href='{studio_url}' target='_blank'><button>Open Studio Link</button></a>", unsafe_allow_html=True)

        with col_text:
            st.subheader("Login Info")
            st.write("**Username:** rpxuser")
            st.write("**Password:** PpD4u2RK")

        if len(reviewed_cases) < len(df):
            tp_fp = st.radio("True Positive / False Positive", ["TP", "FP"], key=f"tp-fp_{current_index}")
            second_opinion = st.checkbox("Request Second Opinion", key=f"second-opinion_{current_index}")
            request_report = st.radio("Request Report", ["No", "Yes"], key=f"request-report_{current_index}")
            location_type = st.text_area("Location/Type", key=f"location-type_{current_index}")
            comments = st.text_area("Comments (Optional)", key=f"comment_{current_index}")

            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                if st.button("Submit and Next"):
                    df.at[current_index, "review_(tp/fp)"] = str(tp_fp)
                    df.at[current_index, "2nd_opinion_(y/n)"] = str("Yes" if second_opinion else "No")
                    df.at[current_index, "request_report_(y/n)"] = str(request_report)
                    df.at[current_index, "location/type"] = str(location_type).strip()
                    df.at[current_index, "comments"] = str(comments).strip()

                    if current_index + 1 < len(df):
                        st.session_state.current_case_index += 1
                        st.rerun()
                    else:
                        st.success("All cases have been completed.")
                        st.rerun()

# Show download button if any file is uploaded
if uploaded_file and df is not None:
    st.subheader("Download Completed Workbook")

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet, data in all_sheets.items():
            data.to_excel(writer, sheet_name=sheet, index=False)
            ws = writer.sheets[sheet]
            for i, column in enumerate(data.columns, 1):
                ws.column_dimensions[chr(64 + i)].width = max(15, len(str(column)) + 2)

    output.seek(0)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    base_filename = uploaded_file.name.replace(".xlsx", "")
    download_filename = f"{base_filename}-updated-{timestamp}.xlsx"

    st.download_button(
        label="📥 Download Updated Excel",
        data=output,
        file_name=download_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.write("👈 **Click here to launch a case.**")
