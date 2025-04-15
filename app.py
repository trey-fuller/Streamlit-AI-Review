from pathlib import Path
import numpy as np
import pandas as pd
import streamlit as st
import webbrowser
import logging
import os
import openpyxl
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles.alignment import Alignment

# Configure logging to save log messages to a file named "adjudicator.log"
logging.basicConfig(
    filename="adjudicator.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# Get the script's parent directory (two levels up from current file)
BASE_DIR = Path(__file__).parent.parent  # Points to "High Confidence Review" folder

# Set the title and display instructions for users
st.title("AI Review Workflow")
st.write("Please select an Excel file from the workflow directory to proceed.")

# Find all Excel files in the parent directory matching the pattern
excel_files = sorted(list(BASE_DIR.glob("High_Confidence_Review_ICH_C_Spine_*.xlsx")))

# Ensure there are Excel files available
if not excel_files:
    st.error(f"No Excel files found in {BASE_DIR}. Please add files to that folder.")
    st.stop()

# User selects a file from the dropdown list of available Excel files
if 'selected_file' not in st.session_state:
    st.session_state['selected_file'] = excel_files[0].name if excel_files else None

selected_file = st.selectbox(
    "Select an existing Excel file:",
    [file.name for file in excel_files],
    index=[file.name for file in excel_files].index(st.session_state['selected_file']) if st.session_state['selected_file'] else 0
)

# Store the selected file in the session state
st.session_state['selected_file'] = selected_file

# Get the full path of the selected file
file_path = BASE_DIR / selected_file
st.write(f"📂 Selected File: {file_path}")


# Function to load data from an Excel file
def load_data(sheet_name, file_path):

    try:
        # Open the Excel file using pandas' ExcelFile class with openpyxl engine
        with pd.ExcelFile(file_path, engine="openpyxl") as xls:
            # Read all the sheets in the Excel file and store them in a dictionary
            all_sheets = {sheet: pd.read_excel(xls, sheet) for sheet in xls.sheet_names}
    except Exception as e:
        # Log any errors that occur while reading the Excel file
        logging.error(f"Failed to read Excel file: {e}")
        # Return an empty DataFrame and an empty dictionary if an error occurs
        return pd.DataFrame(), {}

    # Extract the specified sheet from the dictionary of all sheets
    df = all_sheets.get(sheet_name, pd.DataFrame())
 
    # Clean the column names of the dataframe by stripping leading/trailing whitespace,
    # replacing multiple whitespaces with a single underscore, and converting to lowercase
    df.columns = df.columns.map(str).str.strip().str.replace(r'\s+', '_', regex=True).str.lower()

    # Ensure that the required "completed" column exists in the dataframe
    # If the column doesn't exist, create it and set all values to "no"
    # If the column exists, fill any missing values with "no", convert to string, strip whitespace, and convert to lowercase
    if 'completed' not in df.columns:
        df['completed'] = "no"
    else:
        df['completed'] = df['completed'].fillna("no").astype(str).str.strip().str.lower()

    # Retrieve the last reviewed index from the "index" sheet in the Excel file
    # If the sheet exists and contains "Sheet" and "Last_Index" columns, get the last reviewed index for the current sheet
    # If the sheet doesn't exist or doesn't contain the required columns, set the last reviewed index to 0
    index_sheet = all_sheets.get("index", pd.DataFrame())
    if not index_sheet.empty and 'Sheet' in index_sheet.columns and 'Last_Index' in index_sheet.columns:
        last_reviewed_index = index_sheet.set_index('Sheet').to_dict().get('Last_Index', {}).get(sheet_name, 0)
    else:
        last_reviewed_index = 0

    # Return the dataframe, all sheets in the Excel file, and the last reviewed index
    return df, all_sheets, last_reviewed_index

    # Set initial state for editable fields in the form
default_values = {
        "tp-fp": "TP",
        "second-opinion": False,
        "request-report": "No",
        "location-type": "",
        "comment": ""
    }

for k, v in default_values.items():
    if k not in st.session_state:
        st.session_state[k] = v

# Define a function to reset the editable fields in the form for a given case index
def reset_form(case_index):
    # Get the current case data from the dataframe
    current_case = df.loc[case_index]
    # If the current case has not been completed, set the editable fields to their default values
    if current_case['completed'] == "no":
        for k, v in default_values.items():
            st.session_state[f"{k}_{case_index}"] = v
    # If the current case has been completed, set the editable fields to the values from the dataframe
    else:
        st.session_state[f"tp-fp_{case_index}"] = current_case["review_(tp/fp)"]
        st.session_state[f"second-opinion_{case_index}"] = current_case["2nd_opinion_(y/n)"] == "Yes"
        st.session_state[f"request-report_{case_index}"] = current_case["request_report_(y/n)"]
        st.session_state[f"location-type_{case_index}"] = current_case["location/type"] if isinstance(current_case["location/type"], str) else ""
        st.session_state[f"comment_{case_index}"] = current_case["comments"] if isinstance(current_case["comments"], str) else ""

# Define a function to load the previously submitted case entry for the previous case
def fill_previous_case():
    case_index = st.session_state["current_case_index"] - 1
    if case_index in case_indices:
        reset_form(case_index)

# Define a function to load the previously submitted case entry for the next case
def fill_next_case():
    case_index = st.session_state["current_case_index"] + 1
    if case_index in case_indices:
        reset_form(case_index)

# Set the title of the Streamlit browser window
st.title("AI Review Workflow")

# Stop script execution when browser window is closed
if st.session_state.get("_is_running", True) is False:
    st.stop()

# Set the initial value of "_is_running" to True
st.session_state["_is_running"] = True

# Display the filename of the Excel file being edited
st.write(f"Editing File: {os.path.basename(file_path)}")

# Set the name of the sheet containing the case data
sheet_name = "Case Data"

# Load the data from the Excel file and retrieve the last reviewed index
df, all_sheets, last_reviewed_index = load_data(sheet_name, file_path)

# Check if the dataframe was loaded correctly
if df.empty:
    st.error("Failed to load data.")
    st.stop()

# Filter the dataframe to separate reviewed and unreviewed cases
# reviewed_cases contains all rows where the 'completed' column is 'yes'
reviewed_cases = df[df['completed'] == "yes"]
# unreviewed_cases contains all rows where the 'completed' column is 'no'
unreviewed_cases = df[df['completed'] == "no"]

# Get a list of all case indices
# case_indices is a list of all indices in the dataframe
case_indices = df.index.tolist()

# Determine the current case index
# If the current case index is not in the session state or is not a valid case index,
# set it to the index of the first unreviewed case (if any)
# This ensures that the user always starts with an unreviewed case
if "current_case_index" not in st.session_state or st.session_state["current_case_index"] not in case_indices:
    st.session_state["current_case_index"] = unreviewed_cases.index.min() if not unreviewed_cases.empty else None

# If there are no more unreviewed cases, display a success message and stop the script
# This means that the user has reviewed all cases
if st.session_state["current_case_index"] is None:
    st.success("You have completed all available cases! No more cases to review.")
    logging.info("All cases reviewed for user.")
    st.stop()
# If this is the first case, mark that it has started but don't auto-launch
# This ensures that the user explicitly launches the first case
elif "case_auto_launched" not in st.session_state:
    st.session_state["case_auto_launched"] = True

# Get the current case index and retrieve the corresponding case data
current_index = st.session_state["current_case_index"]
case = df.loc[current_index]

# Display the current progress and case information
# total_cases is the total number of cases in the dataframe
total_cases = len(df)
# completed_cases is the number of reviewed cases
completed_cases = len(reviewed_cases)
# Display the progress and the current case information
st.write(f"Progress: {completed_cases}/{total_cases} cases completed")
st.write(f"Case {current_index+1}/{total_cases}: {case['accession']}")

# Create two columns for the studio link button and text
col_open, col_text = st.columns([1, 2])

# Studio link button
with col_open:
    # Display a button that opens the studio link for the current case
    if st.button("Open Studio Link"):
        webbrowser.open(case['studio_link'])

# Text
with col_text:
    # Display a message indicating that the first case will auto-launch
    st.write("👈 **Click here to launch the first case. The rest will auto-launch.**")

# Capture review details
# Display a radio button for the user to select True Positive (TP) or False Positive (FP) for the current case
tp_fp = st.radio("True Positive / False Positive", ["TP", "FP"], key=f"tp-fp_{current_index}")
# Display a checkbox for the user to request a second opinion for the current case
second_opinion = st.checkbox("Request Second Opinion", key=f"second-opinion_{current_index}")
# Display a radio button for the user to select whether they want to request a report for the current case
request_report = st.radio("Request Report", ["No", "Yes"], key=f"request-report_{current_index}")
# Display a text area for the user to enter the location/type of the current case
location_type = st.text_area("Location/Type", key=f"location-type_{current_index}")
# Display a text area for the user to enter any comments they have for the current case
comments = st.text_area("Comments (Optional)", key=f"comment_{current_index}")

# Create three columns for the buttons
col1, col2, col3 = st.columns([1, 2, 1])

# Column for the "Previous Case" button
with col1:
    # Display the "Previous Case" button and call the fill_previous_case function when clicked
    if st.button("Previous Case", on_click=fill_previous_case):
        # If the current case index is greater than 0, update the current case index and rerun the script
        if current_index > 0:
            st.session_state["current_case_index"] = case_indices[case_indices.index(current_index) - 1]
            st.rerun()
        # If the current case index is 0, display an information message
        else:
            st.info("You have reached the first case!")

# Column for the "Next Case" button
with col2:
    # Display the "Next Case" button and call the fill_next_case function when clicked
    if st.button("Next Case", on_click=fill_next_case):
        # If the current case index is less than the total number of cases minus 1, update the current case index and rerun the script
        if current_index < total_cases - 1:
            current_index = case_indices[case_indices.index(current_index) + 1]
            st.session_state["current_case_index"] = current_index
            st.rerun()
        # If the current case index is equal to the total number of cases minus 1, display an information message
        else:
            st.info("You have reached the last case!")

# Column for the "Submit & Next" button
with col3:
    # Display the "Submit & Next" button
    if st.button("Submit & Next"):
        # Update the dataframe with the user's input
        df.at[current_index, 'review_(tp/fp)'] = str(tp_fp).strip()
        df.at[current_index, '2nd_opinion_(y/n)'] = "Yes" if second_opinion else "No"
        df.at[current_index, 'request_report_(y/n)'] = str(request_report).strip()
        df.at[current_index, 'location/type'] = str(location_type).strip()  # Moved here
        df.at[current_index, 'comments'] = str(comments).strip() if comments.strip() else ""
        df.at[current_index, 'completed'] = "yes"

        # Update the index sheet with the current case index
        all_sheets[sheet_name] = df
        index_sheet = all_sheets.get('index', pd.DataFrame(columns=['Sheet', 'Last_Index']))
        if 'Sheet' not in index_sheet.columns or 'Last_Index' not in index_sheet.columns:
            index_sheet = pd.DataFrame(columns=['Sheet', 'Last_Index'])
        if not index_sheet.empty:
            index_sheet = index_sheet.set_index('Sheet')
        index_sheet.loc[sheet_name, 'Last_Index'] = current_index
        index_sheet = index_sheet.reset_index()
        all_sheets['index'] = index_sheet

        # Function to set column widths and alignment for a worksheet
        def set_column_widths_and_alignment(worksheet, widths):
            # Iterate over the column widths and indices
            for i, width in enumerate(widths, 1):
                # Set the width of the current column
                worksheet.column_dimensions[openpyxl.utils.get_column_letter(i)].width = width
                # Iterate over the cells in the current column
                for cell in worksheet[i]:
                   # Set the horizontal alignment of the current cell to left
                    cell.alignment = Alignment(horizontal='left')

        # Write to the the Excel file
        with pd.ExcelWriter(file_path, mode="w", engine="openpyxl") as writer:
            for sheet, data in all_sheets.items():
                data.to_excel(writer, sheet_name=sheet, index=False)
                
                # Get the worksheet object
                ws = writer.sheets[sheet]
                
                # Define column widths
                if sheet == "Case Data":
                    column_widths = [23, 11, 12, 14, 9, 14, 18, 20, 20, 20, 10, 13]   
                    set_column_widths_and_alignment(ws, column_widths)
                elif sheet == "index":
                    column_widths = [15, 15]  
                    set_column_widths_and_alignment(ws, column_widths)         

        # Log the updated case index
        logging.info(f"Updated case {current_index} as completed.")
        # Find the index of the next unreviewed case, if it exists
        next_unreviewed = unreviewed_cases.index[unreviewed_cases.index > current_index].min() if not unreviewed_cases.empty else None
        # Update the current case index in the session state
        st.session_state["current_case_index"] = next_unreviewed
        # If the next unreviewed case is the last case, display a success message and stop the script
        if next_unreviewed is None:
            st.success("You have completed all available cases! No more cases to review.")
            logging.info("All cases reviewed for user.")
            st.stop()
        elif not np.isnan(next_unreviewed):
            next_case = df.loc[next_unreviewed]
            reset_form(next_unreviewed)
            # Open the studio link for the next case in the user's web browser
            webbrowser.open(next_case['studio_link'])
            st.rerun()

# Create tabs for case review and login info
tab1, tab2 = st.tabs(["Case Review", "Login Info"])
# Display login info in the login info tab
with tab2:
    st.subheader("Login Info")
    st.write("**Username:** rpxuser")
    st.write("**Password:** PpD4u2RK")
