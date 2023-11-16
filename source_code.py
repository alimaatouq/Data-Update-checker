import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill
import os

# Function to compare two Excel files
@st.cache(allow_output_mutation=True)
def load_excel(file):
    return openpyxl.load_workbook(file)

def compare_excel_files(original_data, edited_data):
    fill_style = PatternFill(start_color="FDD835", end_color="FDD835", fill_type="solid")

    # Get a list of sheet names from the original workbook
    sheet_names = original_data.sheetnames

    # Create a new workbook to store the compared data
    compared_data = openpyxl.Workbook()

    for sheet_name in sheet_names:
        original_sheet = original_data[sheet_name]
        edited_sheet = edited_data[sheet_name]
        compared_sheet = compared_data.create_sheet(sheet_name)

        for row_original, row_edited, row_compared in zip(original_sheet.iter_rows(), edited_sheet.iter_rows(), compared_sheet.iter_rows()):
            for cell_original, cell_edited, cell_compared in zip(row_original, row_edited, row_compared):
                original_value = cell_original.value
                edited_value = cell_edited.value

                if original_value != edited_value:
                    cell_compared.value = edited_value
                    cell_compared.fill = fill_style
                else:
                    cell_compared.value = original_value

    return compared_data

# Streamlit app
st.title("Excel File Comparison App")

with st.expander("Upload Files"):
    original_file = st.file_uploader("Upload the Original Excel File", type=["xlsx"])
    edited_file = st.file_uploader("Upload the Edited Excel File", type=["xlsx"])

if original_file and edited_file:
    original_data = load_excel(original_file)
    edited_data = load_excel(edited_file)

    compared_data = compare_excel_files(original_data, edited_data)

    # Save the compared workbook and return the filename
    compared_filename = "compared_file.xlsx"
    compared_data.save(compared_filename)

    # Provide a download link for the compared file
    st.success(f"Comparison complete. You can download the compared file from the link below:")
    with open(compared_filename, "rb") as f:
        st.download_button("Download Compared File", f.read(), key="download_button", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Cleanup the file after downloading
os.remove(compared_filename)

st.write("Note: This app assumes that the sheet names are the same in both files for comparison.")
