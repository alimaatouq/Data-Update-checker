import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill
import io

# Function to compare two Excel files
def compare_excel_files(original_file, edited_file):
    fill_style = PatternFill(start_color="FDD835", end_color="FDD835", fill_type="solid")

    # Load the original and edited workbooks
    original_data = openpyxl.load_workbook(original_file)
    edited_data = openpyxl.load_workbook(edited_file)

    # Get a list of sheet names from the original workbook
    sheet_names = original_data.sheetnames


    for sheet_name in sheet_names:
        original_sheet = original_data[sheet_name]
        edited_sheet = edited_data[sheet_name]


        for row_original, row_edited in zip(original_sheet.iter_rows(), edited_sheet.iter_rows()):
            for cell_original, cell_edited in zip(row_original, row_edited):
                original_value = cell_original.value
                edited_value = cell_edited.value
                
                if original_value != edited_value:
                    cell_edited.fill = fill_style

    # Save the compared workbook and return the file bytes
    
    comapred_file = edited_file.save("compared_file.xlsx")
    return comapred_file

# Streamlit app
st.title("Excel File Comparison App")

original_file = st.file_uploader("Upload the Original Excel File", type=["xlsx"])
edited_file = st.file_uploader("Upload the Edited Excel File", type=["xlsx"])

if original_file and edited_file:

    st.success("Comparison complete. You can download the compared file below:")
    st.download_button("Download Compared File", data=compared_bytes, key="compared_file.xlsx")

