import gc

import pandas as pd
import os
import win32com.client as win32
from run_macros import run_macro_on_workbook
from config import numeric_columns, percentage_columns

startsWithColumns = ["The data and information", "Brokerage services are", "Date downloaded"]

def clean_and_combine_sheets(folder_path, output_file,  exclusion_file, macro_file, macro_name, data_types=None):
    """
    Combine and clean spreadsheets from a folder into a single main file.

    :param macro_name:
    :param macro_file:
    :param exclusion_file:
    :param folder_path: Path to the folder containing spreadsheets
    :param output_file: Path to save the combined main file
    :param data_types: Dictionary specifying data types for columns (optional)
    """
    # Close main_sheet if it's open
    # close_if_open(output_file)

    all_data = []

    for file_name in os.listdir(folder_path):
        # Skip unsupported file types
        if not (file_name.endswith('.xlsx') or file_name.endswith('.csv')):
            print(f"Skipping unsupported file: {file_name}")
            continue

        file_path = os.path.join(folder_path, file_name)

        try:
            # Load the file
            if file_name.endswith('.xlsx'):
                df = pd.read_excel(file_path)
            elif file_name.endswith('.csv'):
                df = pd.read_csv(file_path)
        except Exception as e:
            print(f"Error reading file {file_name}: {e}")
            continue

        # Clean the data
        df = clean_sheet(df)

        # Add to the combined data
        all_data.append(df)

    # Combine all data into one main DataFrame
    if all_data:
        main_df = pd.concat(all_data, ignore_index=True)

        remove_non_numeric_characters(main_df, numeric_columns)
        remove_non_numeric_characters(main_df, percentage_columns)

        # Save the main sheet
        main_df.to_excel(output_file, index=False)
        run_macro_on_workbook(macro_file, output_file, macro_name, exclusion_file)
        apply_formatting(output_file, numeric_columns, percentage_columns)
        print(f"Price sheet saved to {output_file}")
    else:
        print("No valid files were found to process.")


def remove_non_numeric_characters(main_df, columns):
    for col in columns:
        if col in main_df.columns:
            main_df[col] = main_df[col].replace(r'[^\d.-]', '', regex=True)  # Remove non-numeric characters
            main_df[col] = pd.to_numeric(main_df[col], errors='coerce')  # Convert to float


def close_if_open(file_path):
    excel = win32.Dispatch("Excel.Application")
    try:
        for workbook in excel.Workbooks:
            if workbook.FullName.lower() == file_path.lower():
                workbook.Close(SaveChanges=False)
                gc.collect()
                excel.Quit()
                print(f"Closed open workbook: {file_path}")
    except Exception as e:
        print(f"An error occurred during close_if_open: {e}")
    finally:
        del excel
        gc.collect()

def clean_sheet(df):
    """
    Clean the DataFrame by removing unwanted rows and standardizing columns.
    """
    # Remove rows where the 'Description' column starts with "REMOVE"
    if "Account Number" in df.columns:
        for item in startsWithColumns:
            df = df[~df["Account Number"].str.startswith(item, na=False)]

    # Additional cleaning logic (e.g., currency conversion, text trimming)
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].str.strip()  # Remove extra spaces

    return df

def apply_formatting(target_file, numeric_columns, percentage_columns):
    """
    Apply formatting to specific columns in an Excel file.
    :param target_file: Path to the Excel file.
    """
    excel = win32.Dispatch("Excel.Application")
    # excel.Visible = True  # Keep Excel visible for debugging

    try:
        # Open the workbook
        wb = excel.Workbooks.Open(target_file)
        ws = wb.Sheets(1)  # Adjust the sheet index or name as needed

        # Example 1: Format a column as currency
        for col in numeric_columns:
            # ws.Columns(col).NumberFormat = "$#,##0.00"  # Currency with 2 decimal places
            ws.Columns(get_column_index_by_heading(ws, col)).NumberFormat = "$#,##0.00;[Red]($#,##0.00)"  # Red font for negative numbers

        for col in percentage_columns:
            format_percentage_column(ws, col)
            # ws.Columns(get_column_index_by_heading(ws, col)).NumberFormat = "0.00%"  # Percentage with 2 decimal places

        freeze_panes(wb)
        autofit_columns_by_heading(wb.Sheets(1), numeric_columns + percentage_columns)
        # Save changes
        wb.Save()
        # wb.Close()
        print(f"Formatting applied and workbook saved: {target_file}")
        excel.Quit()
    except Exception as e:
        print(f"An error occurred while formatting: {e}")
    finally:
        del wb, excel
        gc.collect()

def format_percentage_column(ws, heading):
    """
    Format a column as a percentage, ensuring values are normalized.
    :param ws: The worksheet object.
    :param heading: The column heading to format as percentage.
    """
    col_index = get_column_index_by_heading(ws, heading)
    if col_index:
        # Normalize values by dividing by 100
        for row_index in range(2, ws.UsedRange.Rows.Count + 1):  # Start from row 2 to skip the header
            cell = ws.Cells(row_index, col_index)
            if cell.Value is not None:  # Ensure the cell is not empty
                cell.Value = cell.Value / 100

        # Apply percentage format
        ws.Columns(col_index).NumberFormat = "0.00%"
        print(f"Percentage format applied to column: {heading}")
    else:
        print(f"Column '{heading}' not found.")


def freeze_panes(target_workbook):
    target_workbook.Application.ActiveWindow.SplitRow = 1  # Freeze top row
    target_workbook.Application.ActiveWindow.SplitColumn = 1  # Freeze first column
    target_workbook.Application.ActiveWindow.FreezePanes = True


def autofit_columns_by_heading(ws, headings):
    """
    AutoFit specific columns based on their heading names.
    :param ws: The worksheet object.
    :param headings: List of column heading names to AutoFit.
    """
    try:
        # Iterate through the specified headings
        for heading in headings:
            # Find the column index for the heading
            for col_index in range(1, ws.UsedRange.Columns.Count + 1):
                cell_value = ws.Cells(1, col_index).Value  # Assuming headers are in the first row
                if cell_value and cell_value.strip() == heading:
                    ws.Columns(col_index).AutoFit()  # AutoFit the matched column
                    print(f"AutoFit applied to column: {heading}")
                    break
    except Exception as e:
        print(f"An error occurred: {e}")


def get_column_index_by_heading(ws, heading):
    """
    Get the column index of a specified column heading in an Excel sheet.
    :param ws: The worksheet object.
    :param heading: The column heading to search for.
    :return: The column index (1-based) or None if not found.
    """
    # Iterate through the first row to find the heading
    for col_index in range(1, ws.UsedRange.Columns.Count + 1):
        cell_value = ws.Cells(1, col_index).Value  # Assuming headers are in the first row
        if cell_value and cell_value.strip() == heading:
            return col_index
    return None

# Example usage
# folder_path = "C:/path/to/spreadsheets"
# output_file = "C:/path/to/main_spreadsheet.xlsx"
# clean_and_combine_sheets(folder_path, output_file, data_types)


