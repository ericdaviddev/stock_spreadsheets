import gc

import pandas as pd
import os
import win32com.client as win32
from run_macros import run_macro_on_workbook
from config import startsWithColumns, numeric_columns, percentage_columns
from datetime import datetime

import logging

# Configure logging
logging.basicConfig(
    filename="error_log.txt",  # Log file name
    filemode="a",  # Append mode (use "w" for overwrite)
    format="%(asctime)s - %(levelname)s - %(message)s",  # Log format
    datefmt="%Y-%m-%d %H:%M:%S",  # Timestamp format
    level=logging.ERROR  # Log only errors and critical messages
)

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

    excel = win32.Dispatch("Excel.Application")
    all_data = process_files(folder_path)

    # Combine all data into one main DataFrame
    if all_data:
        try:
            output_file = process_data(all_data, excel, exclusion_file, macro_file, macro_name, output_file)
            print(f"Price sheet saved to {output_file}")
        except Exception as e:
            print(f"Error processing data: {e}")
            logging.error(f"Error in {__name__}: {e}", exc_info=True)
        finally:
            excel.Quit()
            del excel
            gc.collect()
    else:
        print("No valid files were found to process.")

def process_files(folder_path):
    """
    :param folder_path:
    :return: data list
    """
    all_data = []
    for file_name in os.listdir(folder_path):
        if not file_name.endswith(('.xlsx', '.csv')):  # Cleaner check
            print(f"Skipping unsupported file: {file_name}")
            continue

        file_path = os.path.join(folder_path, file_name)
        try:
            df = pd.read_excel(file_path) \
                if file_name.endswith('.xlsx') else pd.read_csv(file_path)
            df = clean_sheet(df)
            all_data.append(df)
        except Exception as e:
            logging.error(f"Error in {__name__}: {e}", exc_info=True)

    return all_data


def process_data(all_data, excel, exclusion_file, macro_file, macro_name, output_file):
    """
    :param all_data:
    :param excel:
    :param exclusion_file:
    :param macro_file:
    :param macro_name:
    :param output_file:
    :return:
    """
    main_df = pd.concat(all_data, ignore_index=True)
    remove_non_numeric_characters(main_df, numeric_columns)
    remove_non_numeric_characters(main_df, percentage_columns)
    # Save the main sheet
    output_file = add_time_stamp(output_file)
    # main_df.to_excel(output_file, index=False)

    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        main_df.to_excel(writer, index=False)

    # Open the workbook containing the macro
    macro_workbook = excel.Workbooks.Open(macro_file)
    # Open the target workbook (where the macro will operate)
    target_workbook = excel.Workbooks.Open(output_file)
    run_macro_on_workbook(excel, macro_workbook, target_workbook, macro_name, exclusion_file)
    apply_formatting(excel, output_file, numeric_columns, percentage_columns)
    return output_file


def add_time_stamp(filename):
    """
    :param filename:
    :return: filename with time stamp before extension
    """
    # Generate timestamp
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    # Insert timestamp before file extension
    name, ext = filename.rsplit(".", 1)  # Splitting filename and extension
    new_filename = f"{name}_{timestamp}.{ext}"
    return new_filename

def remove_non_numeric_characters(main_df, columns):
    main_df[columns] = main_df[columns].replace(r'[^\d.-]', '', regex=True).apply(pd.to_numeric, errors='coerce')

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

def apply_formatting(excel, target_file, numbers, percentages):
    """
    Apply formatting to specific columns in an Excel file.
    :param excel: 
    :param numbers: 
    :param percentages: 
    :param target_file: Path to the Excel file.
    """

    try:
        # Open the workbook
        wb = excel.Workbooks.Open(target_file)
        ws = wb.Sheets(1)  # Adjust the sheet index or name as needed

        # Example 1: Format a column as currency
        for col in numbers:
            ws.Columns(get_column_index_by_heading(ws, col)).NumberFormat = "$#,##0.00;[Red]($#,##0.00)"  # Red font for negative numbers

        for col in percentages:
            format_percentage_column(ws, col)

        freeze_panes(wb)
        autofit_columns_by_heading(wb.Sheets(1), numbers + percentages)
        wb.Save()
        print(f"Formatting applied and workbook saved: {target_file}")
    except Exception as e:
        print(f"An error occurred while formatting: {e}")
        logging.error(f"Error in {__name__}: {e}", exc_info=True)
    finally:
        del wb
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
    """
    :param target_workbook: freezes rows/columns
    """
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
        logging.error(f"Error in {__name__}: {e}", exc_info=True)


def get_column_index_by_heading(ws, heading):
    return next(
        (col_index for col_index in range(1, ws.UsedRange.Columns.Count + 1)
         if ws.Cells(1, col_index).Value and ws.Cells(1, col_index).Value.strip() == heading),
        None
    )


# Example usage
# folder_path = "C:/path/to/spreadsheets"
# output_file = "C:/path/to/main_spreadsheet.xlsx"
# clean_and_combine_sheets(folder_path, output_file, data_types)


