from combine_spreadsheets import clean_and_combine_sheets
from run_macros import run_macro_on_workbook
import argparse
from config import data_types
import win32com.client as win32


def main(folder_path, output_file, exclusion_file, macro_file, macro_name):
    # Define folder path and output file
    # folder_path = "C:/path/to/spreadsheets"
    # output_file = "C:/path/to/output/main_spreadsheet.xlsx"
    # exclusion_file = "C:/path/to/output/exclusions.xlsx"
    # macro_file = "C:/path/to/output/macro.xlsm"
    # macro_name = "ProcessExclusionsAndTotals"

    clean_and_combine_sheets(folder_path, output_file, exclusion_file, macro_file, macro_name, data_types)

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("folder_path", help="Path to the folder containing spreadsheets")
    parser.add_argument("output_file", help="Path to save the combined spreadsheet")
    parser.add_argument("exclusion_file", help="Path to get excluded symbols and columns to sum")
    parser.add_argument("macro_file", help="Path to get the macro workbook")
    parser.add_argument("macro_name", help="Name of the macro")

    args = parser.parse_args()
    main(args.folder_path, args.output_file, args.exclusion_file, args.macro_file, args.macro_name)