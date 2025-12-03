import logging
import os
from contextlib import contextmanager
from pathlib import Path
from typing import List, Dict, Any, Optional

import pandas as pd
import win32com.client as win32

from config import startsWithColumns, numeric_columns, percentage_columns, account_columns
from run_macros import run_macro_on_workbook
from utils import ExcelFormatter, add_timestamp_to_filename

# Configure logging
# Use INFO so info/warning logs are actually written to the file.
logging.basicConfig(
    filename="error_log.txt",
    filemode="a",
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    level=logging.INFO,
)


@contextmanager
def excel_application(quit_on_exit: bool = False):
    """
    Context manager for Excel application to ensure proper cleanup.

    Args:
        quit_on_exit: If True, Excel will quit when context manager exits.
                      If False, Excel will remain open.
    """
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True  # Keep Excel visible so user can see the workbook
    try:
        yield excel
    finally:
        if quit_on_exit:
            excel.Quit()
        # Intentionally do NOT force garbage collection here; it can interfere
        # with Excel's lifetime in some cases.


def validate_inputs(
    folder_path: str,
    output_file_path: str,
    exclusion_file_path: str,
    macro_file_path: str,
) -> None:
    """
    Validate input paths before processing.

    Args:
        folder_path: Path to input folder
        output_file_path: Path to output file
        exclusion_file_path: Path to exclusion file
        macro_file_path: Path to macro-enabled workbook

    Raises:
        ValueError: If any path is invalid
    """
    folder = Path(folder_path)
    output_file = Path(output_file_path)
    exclusion_file = Path(exclusion_file_path)
    macro_file = Path(macro_file_path)

    if not folder.exists():
        raise ValueError(f"Input folder not found: {folder}")
    if not exclusion_file.exists():
        raise ValueError(f"Exclusion file not found: {exclusion_file}")
    if not macro_file.exists():
        raise ValueError(f"Macro file not found: {macro_file}")
    if not output_file.parent.exists():
        raise ValueError(f"Output directory does not exist: {output_file.parent}")


def combine_and_clean_sheets(
    folder_path: str,
    output_file_path: str,
    exclusion_file_path: str,
    macro_file_path: str,
    macro_name: str,
    data_types: Optional[Dict[str, Any]] = None,
) -> str:
    """
    Combine and clean spreadsheets from a folder into a single main file, then
    run the configured Excel macro and format the result.

    Args:
        folder_path: Path to the folder containing spreadsheets
        output_file_path: Path to save the combined main file
        exclusion_file_path: Path to get excluded symbols and columns to sum
        macro_file_path: Path to the macro workbook
        macro_name: Name of the macro to run
        data_types: Optional dictionary specifying data types for columns

    Returns:
        str: Path to the processed output file

    Raises:
        ValueError: If input validation fails or no valid files found
    """
    try:
        validate_inputs(folder_path, output_file_path, exclusion_file_path, macro_file_path)

        all_dataframes = process_files(folder_path, data_types)
        if not all_dataframes:
            raise ValueError("No valid files were found to process")

        # Keep Excel open at the end so the user can see the result
        with excel_application(quit_on_exit=False) as excel:
            processed_output = process_data(
                all_dataframes,
                excel,
                exclusion_file_path,
                macro_file_path,
                macro_name,
                output_file_path,
            )
            logging.info("Price sheet saved to %s", processed_output)
            return processed_output

    except Exception as e:
        logging.error("Error in combine_and_clean_sheets: %s", e, exc_info=True)
        raise


def process_files(
    folder_path: str,
    data_types: Optional[Dict[str, Any]] = None,
) -> List[pd.DataFrame]:
    """
    Process Excel and CSV files in the given folder.

    Args:
        folder_path: Path to folder containing spreadsheets
        data_types: Optional dictionary specifying data types for columns

    Returns:
        List[pd.DataFrame]: List of processed DataFrames

    Raises:
        ValueError: If no valid files are found
    """
    folder = Path(folder_path)
    all_data: List[pd.DataFrame] = []
    valid_extensions = {".xlsx", ".csv"}

    files = [f for f in folder.iterdir() if f.suffix.lower() in valid_extensions]

    if not files:
        raise ValueError(f"No valid files found in {folder}")

    for file_path in files:
        try:
            if file_path.suffix.lower() == ".xlsx":
                df = pd.read_excel(file_path)
            else:
                df = pd.read_csv(file_path)

            # Clean the data first
            df = clean_dataframe(df, startsWithColumns)

            # If you later want data_types back, this is where to apply them.
            # Example (kept commented for now):
            #
            # if data_types:
            #     for col, dtype in data_types.items():
            #         if col in df.columns:
            #             try:
            #                 if dtype in (float, "float64"):
            #                     df[col] = pd.to_numeric(df[col], errors="coerce")
            #                 else:
            #                     df[col] = df[col].astype(dtype)
            #             except Exception as type_error:
            #                 logging.warning(
            #                     "Could not convert column %s to %s: %s",
            #                     col,
            #                     dtype,
            #                     type_error,
            #                 )

            all_data.append(df)
            logging.info("Successfully processed %s", file_path.name)
        except Exception as e:
            logging.error("Error processing %s: %s", file_path, e, exc_info=True)
            continue

    return all_data


def process_data(
    all_data: List[pd.DataFrame],
    excel: Any,
    exclusion_file: str,
    macro_file: str,
    macro_name: str,
    output_file: str,
) -> str:
    """
    Process and format combined data.

    Args:
        all_data: List of DataFrames to process
        excel: Excel application object
        exclusion_file: Path to exclusion file
        macro_file: Path to macro file
        macro_name: Name of macro to run
        output_file: Output file path (base name before timestamp)

    Returns:
        str: Path to processed output file

    Raises:
        ValueError: If data processing fails
    """
    macro_workbook = None
    target_workbook = None

    try:
        # Combine data
        main_df = pd.concat(all_data, ignore_index=True)

        # Clean numeric and percentage columns
        remove_non_numeric_characters(main_df, numeric_columns)
        remove_non_numeric_characters(main_df, percentage_columns)

        # Add timestamp to output filename
        output_file = add_timestamp_to_filename(output_file)

        # Write initial data frame to Excel
        with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
            main_df.to_excel(writer, index=False)

        # Open workbooks
        macro_workbook = excel.Workbooks.Open(macro_file)
        target_workbook = excel.Workbooks.Open(output_file)

        # Run macro
        run_macro_on_workbook(excel, macro_workbook, target_workbook, macro_name, exclusion_file)

        # Format workbook
        formatter = ExcelFormatter()
        sheet = target_workbook.Sheets(1)

        formatter.format_numeric_columns(sheet, numeric_columns)
        formatter.format_percentage_columns(sheet, percentage_columns)
        formatter.freeze_panes(target_workbook)
        formatter.autofit_columns_by_heading(
            sheet,
            numeric_columns + percentage_columns + account_columns,
        )

        # Save the target workbook but do NOT close it so the user can see it in Excel
        target_workbook.Save()

        return output_file

    except Exception as e:
        logging.error("Error in process_data: %s", e, exc_info=True)
        # On error, attempt to close workbooks carefully
        try:
            if macro_workbook is not None:
                macro_workbook.Close(SaveChanges=False)
        except Exception:
            pass

        try:
            if target_workbook is not None:
                # If we hit an error after writing, prefer not to save partial changes
                target_workbook.Close(SaveChanges=False)
        except Exception:
            pass

        raise ValueError(f"Failed to process data: {e}") from e

    finally:
        # Always close the macro workbook if it was opened
        try:
            if macro_workbook is not None:
                macro_workbook.Close(SaveChanges=False)
        except Exception:
            pass


def remove_non_numeric_characters(df: pd.DataFrame, columns: List[str]) -> None:
    """
    Remove non-numeric characters from specified columns in-place.

    Args:
        df: DataFrame to process
        columns: List of column names to clean
    """
    # Only operate on columns that actually exist
    present_columns = [c for c in columns if c in df.columns]
    if not present_columns:
        return

    try:
        df[present_columns] = (
            df[present_columns]
            .replace(r"[^\d.-]", "", regex=True)
            .apply(pd.to_numeric, errors="coerce")
        )
    except Exception as e:
        logging.error("Error cleaning numeric columns: %s", e, exc_info=True)
        raise


def clean_dataframe(df: pd.DataFrame, startsWithColumns: List[str]) -> pd.DataFrame:
    """
    Clean the DataFrame by removing unwanted rows and standardizing columns.

    Logic:
    - Remove rows where "Account Number" starts with any prefix in `startsWithColumns`.
    - Strip whitespace from all object (string) columns.

    Args:
        df: DataFrame to clean
        startsWithColumns: List of column prefixes to exclude

    Returns:
        pd.DataFrame: Cleaned DataFrame
    """
    if "Account Number" in df.columns:
        for prefix in startsWithColumns:
            df = df[~df["Account Number"].str.startswith(prefix, na=False)]

    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].str.strip()

    return df
