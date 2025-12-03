import logging
import os
from contextlib import contextmanager
from pathlib import Path
from typing import List, Dict, Any, Optional

import pandas as pd
import win32com.client as win32

from config import startsWithColumns, numeric_columns, percentage_columns, account_columns
from run_macros import run_macro_on_workbook
from utils import ExcelFormatter, add_timestamp_to_filename, clean_dataframe

# Configure logging
logging.basicConfig(
    filename="error_log.txt",
    filemode="a",
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    level=logging.ERROR
)

@contextmanager
def excel_application(quit_on_exit: bool = False):
    """Context manager for Excel application to ensure proper cleanup.
    
    Args:
        quit_on_exit: If True, Excel will quit when context manager exits. 
                     If False, Excel will remain open.
    """
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True  # Make Excel window visible
    try:
        yield excel
    finally:
        if quit_on_exit:
            excel.Quit()
        # Removed gc.collect() as it may force Excel to close

def validate_inputs(folder_path: str, output_file_path: str, exclusion_file_path: str, macro_file_path: str) -> None:
    """
    Validate input paths before processing.
    
    Args:
        folder_path: Path to input folder
        output_file_path: Path to output file
        exclusion_file_path: Path to exclusion file
        macro_file_path: Path to macro file
    
    Raises:
        ValueError: If any path is invalid
    """
    if not os.path.exists(folder_path):
        raise ValueError(f"Input folder not found: {folder_path}")
    if not os.path.exists(exclusion_file_path):
        raise ValueError(f"Exclusion file not found: {exclusion_file_path}")
    if not os.path.exists(macro_file_path):
        raise ValueError(f"Macro file not found: {macro_file_path}")
    output_dir = os.path.dirname(output_file_path)
    if not os.path.exists(output_dir):
        raise ValueError(f"Output directory does not exist: {output_dir}")

def combine_and_clean_sheets(
    folder_path: str,
    output_file_path: str,
    exclusion_file_path: str,
    macro_file_path: str,
    macro_name: str,
    data_types: Optional[Dict[str, Any]] = None
) -> str:
    """
    Combine and clean spreadsheets from a folder into a single main file.

    Args:
        folder_path: Path to the folder containing spreadsheets
        output_file_path: Path to save the combined main file
        exclusion_file_path: Path to get excluded symbols and columns to sum
        macro_file_path: Path to get the macro workbook
        macro_name: Name of the macro
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
        
        with excel_application(quit_on_exit=False) as excel:
            output_file_path = process_data(
                all_dataframes,
                excel,
                exclusion_file_path,
                macro_file_path,
                macro_name,
                output_file_path
            )
            logging.info(f"Price sheet saved to {output_file_path}")
            return output_file_path
            
    except Exception as e:
        logging.error(f"Error in combine_and_clean_sheets: {e}", exc_info=True)
        raise

def process_files(folder_path: str, data_types: Optional[Dict[str, Any]] = None) -> List[pd.DataFrame]:
    """
    Process Excel and CSV files in the given folder.
    
    Args:
        folder_path: Path to folder containing spreadsheets
        data_types: Optional dictionary specifying data types for columns
    
    Returns:
        List of processed DataFrames
    
    Raises:
        ValueError: If no valid files are found
    """
    all_data = []
    valid_extensions = {'.xlsx', '.csv'}
    files = [f for f in os.listdir(folder_path) if Path(f).suffix.lower() in valid_extensions]
    
    if not files:
        raise ValueError(f"No valid files found in {folder_path}")
    
    for file_name in files:
        file_path = os.path.join(folder_path, file_name)
        try:
            # First read without dtype to avoid conversion errors
            df = pd.read_excel(file_path) if file_name.endswith('.xlsx') else pd.read_csv(file_path)
            
            # Clean the data first
            df = clean_dataframe(df, startsWithColumns)
            
            # Then apply data types after cleaning, with coerce for numeric columns
            # if data_types:
            #     for col, dtype in data_types.items():
            #         if col in df.columns:
            #             try:
            #                 if dtype in (float, 'float64'):
            #                     df[col] = pd.to_numeric(df[col], errors='coerce')
            #                 else:
            #                     df[col] = df[col].astype(dtype)
            #             except Exception as type_error:
            #                 logging.warning(f"Could not convert column {col} to {dtype}: {type_error}")
            #
            all_data.append(df)
            logging.info(f"Successfully processed {file_name}")
        except Exception as e:
            logging.error(f"Error processing {file_name}: {e}", exc_info=True)
            continue

    return all_data

def process_data(
    all_data: List[pd.DataFrame],
    excel: Any,
    exclusion_file: str,
    macro_file: str,
    macro_name: str,
    output_file: str
) -> str:
    """
    Process and format combined data.
    
    Args:
        all_data: List of DataFrames to process
        excel: Excel application object
        exclusion_file: Path to exclusion file
        macro_file: Path to macro file
        macro_name: Name of macro to run
        output_file: Output file path
    
    Returns:
        Path to processed output file
    
    Raises:
        ValueError: If data processing fails
    """
    macro_workbook = None
    target_workbook = None
    try:
        main_df = pd.concat(all_data, ignore_index=True)
        remove_non_numeric_characters(main_df, numeric_columns)
        remove_non_numeric_characters(main_df, percentage_columns)
        
        output_file = add_timestamp_to_filename(output_file)
        
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            main_df.to_excel(writer, index=False)
        
        # Open workbooks
        macro_workbook = excel.Workbooks.Open(macro_file)
        target_workbook = excel.Workbooks.Open(output_file)
        
        # Run macro
        run_macro_on_workbook(excel, macro_workbook, target_workbook, macro_name, exclusion_file)
        
        # Format workbook
        formatter = ExcelFormatter()
        formatter.format_numeric_columns(target_workbook.Sheets(1), numeric_columns)
        formatter.format_percentage_columns(target_workbook.Sheets(1), percentage_columns)
        formatter.freeze_panes(target_workbook)
        formatter.autofit_columns_by_heading(target_workbook.Sheets(1), numeric_columns + percentage_columns + account_columns)

        # Save
        target_workbook.Save()
        return output_file
        
    except Exception as e:
        logging.error(f"Error in process_data: {e}", exc_info=True)
        # Only close workbooks if there was an error
        try:
            if macro_workbook is not None:
                macro_workbook.Close(SaveChanges=False)
        except:
            pass
            
        try:
            if target_workbook is not None:
                target_workbook.Close(SaveChanges=True)
        except:
            pass
        raise ValueError(f"Failed to process data: {str(e)}")
    finally:
        # Only close the macro workbook in the finally block
        try:
            if macro_workbook is not None:
                macro_workbook.Close(SaveChanges=False)
        except:
            pass

def remove_non_numeric_characters(df: pd.DataFrame, columns: List[str]) -> None:
    """
    Remove non-numeric characters from specified columns.
    
    Args:
        df: DataFrame to process
        columns: List of column names to clean
    """
    try:
        df[columns] = df[columns].replace(r'[^\d.-]', '', regex=True).apply(pd.to_numeric, errors='coerce')
    except Exception as e:
        logging.error(f"Error cleaning numeric columns: {e}", exc_info=True)
        raise

def clean_dataframe(df: pd.DataFrame, startsWithColumns: List[str]) -> pd.DataFrame:
    """
    Clean the DataFrame by removing unwanted rows and standardizing columns.
    
    Args:
        df: DataFrame to clean
        startsWithColumns: List of column prefixes to exclude
    
    Returns:
        pd.DataFrame: Cleaned DataFrame
    """
    if "Account Number" in df.columns:
        for item in startsWithColumns:
            df = df[~df["Account Number"].str.startswith(item, na=False)]

    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].str.strip()

    return df
