"""Utility functions for Excel operations and data processing."""
from typing import List, Dict, Any, Optional
import pandas as pd
from datetime import datetime
import logging
from pathlib import Path


class ExcelFormatter:
    """Class to handle Excel formatting operations."""
    
    @staticmethod
    def format_numeric_columns(ws: Any, columns: List[str]) -> None:
        """Format numeric columns with currency format.
        
        Args:
            ws: Excel worksheet object
            columns: List of column names to format
        """
        try:
            for col in columns:
                col_index = get_column_index_by_heading(ws, col)
                if col_index:
                    ws.Columns(col_index).NumberFormat = "$#,##0.00;[Red]($#,##0.00)"
        except Exception as e:
            logging.error(f"Error formatting numeric columns: {e}")
            raise

    @staticmethod
    def format_percentage_columns(ws: Any, columns: List[str]) -> None:
        """Format percentage columns and normalize values.
        
        Args:
            ws: Excel worksheet object
            columns: List of column names to format
        """
        try:
            for col in columns:
                col_index = get_column_index_by_heading(ws, col)
                if col_index:
                    for row_index in range(2, ws.UsedRange.Rows.Count + 1):
                        cell = ws.Cells(row_index, col_index)
                        if cell.Value is not None:
                            cell.Value = cell.Value / 100
                    ws.Columns(col_index).NumberFormat = "0.00%"
        except Exception as e:
            logging.error(f"Error formatting percentage columns: {e}")
            raise

    @staticmethod
    def freeze_panes(target_workbook: Any) -> None:
        """Freeze the top row and first column of the worksheet.
        
        Args:
            target_workbook: Excel workbook object
        """
        target_workbook.Application.ActiveWindow.SplitRow = 1
        target_workbook.Application.ActiveWindow.SplitColumn = 1
        target_workbook.Application.ActiveWindow.FreezePanes = True

    @staticmethod
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

def get_column_index_by_heading(ws: Any, heading: str) -> Optional[int]:
    """Get column index by heading name.
    
    Args:
        ws: Excel worksheet object
        heading: Column heading to find
    
    Returns:
        Column index if found, None otherwise
    """
    try:
        for i in range(1, ws.UsedRange.Columns.Count + 1):
            if ws.Cells(1, i).Value == heading:
                return i
        return None
    except Exception as e:
        logging.error(f"Error finding column index for {heading}: {e}")
        return None


def add_timestamp_to_filename(filename: str) -> str:
    """Add timestamp to filename before extension.
    
    Args:
        filename: Original filename
    
    Returns:
        Modified filename with timestamp
    """
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    path = Path(filename)
    return str(path.parent / f"{path.stem}_{timestamp}{path.suffix}")


def clean_dataframe(df: pd.DataFrame, exclude_patterns: List[str]) -> pd.DataFrame:
    """Clean DataFrame by removing unwanted rows and standardizing columns.
    
    Args:
        df: Input DataFrame
        exclude_patterns: List of patterns to exclude in Account Number
    
    Returns:
        Cleaned DataFrame
    """
    try:
        if "Account Number" in df.columns:
            for pattern in exclude_patterns:
                df = df[~df["Account Number"].str.startswith(pattern, na=False)]
        
        # Clean string columns
        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].str.strip()
        
        return df
    except Exception as e:
        logging.error(f"Error cleaning DataFrame: {e}")
        raise