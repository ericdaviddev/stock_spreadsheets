"""Entry point for automating the stock spreadsheet workflow.

This script:
- Locates and combines brokerage position spreadsheets
- Applies exclusions and aggregation rules
- Triggers an Excel macro in a separate workbook to finish processing
"""

from __future__ import annotations

import argparse
from pathlib import Path

from combine_spreadsheets import combine_and_clean_sheets
from config import data_types
from run_macros import run_macro_on_workbook  # noqa: F401  # imported for side effects / clarity


def main(
    folder_path: str | Path,
    output_file: str | Path,
    exclusion_file: str | Path,
    macro_file: str | Path,
    macro_name: str,
) -> None:
    """Run the end-to-end combination and cleaning workflow.

    Args:
        folder_path: Directory containing the input brokerage CSV/Excel files.
        output_file: Path to the consolidated/processed Excel file.
        exclusion_file: Path to the workbook defining excluded symbols and columns to sum.
        macro_file: Path to the macro-enabled workbook (.xlsm) that finalizes processing.
        macro_name: Name of the macro to execute in the macro workbook.
    """
    folder_path = Path(folder_path)
    output_file = Path(output_file)
    exclusion_file = Path(exclusion_file)
    macro_file = Path(macro_file)

    combine_and_clean_sheets(
        folder_path=folder_path,
        output_file=output_file,
        exclusion_file=exclusion_file,
        macro_file=macro_file,
        macro_name=macro_name,
        data_types=data_types,
    )


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description=(
            "Combine brokerage position files, apply exclusions, and "
            "run the Excel macro to produce a cleaned summary workbook."
        )
    )
    parser.add_argument(
        "folder_path",
        help="Path to the folder containing downloaded brokerage position files (CSV/Excel).",
    )
    parser.add_argument(
        "output_file",
        help="Path to save the combined/processed Excel workbook.",
    )
    parser.add_argument(
        "exclusion_file",
        help="Path to the Excel file listing symbols to exclude and columns to aggregate.",
    )
    parser.add_argument(
        "macro_file",
        help="Path to the macro-enabled Excel workbook (.xlsm) to run.",
    )
    parser.add_argument(
        "macro_name",
        help="Name of the macro to execute in the macro workbook (e.g. 'ProcessExclusionsAndTotals').",
    )

    args = parser.parse_args()

    main(
        folder_path=args.folder_path,
        output_file=args.output_file,
        exclusion_file=args.exclusion_file,
        macro_file=args.macro_file,
        macro_name=args.macro_name,
    )
