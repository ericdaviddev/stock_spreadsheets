import gc
import logging

def run_macro_on_workbook(excel, macro_workbook, target_workbook, macro_name, exclusion_file_path):
    """
    Run a macro from the macro workbook on the target workbook.

    Args:
        excel: Excel application object
        macro_workbook: Workbook containing macro
        target_workbook: Workbook to run macro against
        macro_name: Name of the macro to run
        exclusion_file_path: Path to workbook containing excluded symbols
    """
    try:
        # Run the macro
        full_macro_name = f"'{macro_workbook.Name}'!{macro_name}"
        excel.Application.Run(full_macro_name, exclusion_file_path)
        
        # Save changes made by macro
        target_workbook.Save()
        logging.info(f"Macro {macro_name} completed successfully")
        
    except Exception as e:
        logging.error(f"Error running macro {macro_name}: {e}", exc_info=True)
        raise  # Re-raise the exception to be handled by the caller
