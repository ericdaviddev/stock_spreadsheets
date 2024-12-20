import win32com.client as win32
import gc
from config import numeric_columns, percentage_columns

def run_macro_on_workbook(macro_workbook_path, target_workbook_path, macro_name, exclusion_file_path):
    # Open Excel application
    global target_workbook
    excel = win32.DispatchEx("Excel.Application")
    # excel.Visible = True  # Set to True for debugging

    try:
        # Open the workbook containing the macro
        macro_workbook = excel.Workbooks.Open(macro_workbook_path)

        # Open the target workbook (where the macro will operate)
        target_workbook = excel.Workbooks.Open(target_workbook_path)

        # Run the macro
        full_macro_name = f"'{macro_workbook.Name}'!{macro_name}"
        excel.Application.Run(full_macro_name, exclusion_file_path)

        # Save and close the target workbook
        target_workbook.Save()
        macro_workbook.Close()
        excel.Quit()
        print(f"Macros completed, and workbook saved: {target_workbook_path}")
    except Exception as e:
        print(f"An error occurred: {e}")
        target_workbook.Close(SaveChanges=False)  # Ensure the workbook is closed
    finally:
        print("Process complete")
        del target_workbook
        del excel
        gc.collect()
        # excel.Quit()




# Example usage
# run_macro_on_workbook(macro_workbook_path, target_workbook_path, macro_name)
