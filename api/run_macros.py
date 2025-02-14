import gc

def run_macro_on_workbook(excel, macro_workbook, target_workbook, macro_name, exclusion_file_path):
    """

    :param excel: excel object
    :param macro_workbook: workbook containing macro
    :param target_workbook: workbook to run macro against
    :param macro_name:
    :param exclusion_file_path: workbook containing excluded symbols
    """
    try:
        # Run the macro
        full_macro_name = f"'{macro_workbook.Name}'!{macro_name}"
        excel.Application.Run(full_macro_name, exclusion_file_path)

        # Save and close the target workbook
        target_workbook.Save()
        macro_workbook.Close()
        print(f"Macros completed, and workbook saved")
    except Exception as e:
        print(f"An error occurred: {e}")
        target_workbook.Close(SaveChanges=False)  # Ensure the workbook is closed
    finally:
        print("Process complete")
        del target_workbook
        gc.collect()

