import xlwings as xw

def copy_move_tab(sourcefile, sourcetab, targettab, targetfile=None, targetwb=None):
    """
    Copies a tab from a source Excel file to a target Excel file.

    Args:
    sourcefile (str): The path of the source Excel file.
    targetfile (str): The path of the target Excel file.
    sourcetab (str): The name of the tab to copy from the source Excel file.
    targettab (str): The name of the new tab in the target Excel file.

    This function opens both the source and target Excel files, copies the specified tab from the source
    to the target, renames the copied tab, and then saves and closes both files.
    """

    # Open the source and target Excel workbooks
    wb_source = xw.Book(sourcefile)
    if targetwb:
        wb_target = targetwb
    else:
        wb_target = xw.Book(targetfile)

    # Select the tab to copy from the source workbook
    sheet_to_copy = wb_source.sheets[sourcetab]

    # Copy the tab to the beginning of the target workbook and rename it
    sheet_to_copy.api.Copy(Before=wb_target.sheets[0].api)
    wb_target.sheets[0].name = targettab

    # Close the source workbook without saving changes
    wb_source.close()

    # Save and close the target workbook
    wb_target.save()
    # wb_target.close()

