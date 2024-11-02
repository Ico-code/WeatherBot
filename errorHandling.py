from openpyxl import load_workbook
from datetime import datetime

errorLogFilePath = "error_log.xlsx";

def logErrorToExcel(error_details={"errMsg":"api key is empty","errLVL":"medium", "errLocation":"tasks.py"}):
    """
    Write error message in error_log.xlsx
    Parameters:
        error_details should contain all the data that is needed for logging the info into excel.
        This would include the following:
            errMsg:
                message containing a description of the error
            errLVL:
                how critical is the error
            errLocation:
                where, or what file did the error occur in

    # Example usage
    error_info = {
        "Error Level": "Error",
        "Location": "data_processing.py, line 54",
        "Error Message": "File not found",
    }
    logErrorToExcel(error_info)
    """
    try:
        workbook = load_workbook(errorLogFilePath)
        sheet = workbook.active
    except FileNotFoundError:
        from openpyxl import Workbook
        workbook = Workbook()
        sheet = workbook.active
        # Write headers if file is new
        sheet.append(["Timestamp", "Error Level", "Location","Error Message"])
    # Append error details as a new row
    sheet.append([
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        error_details.get("errLVL", ""),
        error_details.get("errLocation",""),
        error_details.get("errMsg", ""),
    ])

    # Save the workbook
    workbook.save(errorLogFilePath)