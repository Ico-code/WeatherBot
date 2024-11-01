from RPA.Excel.Files import Files

def getLocationsFromExcel(excelFileName:str):
    """
    Gets location data from Locations.xlsx
    Locations.xlsx is structured so that the first line is for headers and the rest are made up of location names
    """
    locations = []
    excel = Files()

    try:
        excel.open_workbook(excelFileName)        
        locationsEnd = excel.find_empty_row();
        for row in range(2, locationsEnd):
            excelRow = excel.get_cell_value(row,"A")
            if __emptyRowSkipper(excelRow):
                locations.append(excelRow)
    finally:
        excel.close_workbook()
    
    return locations;

def getRecipienEmails(excelFileName):
    """
    Fetches the list of recipient emails from an excel document
    """
    recipients = []
    excel = Files()

    try:
        excel.open_workbook(excelFileName)        
        locationsEnd = excel.find_empty_row();
        for row in range(2, locationsEnd):
            excelRow = excel.get_cell_value(row,"A")
            if __emptyRowSkipper(excelRow):
                recipients.append(excelRow)
    finally:
        excel.close_workbook()

    return recipients

def __emptyRowSkipper(value):
    if value is None: return False;
    return True;