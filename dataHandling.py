from RPA.Excel.Files import Files
from openpyxl import Workbook

def getLocationsFromExcel(excelFileName:str,default_locations):
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
    except FileNotFoundError:
        # Create a new workbook and select the active sheet
        workbook = Workbook()
        sheet = workbook.active
        
        # Set the value of cell A1 to "Locations"
        sheet["A1"] = "Locations"
        
        for i, location in enumerate(default_locations, start=2):
            sheet[f"A{i}"] = location
        
        workbook.save(excelFileName)

        locations = default_locations

    except Exception as e:
        raise ValueError(f"An error occurred while accessing '{excelFileName}': {e}")

    finally:
        try:
            excel.close_workbook()
        except Exception as e:
            print(f"Failed to close workbook: {e}")
    
    return locations;

def getRecipienEmails(excelFileName, administrators):
    """
    Fetches the list of recipient emails from an excel document
    """
    recipients = []
    excel = Files()

    try:
        excel.open_workbook(excelFileName)

        last_row = excel.find_empty_row()

        for row in range(2, last_row):
            email = excel.get_cell_value(row,"A")
            if __emptyRowSkipper(email):
                recipients.append(email)
    except FileNotFoundError:
        # If the file is missing, create a new workbook with a default structure
        workbook = Workbook()
        sheet = workbook.active

        # Set header values
        sheet["A1"] = "Email"
        sheet["B1"] = "Role"
        
        # Populate default user emails with "Administrator" role
        for i, email in enumerate(administrators, start=2):
            sheet[f"A{i}"] = email
            sheet[f"B{i}"] = "Administrator"
            recipients.append(email)  # Add to recipients list

        # Save the new workbook
        workbook.save(excelFileName)

    except Exception as e:
        raise ValueError(f"An error occurred while reading '{excelFileName}': {e}")

    finally:
        try:
            excel.close_workbook()
        except Exception as e:
            print(f"Failed to close workbook: {e}")

    return recipients

def __emptyRowSkipper(value):
    if value is None: return False;
    return True;

def sortDataByCity(sources):
    """
    Sorts data by their city, so that the averages can be calculated for each of the cities.
    """
    dataSortedByCity = {}

    for source in sources.values():
        for cityData in source:
            city_name = cityData["kaupunki"]
            
            # Check if the city already has an entry in the dictionary
            if city_name in dataSortedByCity:
                dataSortedByCity[city_name].append(cityData)
            else:
                dataSortedByCity[city_name] = [cityData]
    
    return dataSortedByCity