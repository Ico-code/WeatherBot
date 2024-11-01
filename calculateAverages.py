from openpyxl import Workbook, load_workbook
import zipfile

def combine_weather_data(data1, data2):
    """
    Merge data from two sources, averaging numerical values where possible and 
    keeping unique values otherwise.
    """
    combined_data = {}

        # Extract and convert temperatures, handling any units (like "Celsius")
    def parse_temperature(temp_str):
        # Remove any non-numeric parts (e.g., " Celsius") and convert to float
        return float(temp_str.split()[0])

    # Use the new helper function to parse temperatures
    temp1 = parse_temperature(data1.get("lämpötila", "0"))
    temp2 = parse_temperature(data2.get("lämpötila", "0"))

    # Average temperatures if both are available, otherwise use whichever exists
    if temp1 and temp2:
        combined_data["Keskimääräinen lämpötila"] = f"{(temp1 + temp2) / 2:.1f} Celsius"
    else:
        combined_data["lämpötila"] = data1.get("lämpötila") or data2.get("lämpötila")

    # Format wind speed to include m/s
    wind_speed = data1.get("tuulen_nopeus") or data2.get("tuulen_nopeus")
    if wind_speed:
        combined_data["Tuulen nopeus"] = f"{wind_speed} m/s"

    # Using Finnish weather description or OpenWeatherMap condition
    combined_data["Säätila"] = data2.get("säätila") or data1.get("säätila")

    return combined_data