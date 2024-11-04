import re

def calculateAverages(weatherDataPerCity):
    """
    Merge data from different sources, averaging numerical values where possible and 
    keeping unique values otherwise.
    """
    temps = []
    windSpeeds = []
    weatherStates = []

    for city in weatherDataPerCity:
        # Extract and convert temperatures, handling any units (like "Celsius")
        temps.append(parse_temperature(city.get("lämpötila", "0")))
        windSpeeds.append(city.get("tuulen_nopeus", "0"))
        weatherStates.append(city.get("säätila", "Ei tietoa"))

    averagesInData = {
        "kaupunki" : city["kaupunki"],
        "Keskimääräinen lämpötila" : f"{sum(temps) / 2:.1f} Celsius",
        "Keskimääräinen Tuulen nopeus" : f"{sum_wind_speed(windSpeeds)} m/s",
        "Säätila" : weatherStates[1] if weatherStates is not None else "Unknown",
    }

    return averagesInData

def parse_temperature(temp_str):
    """
    Parses temperature data
    """
    # Remove any non-numeric parts (e.g., " Celsius") and convert to float
    try:
        return float(temp_str.split()[0])  # Extract numerical part if formatted like "20 Celsius"
    except (ValueError, IndexError):
        return None

def sum_wind_speed(data):
    total_speed = 0.0
    variable = data
    for tuulen_nopeus in data:
        match = re.search(r"[-+]?\d*\.?\d+", tuulen_nopeus)  # Finds the first decimal or whole number in the string
        
        # If a number is found, convert it to float and add it to total_speed
        if match:
            speed = float(match.group())
            total_speed += speed

    return f"{total_speed:1.1f}"