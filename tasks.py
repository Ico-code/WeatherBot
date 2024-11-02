from robocorp.tasks import task
import requests
import datetime
import os
from dotenv import load_dotenv
import traceback

from weather_bot import send_weather_email, log_email_send
from weatherInstitute import get_rounded_time, parse_weather_data
from dataHandling import getLocationsFromExcel, getRecipienEmails, sortDataByCity
from calculateAverages import calculateAverages
from errorHandling import logErrorToExcel

# Load environment variables from .env file
load_dotenv()

#defines error_info
error_info = {
    "errLVL": "Error",
    "errLocation": "tasks.py, line 25",
    "errMsg": "File not found",
}

@task
def weather_task():
    print("Starting weather task...")

    # try:
    #     raise ValueError("Testing");
    # except Exception as e:
    #     error_info["errLVL"] = "High";
    #     error_info["errMsg"] = str(e)        
        
    #      # Use traceback to capture the file name and line number
    #     tb = traceback.extract_tb(e.__traceback__)
    #     file_path, line_number = tb[-1].filename, tb[-1].lineno

    #     # Extract just the file name
    #     file_name = os.path.basename(file_path)
    #     error_info["errLocation"] = f"{file_name}, line {line_number}"     
        
    #     logErrorToExcel(error_info)
    # return;

    # Your settings
    api_key = os.getenv("OPENWEATHER_API_KEY")

    #Fetches locations and recipients from an excel documents
    excelFileEnd = ".xlsx"
    try:
        locations = getLocationsFromExcel(f"Locations{excelFileEnd}")
        recipient_emails = getRecipienEmails(f"Users{excelFileEnd}")
    except Exception as e:
        error_info["errLVL"] = "High";
        error_info["errMsg"] = str(e)        
        
         # Use traceback to capture the file name and line number
        tb = traceback.extract_tb(e.__traceback__)
        file_path, line_number = tb[-1].filename, tb[-1].lineno

        # Extract just the file name
        file_name = os.path.basename(file_path)
        error_info["errLocation"] = f"{file_name}, line {line_number}"     
        
        logErrorToExcel(error_info)

    #defines arrays for data
    openweather_data = [];
    weather_institute_data = [];
    weatherData_Averages = []

    # Check if API key is set
    if api_key is None:
        print("Error: OPENWEATHER_API_KEY environment variable is not set.")
        error_info["errLVL"] = "High";
        error_info["errMsg"] = "Error: OPENWEATHER_API_KEY environment variable is not set.";
        error_info["errLocation"] = "tasks.py, line 55";
        logErrorToExcel(error_info)

    for location in locations:
        # Fetch data from OpenWeatherMap
        if api_key is not None:
            openweather_data.append(get_weather_data(location, api_key))

        # Fetch data from Finnish wheatherInstitute
        weather_institute_data.append(fetch_weather_data_from_weatherInstitute(location, retries=3))

    if openweather_data and weather_institute_data:
        # Combine the data from both sources
        combined_weather_data = {
            "OpenWeatherMap": openweather_data,
            "Finnish Meteorological Institute": weather_institute_data
        }
        sortedData = sortDataByCity(combined_weather_data)

        for city in sortedData.values():
            # Combines the data from sources
            weatherData_Averages.append(calculateAverages(city))

        # Send email with combined data
        if send_weather_email(combined_weather_data, weatherData_Averages, recipient_emails):
            print("Email sent and logged successfully.")  # Success message
        else:
            print("Failed to send email.")  # Failure message
            error_info["errLVL"] = "Medium";
            error_info["errMsg"] = "Failed to send email";
            error_info["errLocation"] = "tasks.py, line 81";
            logErrorToExcel(error_info)
    else:
        print("Failed to retrieve weather data from one or both sources.")
        error_info["errLVL"] = "Medium";
        error_info["errMsg"] = "Failed to retrieve weather data from one or both sources.";
        error_info["errLocation"] = "tasks.py, line 61";
        logErrorToExcel(error_info)

def fetch_weather_data_from_weatherInstitute(city: str, retries: int):
    """
    This function fetches data from the finnish weatherinstitutes API
    
    Parameters:
        city (str): The city name to get weather data for.
        retries (int): The number of times this function tries to fetch data from the api

    Returns:
        dict: A dictionary containing city, temperature, weather condition, and wind speed.
    """

    # Get the current time in ISO 8601 format
    rounded_time = get_rounded_time()
    starttime = (rounded_time - datetime.timedelta(hours=1)).isoformat()  # 1 hour before
    endtime = (rounded_time + datetime.timedelta(minutes=10)).isoformat()  # 10 minutes after

    # Defines the API endpoint and parameters for the request
    url = "https://opendata.fmi.fi/wfs"
    params = {
        "service": "WFS",
        "version": "2.0.0",
        "request": "getFeature",
        "storedquery_id": "fmi::observations::weather::simple",
        "place": city,
        "parameters": "t2m,ws_10min,wawa",  # t2m: temperature, ws_10min: wind speed, wawa: weather state
         "starttime": starttime,
         "endtime": endtime,
    }

    attempt = 0
    while attempt <= retries:
        try:
            # Attempt the request
            response = requests.get(url, params=params, timeout=10)
            
            # If the response is successful (status code 200), parses and returns the data
            if response.status_code == 200:
                return parse_weather_data(response.text,city)
            else:
                print(f"Attempt {attempt + 1} failed with status code: {response.status_code}")
        
        except requests.RequestException as e:
            print(f"Attempt {attempt + 1} encountered an error: {e}")
        
        attempt += 1

    # If all attempts fail, raises an exception
    raise Exception("Failed to fetch data after multiple attempts")
  
# Defines the function to get weather data
def get_weather_data(city, api_key):
    """
    Fetch weather data for a specified city from OpenWeatherMap API.

    Parameters:
        city (str): The city name to get weather data for.
        api_key (str): Your OpenWeatherMap API key.

    Returns:
        dict: A dictionary containing city, temperature, weather condition, and wind speed.
    """
    url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}&units=metric"
    try:
        response = requests.get(url)
        response.raise_for_status()  # Raise an error for bad responses (4xx, 5xx)

        data = response.json()

        # Extract the required information
        temperature = data['main']['temp']
        weather_condition = data['weather'][0]['description']
        wind_speed = data['wind']['speed']

        # Return the data in the desired format
        return {
            "kaupunki": city,
            "lämpötila": str(temperature),  # Temperature as string
            "säätila": str(weather_condition),  # Weather condition as string
            "tuulen_nopeus": str(wind_speed)  # Wind speed as string
        }

    except requests.exceptions.HTTPError as http_err:
        print(f"HTTP error occurred: {http_err}")  # Log HTTP error
    except Exception as err:
        print(f"An error occurred: {err}")  # Log any other error

    return None  # Return None if an error occurs