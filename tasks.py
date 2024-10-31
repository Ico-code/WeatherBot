from robocorp.tasks import task
import xml.etree.ElementTree as ET
import requests
import datetime
import json
import os
from dotenv import load_dotenv
from weather_bot import send_weather_email, log_email_send

# Load environment variables from .env file
load_dotenv()

@task
def minimal_task():
    location = "Helsinki"
    weather_info = fetch_weather_data_from_weatherInstitute(location, 1)
    print(weather_info)

def getLocationsFromExcel(excelFileName:str):
    locations = []
    return locations

def fetch_weather_data_from_weatherInstitute(location: str, retries: int):
    """This fetches function fetches data from the finnish weatherinstitutes API"""

    """
    
    Other Possibility for making this work: https://github.com/pnuu/fmiopendata?tab=readme-ov-file#download-and-parse-observation-data
    
    """

    # Get the current time in ISO 8601 format
    rounded_time = get_rounded_time()
    # starttime = (rounded_time).isoformat() + "Z"
    starttime = (rounded_time - datetime.timedelta(hours=1)).isoformat()  # 1 hour before
    endtime = (rounded_time + datetime.timedelta(minutes=10)).isoformat()  # 10 minutes after

    # Defines the API endpoint and parameters for the request
    url = "https://opendata.fmi.fi/wfs"
    params = {
        "service": "WFS",
        "version": "2.0.0",
        "request": "getFeature",
        "storedquery_id": "fmi::observations::weather::simple",
        "place": location,
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
                return parse_weather_data(response.text)
            else:
                print(f"Attempt {attempt + 1} failed with status code: {response.status_code}")
        
        except requests.RequestException as e:
            print(f"Attempt {attempt + 1} encountered an error: {e}")
        
        attempt += 1

    # If all attempts fail, raises an exception
    raise Exception("Failed to fetch data after multiple attempts")

def get_rounded_time():
    """Gets the nearest 10-minute interval for fetching data from the finish weather institutes API"""

    # Gets the current UTC time and round down to the nearest 10-minute interval
    now = datetime.datetime.now(datetime.timezone.utc)
    rounded_minute = (now.minute // 10) * 10
    rounded_time = now.replace(minute=rounded_minute, second=0, microsecond=0)
    return rounded_time

def parse_weather_data(xml_data):
    """Parses XMl data from the request made to the finish weather insitute"""

    # wawa-koodien sanakirja säätilojen kuvausten kanssa
    wawa_mapping = {
        0: "Ei merkittäviä sääilmiöitä",
        4: "Auerta, savua tai ilmassa leijuvaa pölyä, näkyvyys vähintään 1 km",
        5: "Auerta, savua tai ilmassa leijuvaa pölyä, näkyvyys alle 1 km",
        10: "Utua",
        20: "Sumua",
        21: "Sadetta (olomuoto on määrittelemätön)",
        22: "Tihkusadetta tai lumijyväsiä",
        23: "Vesisadetta (ei jäätävää)",
        24: "Lumisadetta",
        25: "Jäätävää vesisadetta tai tihkua",
        30: "Sumua havaintohetkellä",
        31: "Sumua tai jääsumua erillisinä hattaroina",
        32: "Sumua tai jääsumua, ohentunut edellisen tunnin aikana",
        33: "Sumua tai jääsumua, ilman muutoksia",
        34: "Sumua tai jääsumua, tullut sakeammaksi",
        40: "Sadetta havaintohetkellä",
        41: "Heikkoa tai kohtalaista sadetta",
        42: "Kovaa sadetta",
        50: "Tihkusadetta (heikkoa, ei jäätävää)",
        51: "Heikkoa tihkua, ei jäätävää",
        52: "Kohtalaista tihkua, ei jäätävää",
        53: "Kovaa tihkua, ei jäätävää",
        54: "Jäätävää heikkoa tihkua",
        55: "Jäätävää kohtalaista tihkua",
        56: "Jäätävää kovaa tihkua",
        60: "Vesisadetta (heikkoa, ei jäätävää)",
        61: "Heikkoa vesisadetta, ei jäätävää",
        62: "Kohtalaista vesisadetta, ei jäätävää",
        63: "Kovaa vesisadetta, ei jäätävää",
        64: "Jäätävää heikkoa vesisadetta",
        65: "Jäätävää kohtalaista vesisadetta",
        66: "Jäätävää kovaa vesisadetta",
        67: "Heikkoa lumensekaista vesisadetta tai tihkua (räntää)",
        68: "Kohtalaista tai kovaa lumensekaista vesisadetta tai tihkua (räntää)",
        70: "Lumisadetta",
        71: "Heikkoa lumisadetta",
        72: "Kohtalaista lumisadetta",
        73: "Tiheää lumisadetta",
        74: "Heikkoa jääjyvässadetta",
        75: "Kohtalaista jääjyväsadetta",
        76: "Kovaa jääjyväsadetta",
        77: "Lumijyväsiä",
        78: "Jääkiteitä",
        80: "Heikkoja kuuroja tai ajoittaista sadetta",
        81: "Heikkoja vesikuuroja",
        82: "Kohtalaisia vesikuuroja",
        83: "Kovia vesikuuroja",
        84: "Ankaria vesikuuroja (>32 mm/h)",
        85: "Heikkoja lumikuuroja",
        86: "Kohtalaisia lumikuuroja",
        87: "Kovia lumikuuroja",
        89: "Raekuuroja mahdollisesti yhdessä vesi- tai räntäsateen kanssa"
    }

    # Forms XML from string that is used for finding data
    root = ET.fromstring(xml_data)

    # Initializes variables to store values
    latest_time = None
    temperature = None
    wind_speed = None
    weather_state = None

    # Iterates through each <BsWfs:BsWfsElement> in the XML
    for element in root.findall(".//{*}BsWfsElement"):
        obs_time = element.find("{*}Time").text
        param_name = element.find("{*}ParameterName").text
        param_value = element.find("{*}ParameterValue").text

        # If this observation is more recent than the stored one, updates the values
        if latest_time is None or obs_time >= latest_time:
            latest_time = obs_time
            if param_name == "t2m":
                temperature = param_value
            elif param_name == "ws_10min":
                wind_speed = param_value
            elif param_name == "wawa":
                weather_state = param_value

    # Ensures we found all values, otherwise raise an exception
    if temperature is None or wind_speed is None or weather_state is None:
        raise ValueError("One or more required weather parameters were not found in the XML data")
    
    # Returns the formatted weather data dictionary for the most recent observation
    return {
        "lämpötila": temperature + " Celsius",
        "säätila": wawa_mapping.get(int(float(weather_state)), "Tuntematon säätila"),
        "tuulen_nopeus": wind_speed + " m/s avarage speed measured in the last 10 minutes",
    }
  
# Define the function to get weather data
def get_weather_data(city, api_key):
    """
    Fetch weather data for a specified city from OpenWeatherMap API.

    Parameters:
        city (str): The city name to get weather data for.
        api_key (str): Your OpenWeatherMap API key.

    Returns:
        dict: A dictionary containing temperature, weather condition, and wind speed.
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
            "lämpötila": str(temperature),  # Temperature as string
            "säätila": str(weather_condition),  # Weather condition as string
            "tuulen_nopeus": str(wind_speed)  # Wind speed as string
        }

    except requests.exceptions.HTTPError as http_err:
        print(f"HTTP error occurred: {http_err}")  # Log HTTP error
    except Exception as err:
        print(f"An error occurred: {err}")  # Log any other error

    return None  # Return None if an error occurs

@task
def weather_task():
    print("Starting weather task...")

    # Your settings
    city = "Helsinki"
    api_key = os.getenv("OPENWEATHER_API_KEY")
    recipient_emails = ["robin.salminen@student.laurea.fi"]

    # Check if API key is set
    if api_key is None:
        print("Error: OPENWEATHER_API_KEY environment variable is not set.")
        return

    # Fetch data from OpenWeatherMap
    openweather_data = get_weather_data(city, api_key)
    
    # Fetch data from the Finnish Weather Institute API
    weather_institute_data = fetch_weather_data_from_weatherInstitute(city, retries=3)

    if openweather_data and weather_institute_data:
        # Combine the data from both sources
        combined_weather_data = {
            "OpenWeatherMap": openweather_data,
            "Finnish Weather Institute": weather_institute_data
        }

        # Send email with combined data
        if send_weather_email(combined_weather_data, recipient_emails):
            print("Email sent and logged successfully.")  # Success message
        else:
            print("Failed to send email.")  # Failure message
    else:
        print("Failed to retrieve weather data from one or both sources.")