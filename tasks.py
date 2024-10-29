from robocorp.tasks import task
import requests
import json
import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

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
    print("Starting weather task...")  # Logging start
    city = "Helsinki"  # Change to the desired city
    api_key = os.getenv("OPENWEATHER_API_KEY")  # Store your API key in an environment variable

    # Debugging output to check if the API key is loaded
    print(f"Loaded API Key: {api_key}")  # This will show the API key or None

    if api_key is None:
        print("Error: OPENWEATHER_API_KEY environment variable is not set.")
    else:
        print("API Key loaded successfully.")
        weather_data = get_weather_data(city, api_key)
        if weather_data:
            print(json.dumps(weather_data, indent=4))
        else:
            print("Failed to retrieve weather data.")
