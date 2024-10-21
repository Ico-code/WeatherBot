import requests

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

# Example usage
if __name__ == "__main__":
    api_key = "279bf9693b8a0adcddab6083c0b1d7b5"  # Replace with your actual OpenWeatherMap API key
    city = "Helsinki"
    weather_data = get_weather_data(city, api_key)
    
    if weather_data:
        print(weather_data)
