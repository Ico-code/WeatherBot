import datetime
import xml.etree.ElementTree as ET

def get_rounded_time():
    """
    Gets the nearest 10-minute interval for fetching data from the finish weather institutes API
    """

    # Gets the current UTC time and rounds it down to the nearest 10-minute interval
    now = datetime.datetime.now(datetime.timezone.utc)
    rounded_minute = (now.minute // 10) * 10
    rounded_time = now.replace(minute=rounded_minute, second=0, microsecond=0)
    return rounded_time

def parse_weather_data(xml_data, location):
    """
    Parses XMl data from the request made to the finish weather insitute
    """

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
        "kaupunki": location,
        "lämpötila": temperature + " Celsius",
        "säätila": wawa_mapping.get(int(float(weather_state)), "Tuntematon säätila"),
        "tuulen_nopeus": wind_speed + " m/s avarage speed measured in the last 10 minutes",
    }