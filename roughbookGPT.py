import requests
from dotenv import load_dotenv
from os import path
import os

root = path.dirname(path.abspath(__file__))
dotenvpath = str(root) + '/weather-checker.env'
load_dotenv(dotenv_path=dotenvpath)

yahooapiclientid = os.getenv('yahooapiclientid')

def getweather(lat, lon, appid, outputformat):
    url = f"https://map.yahooapis.jp/weather/V1/place?coordinates={lon},{lat}&appid={appid}&output={outputformat}"
    response = requests.get(url)
    responsejson = response.json()
    print(responsejson)  # Print the entire JSON response
    temp_list = responsejson["Feature"][0]["Property"]["WeatherList"]["Weather"][0]["Temperature"]
    return temp_list


temperature = getweather(34.784835, 135.874422, yahooapiclientid, 'json')
print(f"The current temperature is {temperature}°C.")