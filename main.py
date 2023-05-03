import os
import re
import requests
from pprint import pprint

from os import path
from dotenv import load_dotenv
root = path.dirname(path.abspath(__file__))
dotenvpath = str(root) + '/weather-checker.env'
load_dotenv(dotenv_path=dotenvpath)


clientid = os.getenv('apiclientid')

def create_query_string(params):
    query_string = ''
    for key, value in params.items():
        query_string += key + '=' + str(value) + '&'
    query_string = query_string[:-1] # remove the last '&'
    return query_string


def getweather(latitude, longitude, appid='', outputformat='json', past=2, intervals=5):
    coordinates_combined = str(latitude) + ',' + str(longitude)
    params = {
        'coordinates' : coordinates_combined, 
        'appid' : appid, 
        'output' : outputformat,
        'past' : past,
        'interval' : intervals,
        }
    
    query = create_query_string(params=params)
    rooturl = 'https://map.yahooapis.jp/weather/V1/place?'

    response = requests.get(rooturl + query)

    if response.status_code == 200:
        responsejson = response.json()
        weather_list = responsejson['Feature'][0]['Property']['WeatherList']['Weather']
        pprint(weather_list)
    else:
        print(f"Error: {response.status_code}")






getweather(latitude=135.9366703, longitude=34.7979113, appid=clientid, outputformat='json', past=2, intervals=5)
