import os
import re
import requests
from pprint import pprint
from xslxfunctions import read_xlsx_file, createxlsxworkbook, write_xlsx_worksheet, rename_active_sheet, copy_sheet_contents, delete_xlsx_rows, delete_rows_with_duplicates, get_sheet_names
from os import path
from dotenv import load_dotenv
from datetime import date

date_string = str(date.today())



root = path.dirname(path.abspath(__file__))
dotenvpath = str(root) + '/weather-checker.env'
load_dotenv(dotenv_path=dotenvpath)

obubudatadir = os.getenv('obubudatadir')
yahooapiclientid = os.getenv('yahooapiclientid')

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
        temp_list = responsejson["Feature"][0]["Property"]["WeatherList"]["Weather"][0]["Temperature"]

        return [weather_list, temp_list]
        pprint(weather_list)
    else:
        print(f"Error: {response.status_code}")



def rain_check():

    teagardendataxlsx = 'TeaGardensData.xlsx'
    teagardendata = obubudatadir + teagardendataxlsx

    data = read_xlsx_file(teagardendata)

    gardennames = []

    for dict in data:
        gardennames.append(dict['Tea Garden Name'])

    xlsx_path_name_list = createxlsxworkbook(date_string=date_string, filename_prefix='TeaGardensData', sheet_names=gardennames, xlsx_folder_path='/Users/georgeguttridge-smith/code/obubu/obubu-data/')

    for dict in data:
        weather_list = getweather(dict['Latitude'], dict['Longitude'], yahooapiclientid)
        weather_list = [weather for weather in weather_list if weather.get('Type') != 'forecast']
        write_xlsx_worksheet(weather_list, dict['Tea Garden Name'], xlsx_path_name_list[0])

    rename_active_sheet(xlsx_path_name_list[0], 'Tea Gardens')

    copy_sheet_contents(teagardendata, 'Tea Gardens', xlsx_path_name_list[0], 'Tea Gardens')

    sheetnames = get_sheet_names(xlsx_path_name_list[0])

    for sheet in sheetnames:
        if sheet == 'Tea Gardens':
            delete_rows_with_duplicates(sheet, 'Tea Garden Name', xlsx_path_name_list[0])
        else:
            delete_rows_with_duplicates(sheet, 'Date', xlsx_path_name_list[0])
            print(f'Rows from {sheet} have been delete')










