import os
from os import path
from dotenv import load_dotenv
import requests
import time
import datetime
from xslxfunctions import read_xlsx_file, createxlsxworkbook, write_xlsx_worksheet, rename_active_sheet, copy_sheet_contents, get_sheet_names, delete_rows_with_duplicates

start_time = time.time()

root = path.dirname(path.abspath(__file__))
dotenvpath = str(root) + '/weather-checker.env'
load_dotenv(dotenv_path=dotenvpath)

apikey = os.getenv('openweatherapikey')
obubudatadir = os.getenv('obubudatadir')

date_string = datetime.date.today()

import requests

def extract_nested_dicts_lists(data_dict):
    new_data_dict = {}
    for key, value in data_dict.items():
        if isinstance(value, dict):
            for subkey, subvalue in extract_nested_dicts_lists(value).items():
                new_key = subkey + '_' + key
                new_data_dict[new_key] = subvalue
        elif isinstance(value, list):
            for i, item in enumerate(value):
                if isinstance(item, dict):
                    for subkey, subvalue in extract_nested_dicts_lists(item).items():
                        new_key = key + '_' + str(i) + '_' + subkey
                        new_data_dict[new_key] = subvalue
                else:
                    new_key = key + '_' + str(i)
                    new_data_dict[new_key] = item
        else:
            new_data_dict[key] = value
    return new_data_dict


def get_weather_data_current_forecast(lat, lon, appid, units='metric', exclude='minutely'):
    """
    Retrieves weather data from OpenWeather One Call API for a specified location.
    
    Args:
    - lat (float): Latitude of the location
    - lon (float): Longitude of the location
    - appid (str): Your OpenWeather API key
    - units (str): Units of measurement (default: metric)
    - exclude (str): Data to exclude from the response (default: minutely)
    
    Returns:
    - dict: A dictionary containing weather data for the specified location
    """
    
    url = f'https://api.openweathermap.org/data/2.5/onecall?lat={lat}&lon={lon}&exclude={exclude}&appid={appid}&units={units}'
    
    response = requests.get(url)
    data = response.json()
    
    return data

def get_weather_data_historical(lat, lon, appid, time, units='metric'):
    """
    Retrieves historical weather data from OpenWeather One Call API for a specified location.
    
    Args:
    - lat (float): Latitude of the location
    - lon (float): Longitude of the location
    - appid (str): Your OpenWeather API key
    - time (int): Timestamp Unix time UTC Zone 
    
    (Data is available from January 1st, 1979)
    
    Returns:
    - dict: A dictionary containing weather data for the specified location and time
    """
    
    url = f'https://api.openweathermap.org/data/3.0/onecall/timemachine?lat={lat}&lon={lon}&dt={time}&appid={appid}&units={units}'
    
    response = requests.get(url)
    data = response.json()
    
    return data

def replace_weather_data_with_desc(data_dict):
    # Get the list of weather data
    weather_data_list = data_dict['data']

    # Iterate over each item in the list
    for weather_data in weather_data_list:
        # Get the weather description
        weather_desc_list = weather_data['weather']
        weather_desc = weather_desc_list[0]['description']
        
        # Replace the weather dictionary with the weather description
        weather_data['weather'] = weather_desc

    # Return the modified dictionary
    return data_dict

def yyyy_mm_yy_unix(dates):
    unix_dates = []
    for date in dates:
        date_obj = datetime.datetime.strptime(date, '%Y-%m-%d')
        unix_time = int(date_obj.timestamp())
        unix_dates.append(unix_time)

    return unix_dates

def generate_date_list(startdate=19790101, enddate=20230506):
    start_date = datetime.datetime.strptime(str(startdate), '%Y%m%d')
    end_date = datetime.datetime.strptime(str(enddate), '%Y%m%d')

    delta = datetime.timedelta(days=1)
    dates = []
    unix_dates = []

    while start_date <= end_date:
        dates.append(start_date.strftime('%Y-%m-%d'))
        start_date += delta

    for date in dates:
        date_obj = datetime.datetime.strptime(date, '%Y-%m-%d')
        unix_time = int(date_obj.timestamp())
        unix_dates.append(unix_time)

    print(len(dates))

    # days = len(dates)
    # # gardens = len(gardenlist)
    # gardens = 1
    # free_calls = 1000
    # single_call_price_gbp = 0.0012
    # pound_yen_exchange = 170

    # total_api_calls = days * gardens
    # minus_free_calls = total_api_calls - free_calls
    # data_cost_gbp = minus_free_calls * single_call_price_gbp
    # data_cost_jpy = data_cost_gbp * pound_yen_exchange

    # print(f'The cost for {minus_free_calls} API calls in JPY is ¥{data_cost_jpy}')
    # print(f'The cost for {minus_free_calls} API calls in GBP is £{data_cost_gbp}')

    return [dates, unix_dates]

def sort_dicts_by_year(data_list):
    # Initialize an empty dictionary to store lists of dicts by year
    year_dict = {}

    # Iterate over each dict in the input list
    for data in data_list:
        # Convert the 'dt' Unix timestamp to a datetime object
        dt = datetime.datetime.utcfromtimestamp(data['dt'])

        # Extract the year from the datetime object
        year = dt.year

        # If this year is not yet a key in the dictionary, initialize it with an empty list
        if year not in year_dict:
            year_dict[year] = []

        # Append the current dict to the list corresponding to this year
        year_dict[year].append(data)

    return year_dict

def get_dict_by_value(list_of_dicts, key, value):
    """
    Retrieves a specific dictionary from a list of dictionaries based on a key-value pair.

    Args:
    - list_of_dicts (list): List of dictionaries to search through.
    - key (str): Key to check for the desired value.
    - value: Value to match in the dictionaries.

    Returns:
    - dict or None: The first dictionary that matches the specified key-value pair, or None if no match is found.
    """

    for dictionary in list_of_dicts:
        if dictionary.get(key) == value:
            return dictionary

    return None

def weather_check(dataxlsxname='', datadir='', start_date=20200101, end_date=20230101, teagarden=''):

    gardennames = []

    # read the garden data from excel
    teagardendata = datadir + dataxlsxname
    data = read_xlsx_file(teagardendata)
    for dict in data:
        gardennames.append(dict['Tea Garden Name'])
    # Get the names of all the gardens to loops through
    if teagarden == '':
        gardennames.append(teagarden)

    # Get the dates wanted
    date_list = generate_date_list(start_date, end_date)
    
    # Create the excel for the data to be able to be written too with a sheet for each of the gardesn
    xlsx_path_name_list = createxlsxworkbook(date_string=date_string, filename_prefix='OpenWeatherTeaGardensData', sheet_names=gardennames, xlsx_folder_path='/Users/georgeguttridge-smith/code/obubu/obubu-data/')
    if teagarden != '':
        recorded_data = read_xlsx_file(xlsx_path_name_list[0], teagarden)
        
        # remove already recorded dates
        recorded_dates = []
        for dict in recorded_data:
            recorded_dates.append(dict['dt'])
        for date in recorded_dates:
            if date in date_list[1]:
                date_list[1].remove(date)

    # get only the specified garden
    if teagarden != '':
        dict = get_dict_by_value(data, "Tea Garden Name", "Jinja")
        data = [dict]


    for dict in data:
        data_len = 0
        datelistlen = len(date_list[1])
        for date in date_list[1]:
            # print(f'Getting {date} Weather Data')
            weather_dict = get_weather_data_historical(dict['Latitude'], dict['Longitude'], apikey, date)
            # print(f'Got {date} Weather Data')
            # print(f'Extracting Data from dictionary')
            data_dict = extract_nested_dicts_lists(weather_dict['data'][0])
            # dates_data_list.append(data_dict)
            data_len += 1
            print(f'Date {data_len} of {datelistlen} recieved')
            # print(f'Writing {date} Data to excel')
            write_xlsx_worksheet(data_dict, dict['Tea Garden Name'], xlsx_path_name_list[0])
            # print(f'Written {date} Data')




    rename_active_sheet(xlsx_path_name_list[0], 'Tea Gardens')

    copy_sheet_contents(teagardendata, 'Tea Gardens', xlsx_path_name_list[0], 'Tea Gardens')

    sheetnames = get_sheet_names(xlsx_path_name_list[0])

    for sheet in sheetnames:
        if sheet == 'Tea Gardens':
            delete_rows_with_duplicates(sheet, 'Tea Garden Name', xlsx_path_name_list[0])
        else:
            delete_rows_with_duplicates(sheet, 'Date', xlsx_path_name_list[0])
            print(f'Rows from {sheet} have been deleted')

weather_data_dict = {'data': [{'clouds': 56,
                               'dew_point': -1.16,
                               'dt': 1586468027,
                               'feels_like': 6.86,
                               'humidity': 48,
                               'pressure': 1023,
                               'sunrise': 1586464268,
                               'sunset': 1586510643,
                               'temp': 9.11,
                               'visibility': 10000,
                               'weather': [{'description': 'broken clouds',
                                            'icon': '04d',
                                            'id': 803,
                                            'main': 'Clouds'}],
                               'wind_deg': 300,
                               'wind_gust': 0,
                               'wind_speed': 4.1,
                               'rain': {
                                   '1h' : 0.14,
                               }}],
                     'lat': 34.7848,
                     'lon': 135.8744,
                     'timezone': 'Asia/Tokyo',
                     'timezone_offset': 32400}

gardenlist = [
    {
        'teagarden': 'Michinashi', 
        '茶畑': '道無', 
        'latitude': 34.784835, 
        'longitude': 135.874422
    }, 
    {
        'teagarden': 'Ie Oku', 
        '茶畑': '家奥', 
        'latitude': 34.7835, 
        'longitude': 135.876888
    }, 
    {
        'teagarden': 'Ie Yoko', 
        '茶畑': '家横', 
        'latitude': 34.78379, 
        'longitude': 135.877282
    },
    {
        'teagarden': 'Ie no Mae', 
        '茶畑': '家', 
        'latitude': 34.783478, 
        'longitude': 135.877376
    },
    {
        'teagarden': 'Erihara Yoshida', 
        '茶畑': '撰原吉田', 
        'latitude': 34.787155, 
        'longitude': 135.89449
    },
    {
        'teagarden': 'Kouminkan', 
        '茶畑': '公民館', 
        'latitude': 34.782537, 
        'longitude': 135.877413
    },
    {
        'teagarden': 'Erihara', 
        '茶畑': '撰原', 
        'latitude': 34.781082, 
        'longitude': 135.888519
    },
    {
        'teagarden': 'Koshigoe', 
        '茶畑': '腰越', 
        'latitude': 34.795645, 
        'longitude': 135.872506
    },
    {
        'teagarden': 'Tenku', 
        '茶畑': '天空', 
        'latitude': 34.791104, 
        'longitude': 135.870918
    },
    {
        'teagarden': 'Tenku Oku', 
        '茶畑': '天空奥', 
        'latitude': 34.788837, 
        'longitude': 135.870455
    },
    {
        'teagarden': 'Monzen', 
        '茶畑': '門前', 
        'latitude': 34.801198, 
        'longitude': 135.92785
    },
    {
        'teagarden': 'AoiMori', 
        '茶畑': '青い森', 
        'latitude': 34.79791, 
        'longitude': 135.936664
    },
    {
        'teagarden': 'Kamo', 
        '茶畑': '加茂', 
        'latitude': 34.765443, 
        'longitude': 135.858955
    },
    {
        'teagarden': 'Kamikoma A', 
        '茶畑': '上狛A', 
        'latitude': 34.748556, 
        'longitude': 135.809004
    },
    {
        'teagarden': 'Kamikoma B', 
        '茶畑': '上狛B', 
        'latitude': 34.750011, 
    'longitude': 135.8087
    },
    {
        'teagarden': 'Minami', 
        '茶畑': '南', 
        'latitude': 34.792001, 
        'longitude': 135.905413
    },
    {
        'teagarden': 'Jinja', 
        '茶畑': '神社', 
        'latitude': 34.832394, 
        'longitude': 135.95307
    },
    {
        'teagarden': 'Jinja Oku', 
        '茶畑': '神社奥', 
        'latitude': 34.827453, 
        'longitude': 135.964907
    },
    {
        'teagarden': 'Shimojima', 
        '茶畑': '下島', 
        'latitude': 34.780163, 
        'longitude': 135.881876
    },
    {
        'teagarden': 'Somada', 
        '茶畑': '杣田', 
        'latitude': 34.788102, 
        'longitude': 135.913066
    },
    {
        'teagarden': 'Himeno', 
        '茶畑': '姫野', 
        'latitude': 34.813957, 
        'longitude': 135.891731
    },
    {
        'teagarden': 'Kawayoko', 
        '茶畑': '河横', 
        'latitude': 34.780703, 
        'longitude': 135.877667
    },
    {
        'teagarden': 'Minami Oku', 
        '茶畑': '南奥', 
        'latitude': 34.791195, 
        'longitude': 135.90727
    },
    {
        'teagarden': 'Pet Memorial', 
        '茶畑': 'ペットメモリアル', 
        'latitude': 34.78353, 
        'longitude': 135.879623
    },
    {
        'teagarden': 'Shōbō-dan no mae', 
        '茶畑': '消防団の前', 
        'latitude': 34.782308, 
        'longitude': 135.890133
    },
    {
        'teagarden': 'Prison', 
        '茶畑': 'プリズン', 
        'latitude': 34.774734, 
        'longitude': 135.882228
    }]

# michinashi = gardenlist[0]

# # data = get_weather_data_current_forecast(michinashi['latitude'], michinashi['longitude'], appid=apikey)
# data = get_weather_data_historical(michinashi['latitude'], michinashi['longitude'], apikey, 1586468027)

# ppprint(data)


# weather_data = replace_weather_data_with_desc(weather_data_dict)
# ppprint(weather_data)

weather_check('TeaGardensData.xlsx', obubudatadir, 20150101, 20230506, 'Jinja')

# delete_rows_with_duplicates('Michinashi', 'dt', '/Users/georgeguttridge-smith/code/obubu/obubu-data/2023-05-11OpenWeatherTeaGardensData.xlsx')



# extract_nested_dicts_lists(weather_data_dict)


# generate_date_list(20150101)




end_time = time.time()
elapsed_time_seconds = end_time - start_time
elapsed_time_minutes = elapsed_time_seconds / 60
print(f'Time taken: {elapsed_time_seconds} seconds')
print(f'Time taken: {elapsed_time_minutes} minutes')