import pandas as pd
import os
import googlemaps
from datetime import datetime
import csv
import json

def getApiKey():
    f = open('api_key.json')
    data = json.load(f)
    return data['googleApiKey']


def calculateDistanceUsingGoogleDirectionsApi(fromCity, toCity):
    gmaps = googlemaps.Client(key=getApiKey)
    now = datetime.now()
    directions_result = gmaps.directions(fromCity, toCity, mode="driving", departure_time=now)

    print(directions_result[0]['legs'][0]['distance']['text'])
    print(directions_result[0]['legs'][0]['duration']['text'])

    return directions_result[0]['legs'][0]['distance']['text']

def calculateDistanceUsingDistanceMatrix(fromCity, toCity):
    gmaps = googlemaps.Client(key=getApiKey())
    my_dist = gmaps.distance_matrix(fromCity, toCity, units='imperial')['rows'][0]['elements'][0]['distance']['text']

    return float(str(my_dist).replace(',', '').split(' ')[0])

def getSheetNames(templateFile):
    sheetNames = pd.ExcelFile(templateFile).sheet_names
    return sheetNames


def readExcel(excelFile, sheetName):
    xl = pd.read_excel(excelFile, sheet_name=sheetName)
    return xl

