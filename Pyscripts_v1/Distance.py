import geopy
import geopy.distance
import pandas as pd
import openpyxl
import pyodbc
import pypyodbc

import sqlalchemy
from sqlalchemy import create_engine
import os
import urllib

os.chdir(r'\\ac-hq-fs01\users$\daflores\My Documents\\')
os.getcwd()

cnxn2 = pypyodbc.connect("Driver={SQL Server Native Client 11.0};"
                        "Server=AC-PC-097;"
                        "Database=FinanceTeamDB;"
                        "Trusted_Connection=yes;")

LocDistanceDF = pd.read_sql('Select * from v_Location_Distance where homedistrict = awaydistrict',cnxn2)


DistanceDF = pd.read_excel('Distance.xlsx')

DistanceDF['KM'] = None
DistanceDF['Miles'] = None
i = 0
for index, row in DistanceDF.iterrows():
    coords_1 = (row['Loc1_Latitude'],row['Loc1_Longitude'])
    coords_2 = (row['Loc2_Latitude'],row['Loc2_Longitude'])
    km = geopy.distance.distance(coords_1, coords_2).km
    DistanceDF.at[index,'KM'] = km
    DistanceDF.at[index,'Miles'] = km * 0.621371


DistanceDF.to_excel('DistanceCalc.xlsx')

LocDistanceDF['KM'] = None
LocDistanceDF['Miles'] = None
i = 0
for index, row in LocDistanceDF.iterrows():
    coords_1 = (row['homelatitude'],row['homelongitude'])
    coords_2 = (row['awaylatitude'],row['awaylongitude'])
    km = geopy.distance.distance(coords_1, coords_2).km
    LocDistanceDF.at[index,'KM'] = km
    LocDistanceDF.at[index,'Miles'] = km * 0.621371

