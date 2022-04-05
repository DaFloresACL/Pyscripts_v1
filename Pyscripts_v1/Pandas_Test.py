import pandas as pd
import openpyxl
import pyodbc
import pypyodbc
import csv
import sqlalchemy
from sqlalchemy import create_engine
import os
import urllib

#os.chdir(r'\\ac-hq-fs01\users$\daflores\My Documents\\')
#os.getcwd()
#os.listdir()

#SAIL = pd.read_excel('SAIL_Data_PN.xlsx')

#SAIL['PN_Bins'] = pd.qcut(SAIL['PNValue'], q=11, precision = 4)

#lis = list(SAIL['PN_Bins'].unique())
#for i in lis:
#    print(i)

#SAIL['PN_Bins'].value_counts()

#SAIL.to_excel('SAIL_Bins.xlsx', index = False)

os.chdir(r'\\ac-hq-fs01\accounting\Finance\Underwriting\ValidiFi\Data Study\\')
os.getcwd()
os.listdir()

Validifi = pd.read_excel('Validifi Results Analysis 03.24.2022.xlsx')

Validifi = Validifi[Validifi['Funded']=='Y']

Validifi_missing = Validifi[Validifi['PI4Score'] == -1]
Validifi_missing['PI4_Bins'] = '[-1,-1]'

Validifi_withScores = Validifi[Validifi['PI4Score'] != -1]
Validifi_withScores['PI4_Bins'] = pd.qcut(Validifi_withScores['PI4Score'], q=10, precision = 4,duplicates='drop')

Validifi_Bins=pd.concat([Validifi_withScores,Validifi_missing])


lis = list(Validifi_Bins['PI4_Bins'].unique())
for i in lis:
    print(i)

Validifi_Bins['PI4_Bins'].value_counts()

Validifi_Bins.to_excel('Validifi_Bins.xlsx', index = False)