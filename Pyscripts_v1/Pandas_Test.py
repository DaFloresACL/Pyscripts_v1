import pandas as pd
import openpyxl
import pyodbc
import pypyodbc
import csv
import sqlalchemy
from sqlalchemy import create_engine
import os
import urllib

os.chdir(r'\\ac-hq-fs01\users$\daflores\My Documents\\')
os.getcwd()
os.listdir()

SAIL = pd.read_excel('SAIL_Data_PN.xlsx')

SAIL['PN_Bins'] = pd.qcut(SAIL['PNValue'], q=11, precision = 4)

lis = list(SAIL['PN_Bins'].unique())
for i in lis:
    print(i)

SAIL['PN_Bins'].value_counts()

SAIL.to_excel('SAIL_Bins.xlsx', index = False)