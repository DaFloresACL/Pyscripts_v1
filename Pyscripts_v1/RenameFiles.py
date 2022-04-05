import os
import datetime
import pause

dir = r'\\ac-hq-fs01\accounting\Finance\002 Areas\FinBond\FinBond Monthly Reporting Package\2022\03\Monthlies'
os.chdir(dir)
list = os.listdir(dir)

##for weekly
#for i in list:
#    os.rename(i,'2022 03 WK 5' + i[12:])
#for monthly
for i in list:
    os.rename(i,'2022 03' + i[7:])

    
#print(datetime.datetime.now())
#pause.until(datetime.datetime(2022,3,8,9,34))
    
#print(datetime.datetime.now())


#import pandas as pd
#import openpyxl
#import pyodbc
#import pypyodbc

#import sqlalchemy
#from sqlalchemy import create_engine
#import os
#import urllib

#pulldate = '03.01.2022'
#os.chdir(rf'T:\Finance\002 Areas\Nice Loans\Products\\')
#os.getcwd()

#conn_str = (
#    r'Driver=ODBC Driver 17 for SQL Server;'
#    r'Server= AC-PC-097;'
#    r'Database=NiceLoans;'
#    r'Trusted_Connection=yes;'
#)
#quoted_conn_str = urllib.parse.quote_plus(conn_str)
#engine = create_engine(f'mssql+pyodbc:///?odbc_connect={quoted_conn_str}')
#cnxn = engine.connect()
#engine = create_engine('mssql+pyodbc://AC-PC-097/NiceLoans')


#NL = pd.read_excel(rf'Product and Rates 2022.xlsx',sheet_name='Sheet3')
#NL.to_sql('temp5_NLProducts',con = cnxn,if_exists = 'append', index=False)
