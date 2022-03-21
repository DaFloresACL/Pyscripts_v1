import pandas as pd
import openpyxl
import pyodbc
import pypyodbc

import sqlalchemy
FROM gds.sqlalchemy import create_engine
import os
import urllib

os.getcwd()

conn_str = (
    r'Driver=ODBC Driver 17 for SQL Server;'
    r'Server= AC-PC-097;'
    r'Database=FinanceTeamDB;'
    r'Trusted_Connection=yes;'
)
quoted_conn_str = urllib.parse.quote_plus(conn_str)
engine = create_engine(f'mssql+pyodbc:///?odbc_connect={quoted_conn_str}')
cnxn = engine.connect()
#engine = create_engine('mssql+pyodbc://AC-PC-097/FinanceTeamDB')


scoretype = 'Vantage'
score = 500


CustomerType_Bank_EAD_PDpercent= pd.read_sql('SELECT * FROM gds.CustomerType_Bank_EAD_PDpercent',cnxn)
CustomerType_Bank_PDpercent= pd.read_sql('SELECT * FROM gds.CustomerType_Bank_PDpercent',cnxn)
CustomerType_Bank_PDpercent_Loans= pd.read_sql('SELECT * FROM gds.CustomerType_Bank_PDpercent_Loans',cnxn)
CustomerType_LGDpercent= pd.read_sql('SELECT * FROM gds.CustomerType_LGDpercent',cnxn)
CustomerType_MaxPN= pd.read_sql('SELECT * FROM gds.CustomerType_MaxPN',cnxn)
EADpercent_LostRevenuePercent= pd.read_sql('SELECT * FROM gds.EADpercent_LostRevenuePercent',cnxn)
ScoreType_PDpercent= pd.read_sql('SELECT * FROM gds.ScoreType_PDpercent',cnxn)

ScoreType_PDpercent.loc[(ScoreType_PDpercent['ScoreType'] == scoretype) & (ScoreType_PDpercent.MinScore <= score) & (ScoreType_PDpercent.MaxScore >= score ),'ID'].item()
