import pandas as pd
import openpyxl
import pyodbc
import pypyodbc
import numpy as np
import sqlalchemy
from   sqlalchemy import create_engine
import os
import urllib
import numpy_financial as npf

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

CustomerType_Bank_EAD_PDpercent= pd.read_sql('SELECT * FROM gds.CustomerType_Bank_EAD_PDpercent',cnxn)
CustomerType_Bank_PDpercent= pd.read_sql('SELECT * FROM gds.CustomerType_Bank_PDpercent',cnxn)
CustomerType_Bank_PDpercent_Loans= pd.read_sql('SELECT * FROM gds.CustomerType_Bank_PDpercent_Loans',cnxn)
CustomerType_LGDpercent= pd.read_sql('SELECT * FROM gds.CustomerType_LGDpercent',cnxn)
CustomerType_MaxPN= pd.read_sql('SELECT * FROM gds.CustomerType_MaxPN',cnxn)
EADpercent_LostRevenuePercent= pd.read_sql('SELECT * FROM gds.EADpercent_LostRevenuePercent',cnxn)
ScoreType_PDpercent= pd.read_sql('SELECT * FROM gds.ScoreType_PDpercent',cnxn)
FreqConv= pd.read_sql('SELECT * FROM FrequencyConversion',cnxn)

income_frequency = 'Weekly'
scoretype = 'V11'
score = 700
CustomerType = 'Existing'
BankAccount = 'PrePaid'
NumberOfLoans = 10
NetMonthIncome = 1515.50

income_multiplier = FreqConv.loc[(FreqConv['FreqString'] == income_frequency) ,'ConvToMonth'].item()



var_CustomerType_PNValue = 0.0000
if CustomerType == 'New':
    var_CustomerType_PNValue = 0.10
else:
    var_CustomerType_PNValue = (0.3214-0.1429) * (NumberOfLoans/25) + 0.1429

var_Score_PDPercent = ScoreType_PDpercent.loc[(ScoreType_PDpercent['ScoreType'] == scoretype) & (ScoreType_PDpercent.MinScore <= score) & (ScoreType_PDpercent.MaxScore >= score ),'PDpercent'].item()

var_Score_PDPercent_MinScore = ScoreType_PDpercent.loc[(ScoreType_PDpercent['ScoreType'] == scoretype) & (ScoreType_PDpercent.MinScore <= score) & (ScoreType_PDpercent.MaxScore >= score ),'MinScore'].item()

var_Score_PDPercent_Increments = ScoreType_PDpercent.loc[(ScoreType_PDpercent['ScoreType'] == scoretype) & (ScoreType_PDpercent.MinScore <= score) & (ScoreType_PDpercent.MaxScore >= score ),'IncrementDecrease'].item()

var_CustomerType_LGDPercent = CustomerType_LGDpercent.loc[(CustomerType_LGDpercent['CustomerType'] == CustomerType) ,'LGDpercent'].item()

var_CustomerType_Bank_PDpercent_Loans = CustomerType_Bank_PDpercent_Loans.loc[(CustomerType_Bank_PDpercent_Loans['CustomerType'] == CustomerType) & (CustomerType_Bank_PDpercent_Loans['BankAccount'] == BankAccount) & (CustomerType_Bank_PDpercent_Loans.Min_NumberofLoans <= NumberOfLoans) & (CustomerType_Bank_PDpercent_Loans.Max_NumberofLoans >= NumberOfLoans ),'PDpercent'].item()

var_CustomerType_Bank_PDpercent = CustomerType_Bank_EAD_PDpercent.loc[(CustomerType_Bank_EAD_PDpercent['CustomerType'] == CustomerType) & (CustomerType_Bank_EAD_PDpercent['BankAccount'] == BankAccount),'PDpercent'].item()

var_CustomerType_Bank_EADpercent = CustomerType_Bank_EAD_PDpercent.loc[(CustomerType_Bank_EAD_PDpercent['CustomerType'] == CustomerType) & (CustomerType_Bank_EAD_PDpercent['BankAccount'] == BankAccount),'EADpercent'].item()


var_EADpercent_LostRevenuePercent = EADpercent_LostRevenuePercent.loc[(EADpercent_LostRevenuePercent.Min_EADpercent <= var_CustomerType_Bank_EADpercent) & (EADpercent_LostRevenuePercent.Max_EADpercent >= var_CustomerType_Bank_EADpercent ),'LostRevenuePercent'].item()

var_Score_PDPercent + ((score - var_Score_PDPercent_MinScore) * var_Score_PDPercent_Increments)

round(npf.pmt(0.359999 / 12, 24 , -4200),2) / 1516.67 					

(0.3214-0.1429) * (3.000000/25.00000) + 0.1429

max_limit = 0.00

if CustomerType == 'New' and BankAccount != 'PrePaid':
    max_limit = 4000.00
elif CustomerType == 'New' and BankAccount == 'PrePaid':
    max_limit = 1000.00
elif CustomerType == 'Existing' and BankAccount != 'PrePaid':
    max_limit = 4000.00
else:
    max_limit = 6000.00

min_offer = 1000.00
max_offer = 1000.00
pn_offer = 0.00
max_payment = 100


while max_offer <= max_limit and pn_offer <= var_CustomerType_PNValue:
    pn_offer = round(npf.pmt(0.359999 / 12, 24 , -max_offer),2) / 1516.67
    if pn_offer > var_CustomerType_PNValue:
        max_offer = max_offer - 100
        pn_offer = round(npf.pmt(0.359999 / 12, 24 , -max_offer),2) / 1516.67
        break 
    else:
        max_offer = max_offer + 100

max_offer
var_CustomerType_PNValue
pn_offer



