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


#manual Test
income_frequency = 'Monthly'
scoretype = 'Vantage'
score = 811
CustomerType = 'New'
BankAccount = 'Non-PrePaid'
NumberOfLoans = 0
NetMonthIncome = 972.22

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

round(npf.pmt(0.359999 / 12, 24 , -1600),2) / NetMonthIncome					

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
max_payment = 0.00
min_payment = 0.00


while max_offer <= max_limit and pn_offer <= var_CustomerType_PNValue:
    pn_offer = round(npf.pmt(0.359999 / 12, 24 , -max_offer),2) / 1516.67
    if pn_offer > var_CustomerType_PNValue:
        max_offer = max_offer - 100
        max_payment = round(npf.pmt(0.359999 / 12, 24 , -max_offer),2)
        pn_offer = round(npf.pmt(0.359999 / 12, 24 , -max_offer),2) / 1516.67
        break 
    else:
        max_offer = max_offer + 100

max_offer
var_CustomerType_PNValue
pn_offer


dir = os.chdir(r'\\ac-hq-fs01\Public\Development\GDS')
test_cases = pd.read_excel('TestCases v2 03.24.2022'+ '.xlsx')

income_frequency = 'Weekly'
scoretype = 'V11'
score = 700
CustomerType = 'Existing'
BankAccount = 'PrePaid'
NumberOfLoans = 10
NetMonthIncome = 1515.50
results = []
min_PN = 0.00
test_cases_trial = test_cases[test_cases['Test Case'] == 49]
for index, row in test_cases.iterrows():
    if row['Customer Type'] == 'ExistingCustomer':
         CustomerType = 'Existing'
    else:
        CustomerType = 'New'
    if row['Bank Account'] == 'PayCard':
         BankAccount = 'PrePaid'
    else:
        BankAccount = 'Non-PrePaid'
    if CustomerType == 'New' and BankAccount != 'PrePaid':
        max_limit = 4000.00
    elif CustomerType == 'New' and BankAccount == 'PrePaid':
        max_limit = 1000.00
    elif CustomerType == 'Existing' and BankAccount == 'PrePaid':
        max_limit = 4000.00
    else:
        max_limit = 6000.00
    
    min_offer = 1000.00
    max_offer = 1000.00
    pn_offer = 0.00
    max_payment = 0.00
    min_payment = 0.00
    NumberOfLoans = row['Number of Loans']
    scoretype = row['Credit Score Type']
    score = row['Credit Score']
    NetMonthIncome = row['Monthly Net Income']
    CaseStatus = row.CaseStatus
    pn_offer = 0.00
    var_CustomerType_PNValue = 0.0000
    if CustomerType == 'New':
        var_CustomerType_PNValue = 0.10
    else:
        var_CustomerType_PNValue = (0.3214-0.1429) * (NumberOfLoans/25) + 0.1429
    print(var_CustomerType_PNValue)
    terms = [12,24]
    for term in terms:
        pn_offer = 0.00
        max_offer = 0.00
        if CaseStatus == 'Decline' or (scoretype == 'Vantage' and score < 525 and CustomerType == 'New' and BankAccount == 'PrePaid' and (score != None or score != 0)) or (scoretype == 'V11' and score < 685 and CustomerType == 'New' and BankAccount == 'PrePaid' and score != 111):
            CaseStatus = 'Decline'
            min_offer= 0
            min_payment = 0
            max_offer = 0
            max_payment = 0
            pn_offer = 0
        else:
            min_offer = 1000
            min_payment = round(npf.pmt(0.359999 / 12, term , -1000),2)
            min_PN = round(npf.pmt(0.359999 / 12, term , -1000),2) / NetMonthIncome
            while max_offer <= max_limit:
                if max_offer == max_limit:
                    max_offer = min(max_offer,max_limit)
                    pn_offer = round(npf.pmt(0.359999 / 12, term , -max_offer),2) / NetMonthIncome
                    max_payment = round(npf.pmt(0.359999 / 12, term , -max_offer),2)
                    break
                elif pn_offer > var_CustomerType_PNValue or max_offer > max_limit:
                    max_offer = max_offer - 100
                    max_payment = round(npf.pmt(0.359999 / 12, term , -max_offer),2)
                    pn_offer = round(npf.pmt(0.359999 / 12, term , -max_offer),2) / NetMonthIncome
                    break 
                else:
                    max_offer = max_offer + 100
                    pn_offer = round(npf.pmt(0.359999 / 12, term , -max_offer),2) / NetMonthIncome
                    max_payment = round(npf.pmt(0.359999 / 12, term , -max_offer),2)
        result = [row['Test Case'], CustomerType, BankAccount, scoretype, score, NetMonthIncome, term, CaseStatus, min_offer, min_payment, min_PN, max_offer, max_payment, pn_offer]
        results.append(result)

len(results)

df = pd.DataFrame(results,columns = ['TestCase#', 'CustomerType', 'BankAccount', 'scoretype', 'score','NetMonthIncome','TerminMonths', 'CaseStatus', 'MinOffer', 'MinPayment', 'Min_PN', 'Max_Offer', 'Max_Payment', 'Max_PN_Offer'])

df_12 = df[df.TerminMonths == 12]
df_24 = df[df.TerminMonths == 24]


df_all = df_12.merge(df_24, on='TestCase#', how='left')

df.to_excel(r'\\ac-hq-fs01\Public\Development\GDS\PythonTestCases.xlsx', index = False)


round(npf.pmt(0.359999 / 12, term , -1100),2)

