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


ach_info = pd.read_csv(r'9000 customerdata_part1\ach_info.csv')
loan_info = pd.read_csv(r'9000 customerdata_part1\loan_info.csv')
schedule_info = pd.read_csv(r'9000 customerdata_part1\schedule_info.csv')
cust_info = pd.read_excel(r'part 2\cust_info.xlsx')
custnotes_info = pd.read_excel(r'part 2\custnotes_info.xlsx')
email_info = pd.read_csv(r'part 2\email_info.csv')
transactions_info = pd.read_csv(r'part 2\transactions_info.csv')
mail = pd.read_csv(r'9000 customerdata_part3\mail.csv')
sms_info = pd.read_csv(r'9000 customerdata_part3\sms_info.csv')
#cnxn2 = pypyodbc.connect("Driver={SQL Server Native Client 11.0};"
#                        "Server=secondary.database.windows.net;"
#                        "Database=ACL.LMS;"
#                        "Trusted_Connection=yes;")

conn_str = (
    r'Driver=ODBC Driver 17 for SQL Server;'
    r'Server= AC-PC-097;'
    r'Database=FinanceTeamDB;'
    r'Trusted_Connection=yes;'
)
quoted_conn_str = urllib.parse.quote_plus(conn_str)
engine = create_engine(f'mssql+pyodbc:///?odbc_connect={quoted_conn_str}')
cnxn = engine.connect()

engine = create_engine('mssql+pyodbc://AC-PC-097/FinanceTeamDB')
test = pd.read_sql('select top 10 LoanID from lmsloans', cnxn)
ach_info.to_sql('ACH_Info_1_stg',con = cnxn,if_exists = 'replace', index=False)
loan_info.to_sql('Loan_Info_1_stg',con = cnxn,if_exists = 'replace', index=False)
schedule_info.to_sql('Schedule_Info_1_stg',con = cnxn,if_exists = 'replace', index=False)
cust_info.to_sql('Cust_Info_2_stg',con = cnxn,if_exists = 'replace', index=False)
custnotes_info.to_sql('CustNotes_Info_2_stg',con = cnxn,if_exists = 'replace', index=False)
email_info.to_sql('Email_Info_2_stg',con = cnxn,if_exists = 'replace', index=False)
transactions_info.to_sql('Transactions_Info_2_stg',con = cnxn,if_exists = 'replace', index=False)
mail.to_sql('mail_3_stg',con = cnxn,if_exists = 'replace', index=False)
sms_info.to_sql('sms_Info_3_stg',con = cnxn,if_exists = 'replace', index=False)

#asrcall_col = ['UniqueID', 'Tenant', 'FirstName', 'LastName', 'HomePhone', 'WorkPhone', 'WorkExtension', 'CellPhone', 'Status', 'StoreNumber', 'StoreName', 'StorePhoneNumber', 'EmployerName', 'LastCallDate', 'ST', 'LoanType', 'Last4', 'Type', 'Date', 'Period']


#os.chdir(r'\\AC-PC-097\imports\\')
#os.getcwd()
#a = pd.read_csv('asr.csv', sep = '\t', lineterminator= '\r')

#a.to_sql('testing',con = cnxn,if_exists = 'append', index=False)

#b = pd.read_csv('coll.csv', sep = '\t', lineterminator= '\r')

#b['Home'] = b['Home'].str[1:]

#b.to_sql('testing2',con = cnxn,if_exists = 'append', index=False)

#c = pd.read_csv('ASRApp.csv', sep = '\t', lineterminator= '\r')
#c.to_sql('testing3',con = cnxn,if_exists = 'replace', index=False)
