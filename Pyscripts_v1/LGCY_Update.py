import pandas as pd
import openpyxl
import pyodbc
import pypyodbc

import sqlalchemy
from sqlalchemy import create_engine
import os
import urllib

pulldate = '12.01.2021'
os.chdir(rf'\\ac-hq-fs01\users$\daflores\My Documents\LGCY Python\{pulldate}\\')
os.getcwd()

fee_col = ['DB', 'LoanID', 'FeesID', 'Type', 'Date', 'Amount', 'Reference', 'Reference_', 'balance', 'CreateDate', 'UpdateDate']
tblFee = pd.read_csv(rf'LGCYtblFee_{pulldate}.csv', names=fee_col)

history_col = ['DB', 'HistoryID', 'CustomerID', 'Date', 'Loan', 'Location', 'Action', 'Notes', 'Name', 'CreatedID', 'Created', 'EditedID', 'Edited', 'Balance', 'msrepl_tran_version']
tblHistory = pd.read_csv(rf'LGCYtblHistory_{pulldate}.csv', names=history_col)

loans_col = ['DB', 'LoanID', 'CustomerID', 'Number', 'Date', 'LoanAmount', 'CheckNumber', 'CheckAmount', 'Type', 'Interest', 'Balance', 'Unpaid', 'LastDate', 'PaymentsMade', 'Refinance', 'Status', 'StatusChanged', 'Sent', 'ChargedOff', 'Location', 'ReceivedBy', 'Received', 'CreatedID', 'Created', 'EditedID', 'Edited', 'msrepl_tran_version', 'FirstPayment', 'LastPayment', 'Payment', 'Payments', 'PayFrequency', 'PayDate', 'PayDay1', 'PayDay2', 'PayFirst', 'PayLast', 'PayDay', 'ACH', 'Notify', 'Term', 'Rate', 'ACHYN', 'DueAmount', 'DueDays', 'Fees', 'PP', 'Fixed', 'Loan', 'Charge', 'TransactionID', 'Event']
tblLoans = pd.read_csv(rf'LGCYtblLoans_{pulldate}.csv', names = loans_col)

payments_col = ['DB', 'paymentid', 'LoanID', 'PostedDate', 'PaymentNo', 'TrxType', 'Reference', 'Reference_', 'Payment', 'Interest', 'Principal', 'Balance', 'Unpaid', 'CreateDate', 'UpdateDate']
tblPayments = pd.read_csv(rf'LGCYtblPayments_{pulldate}.csv', names = payments_col)


users_col = ['DB', 'UserID', 'User', 'Name', 'FirstName', 'LastName', 'LocationID', 'Location', 'Password', 'Code', 'Group', 'Active', 'LoansInstallment', 'CollectionsInstallment', 'LoansPayday', 'CollectionsPayday', 'LoansDays', 'CollectionsDays', 'Log', 'Logon', 'Logoff', 'Collector', 'User_ID', 'Location_ID', 'Last Edit', 'msrepl_tran_version', 'EmployeeNo'
]
tblusers = pd.read_csv(rf'LGCYtblUsers_{pulldate}.csv', names = users_col)

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
tblFee.to_sql('test_tblFee',con = cnxn,if_exists = 'append', index=False)
tblHistory.to_sql('test_tblHistory',con = cnxn,if_exists = 'append', index=False)
tblLoans.to_sql('test_tblLoan',con = cnxn,if_exists = 'append', index=False)
tblPayments.to_sql('test_tblPayments',con = cnxn,if_exists = 'append', index=False)
tblusers.to_sql('test_tblusers',con = cnxn,if_exists = 'replace', index=False)

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



