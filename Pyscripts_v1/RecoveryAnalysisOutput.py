import time
import pypyodbc
import pandas as pd
import numpy as np
import itertools
import xlsxwriter
import openpyxl
import datetime
import os

col = ['CST_PeriodStart', 'Tenant', 'CustomerType', 'States', 'AR', 'Secured', 'LTD_WO', 'LTD_Recovery', 'Loan_Active_Eventual_WO', 'Loan_Active_Eventual_Recovery', 'Month_WO_Amount', 'Month_Recovery_Amount', '00-06_Months_Recovery', '07-12_Months_Recovery', '01-02_Years_Recovery', '02-03_Years_Recovery', '03+_Years_Recovery']

df = pd.DataFrame(data= None, columns = col)

os.chdir(r'\\ac-hq-fs01\users$\daflores\My Documents\2022-03\\')

df = pd.read_excel('Recovery Analysis DATA 03.18.2022.xlsx')

file = 'RecoverySummary.xlsx'

file_create = os.getcwd() + '\\' + file
writer = pd.ExcelWriter(file_create)
tenants = df.Tenant.unique()
spacing = 0
for tenant in tenants:
    print(tenant)
    tenant_df = df[df['Tenant']==tenant]
    states = tenant_df.States.unique()
    for state in states:
        state_df = tenant_df[tenant_df['States']==state]
        customerTypes = state_df.CustomerType.unique()
        for customerType in customerTypes:
            if customerType == 'New':
                spacing = 0
            else:
                spacing = 65
            customer_df = state_df[state_df['CustomerType'] == customerType]
            customer_df['CustomerType'] = customerType
            table_df = np.round(pd.pivot_table(customer_df,values=['AR', 'Month_WO_Amount', 'Month_Recovery_Amount','LTD_WO', 'LTD_Recovery', '00-06_Months_Recovery', '07-12_Months_Recovery', '01-02_Years_Recovery', '02-03_Years_Recovery', '03+_Years_Recovery','TotalRecovery'],index = ['CustomerType','CST_PeriodStart'], aggfunc={'AR': np.sum, 'Month_WO_Amount': np.sum, 'Month_Recovery_Amount': np.sum, 'LTD_WO': np.sum, 'LTD_Recovery': np.sum, '00-06_Months_Recovery': np.sum, '07-12_Months_Recovery': np.sum, '01-02_Years_Recovery': np.sum, '02-03_Years_Recovery': np.sum, '03+_Years_Recovery': np.sum, 'TotalRecovery': np.sum},fill_value=0),2)
            table_df = table_df[['Month_WO_Amount', 'Month_Recovery_Amount', 'LTD_WO', 'LTD_Recovery','TotalRecovery',  '00-06_Months_Recovery', '07-12_Months_Recovery', '01-02_Years_Recovery', '02-03_Years_Recovery','03+_Years_Recovery']]
            table_df.to_excel(writer, sheet_name = tenant + '-' + state, startrow = spacing)
            spacing = len(df) + 5
            
writer.save()
writer.close()
