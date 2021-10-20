import pyodbc
import pandas as pd
from datetime import timedelta, datetime, date
import smtplib
import mimetypes
from email.message import EmailMessage
#import keyring
from argparse import ArgumentParser
import os
from email import encoders
from email.message import Message
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pretty_html_table import build_table
import sys
from dateutil import parser

LMS_db_config = 'DRIVER={SQL Server};SERVER=AC-PC-097;DATABASE=FinanceTeamDB;Trusted_Connection=yes;'
cnxn = pyodbc.connect(LMS_db_config)

today = datetime.today()
startdate = today - timedelta(1)
startdate = format(startdate.strftime('%Y-%m-%d'))
today = format(today.strftime('%Y-%m-%d'))

#### DEV PURPOSES ONLY (No need to comment/uncomment anything here ###############################################################
SAIL_projection_query = '''
exec sp_SAILProjection_30dayout
'''

SAIL_Records_query = '''
select * from ##SAILSummary
'''

SAIL_projection_df = pd.read_sql(SAIL_projection_query, cnxn)  
SAIL_records_df = pd.read_sql(SAIL_Records_query, cnxn)  

records = []
for index,row in SAIL_projection_df.iterrows():
    start = parser.parse(row.ProjectionStartDate)
    end = (start + timedelta(30)).strftime('%Y-%m-%d')
    record = [row.ProjectionStartDate, end, row.TypeOfStore, row.Avg_LoanAmount, row.FV, row.Projection]

    records.append(record)

sail_locations = list(SAIL_records_df.StoreName.unique())

sail_locations.sort()

for location in sail_locations:
    SAIL_projection_query = f'''
    exec sp_SAILProjection_30dayout @store = '{location}'
    '''
    location_projection = pd.read_sql(SAIL_projection_query, cnxn)

    for index,row in location_projection.iterrows():
        start = parser.parse(row.ProjectionStartDate)
        end = (start + timedelta(30)).strftime('%Y-%m-%d')
        record = [row.ProjectionStartDate, end, location, row.Avg_LoanAmount, row.FV, row.Projection]

        records.append(record)

cnxn.close()
Projection_summary = pd.DataFrame(records, columns = ['Projection Start Date','Projection End Date', 'Store', 'Avg_LoanAmount_Past30Days', 'Projected_Loans_Funded_Future30Days', 'Projection'])
Projection_summary = Projection_summary.sort_values('Projection', ascending = False)
Projection_summary['Projected_Loans_Funded_Future30Days'] = Projection_summary['Projected_Loans_Funded_Future30Days'].apply(lambda x: "{:,.0f}".format((x/1)))
Projection_summary['Projection'] = Projection_summary['Projection'].apply(lambda x: "${:,.2f}".format((x/1)))
Projection_summary['Avg_LoanAmount_Past30Days'] = Projection_summary['Avg_LoanAmount_Past30Days'].apply(lambda x: "${:,.2f}".format((x/1)))


os.chdir(r'\\ac-hq-fs01\users$\daflores\My Documents\\')

Projection_summary.to_excel(f'ProjectionSummary_{today}.xlsx', index = False)

dfs = {'Projection': Projection_summary}
filename = f'ProjectionSummary2_{today}.xlsx'
writer = pd.ExcelWriter(filename, engine='xlsxwriter')
for sheetname, df in dfs.items():  # loop through `dict` of dataframes
    df.to_excel(writer, sheet_name=sheetname, index = False)  # send df to writer
    worksheet = writer.sheets[sheetname]  # pull worksheet object
    for idx, col in enumerate(df):  # loop through all columns
        series = df[col]
        max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
            len(str(series.name))  # len of column name/header
            )) + 1  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column width
writer.save()