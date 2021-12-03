import pandas as pd
import sqlalchemy
from sqlalchemy import create_engine
import os
from dateutil import parser
from dateutil.relativedelta import relativedelta
import datetime 
import calendar
import keyring


#username = 'prodSQLFinUserRO'
#password = keyring.get_password('ACL.LMS',username)

today = datetime.datetime.today()

todaydate = datetime.datetime.strftime(today,'%Y-%m-%d')

LMS_db_config = 'DRIVER={SQL Server};SERVER=AC-PC-097;DATABASE=NiceLoans;Trusted_Connection=yes;'
cnxn = pyodbc.connect(LMS_db_config)

df = pd.read_sql('SET NOCOUNT ON; exec sp_NL_Vergent_Inactive_CallList', cnxn)  # PROD PURPOSES ONLY
cnxn.close()

#os.chdir(r'R:\2021\21-10')
os.chdir(r'\\ac-hq-ras\clarityftp$\prod\Noble\Export')
main_dir = r'\\ac-hq-ras\clarityftp$\prod\Noble\Export'
#year_dir = fr'{main_dir}\{eom[0:4]}'
#current_dir = fr'{year_dir}\{eom[2:7]}'
#current_dir
#try:
#	os.mkdir(fr'{current_dir}\\')
#except:
#	pass
#os.chdir(current_dir)
os.getcwd()

df.to_excel(f'NL_Vergent_Inactive_CallList_{todaydate}.xlsx', index = False)

