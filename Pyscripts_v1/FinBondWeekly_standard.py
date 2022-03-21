'''
Analyst: Ralph Gozun
Date Created: February 25, 2021
Latest Update:  May 11, 2021

Warnings:
    * Manually executing this script will cause the latest copy of the reporting and summary excel workbooks to be overwritten.
    * Please ensure all Excel workbook instances are closed prior to running.  Bug currently exists which will modify any open Excel workbooks if this program is ran while a workbook is open.
    * I have noticed AmountReceived in Collections Progress Report (##PaymentCollection.AmountReceived) will change around prior to 10:30am.  Due to this, I suggest to run this report past that time.

Updates
    2021-09-22:
        * keyring package implemented to protect credential exposure

    2021-08-15:
        * Do not have any Excel workbooks open while running this script.  There may be possibility of packages going onto other open Excel instances and creating unwanted modifications to the file.  

    2021-06-20:
        * Fully manually executed this script on AC-PC-097 successfully at 8:50am.  I believe prior lock-up of scheduled execution on AC-PC-097 is due to traffic issues.
        * Discovered data has change from time of script run at 8:50am up to validation time at 9:25am.  Re-run of query reveals that DB sources have tables that run multiple times in a day.  Data that has changed is for AmountReceived for current week.

    2021-05-11:
        * Fully integrated WeeklyQuery_SQL.py so that Microsoft SSMS is no longer needed to instantiate temp-tables.
        * execute_server.bat now a stand-alone executable which will run the entire package from beginning to end.  Once completed, all reporting and summary excel workbooks will have been created in their respective directories.
        * Each reporting and summary file generated by this package will need to be manually validated.  Once validated, upload to Finbond FTP will need to be manually completed.
            * This is done to aid in data stewardship of reporting.
        * First code refactoring performed.

    2021-05-03:
        * Launched 'Weekly Reporting 20210429_001.py' for creating and updating reporting files.  Summary file also replicated but updating done manually.
'''
# PREPARE ENVIRONMENT
import pandas as pd
import numpy as np
import pyodbc
import shutil
import os
import os.path
import win32com.client as win32
import datetime
from datetime import datetime, timedelta
import timeit as ti
import keyring
from dateutil.relativedelta import relativedelta

# START CLOCKING RUNTIME FOR LOG
Runtime_start = ti.default_timer()

# # INSTANTIATE TEMP-TABLES IN SQL SERVER
# exec(open('//ac-hq-fs01/Accounting/Finance/004 Resources/Python/Finbond/Weekly Reporting/FinBondWeeklyQuery_SQL.py').read())

# PREPARE ODBC
## CONNECT TO AC-PC-097
Server = 'AC-PC-097'
DatabaseName = 'FinanceTeamDB'
FT_db_config = "DRIVER={SQL Server};Trusted_Connection=yes;SERVER=%s;Database=%s;" % (Server, DatabaseName)
cnxn = pyodbc.connect(FT_db_config)
del(Server, DatabaseName, FT_db_config)

# STAGE TEMP-TABLES TO LOCAL MEMORY
sql_PaymentCollection = 'SELECT * FROM ##PaymentCollection'
PaymentCollection = pd.read_sql(sql_PaymentCollection,cnxn)
del(sql_PaymentCollection)

sql_ExpectedCollectedReport = 'SELECT * FROM ##ExpectedCollectedReport'
ExpectedCollectedReport = pd.read_sql(sql_ExpectedCollectedReport,cnxn)
del(sql_ExpectedCollectedReport)

sql_NewClientFigures = 'SELECT * FROM ##NewClientFigures'
NewClientFigures = pd.read_sql(sql_NewClientFigures,cnxn)
del(sql_NewClientFigures)

sql_NewCustomerLoans = 'SELECT * FROM ##NewCustomerLoans'
NewCustomerLoans = pd.read_sql(sql_NewCustomerLoans,cnxn)
del(sql_NewCustomerLoans)

sql_SalesComparisonReport = 'SELECT * FROM ##SalesComparisonReport'
SalesComparisonReport = pd.read_sql(sql_SalesComparisonReport,cnxn)
del(sql_SalesComparisonReport)

sql_StoreSalesReport = 'SELECT * FROM ##StoreSalesReport'
StoreSalesReport = pd.read_sql(sql_StoreSalesReport,cnxn)
del(sql_StoreSalesReport)

# STAGE DATE DIMENSION
from datetime import date

Today = datetime(2022,3,20)  # For manual week.  This must match SQL being executed
#Today = datetime.today()
Today_nthDayOfYear = Today.timetuple().tm_yday

DateDim = pd.DataFrame(pd.date_range(Today-timedelta(days=Today_nthDayOfYear-1), periods=365).strftime('%Y-%m-%d'),columns=['Date'])
DateDim['Month'] = pd.to_datetime(DateDim.Date).dt.month
DateDim['Weekday'] = pd.to_datetime(DateDim.Date).dt.weekday # 0 is Monday
DateDim['Weekday_nthOfMonth'] = DateDim.groupby(['Month','Weekday']).cumcount()+1

DateDim_Sundays = DateDim[DateDim['Weekday'] == 6]
DateDim_Sundays = DateDim_Sundays.reset_index(drop=True)
Date_Sunday_CurrWk = DateDim_Sundays[DateDim_Sundays.Date < str(Today)].Date.max()
WeekAgo = Today-timedelta(days=7)
MonthAgo = Today - relativedelta(months = 1)
del(Today, Today_nthDayOfYear, DateDim)

## DETAILS FOR SUNDAYS OF INTERESTS
DateDim_Sunday_CurrWk = DateDim_Sundays[DateDim_Sundays.Date == Date_Sunday_CurrWk]
Date_Sunday_PriorMth = DateDim_Sundays[(DateDim_Sundays.Month==int(DateDim_Sunday_CurrWk.Month-1))
                                       & (DateDim_Sundays.Weekday_nthOfMonth==int(DateDim_Sunday_CurrWk.Weekday_nthOfMonth))
                                       ].Date.to_string(index=False).strip()
#Date_Sunday_PriorMth = '2021-09-26'  # use in event there are more Sundays this month than last month

# GET LAST YEAR DECEMBER FOR JANUARY REPORTING
Today_LastYear = datetime.today()
Today_LastYear = Today_LastYear.replace(Today_LastYear.year - 1)
Today_LastYear_nthDayOfYear = Today_LastYear.timetuple().tm_yday

DateDim_LastYear = pd.DataFrame(pd.date_range(Today_LastYear-timedelta(days=Today_LastYear_nthDayOfYear-1), periods=365).strftime('%Y-%m-%d'), columns=['Date'])
DateDim_LastYear['Month'] = pd.to_datetime(DateDim_LastYear.Date).dt.month
DateDim_LastYear['Weekday'] = pd.to_datetime(DateDim_LastYear.Date).dt.weekday  # 0 is Monday
DateDim_LastYear['Weekday_nthOfMonth'] = DateDim_LastYear.groupby(['Month', 'Weekday']).cumcount()+1

DateDim_LastYear_Sundays = DateDim_LastYear[DateDim_LastYear['Weekday'] == 6]
DateDim_LastYear_Sundays = DateDim_LastYear_Sundays.reset_index(drop=True)
DateDim_LastYear_Dec = DateDim_LastYear_Sundays[pd.to_datetime(DateDim_LastYear_Sundays.Date).dt.month == 12]

# APPEND DEC LAST YEAR TO CURRENT YEAR
DateDim_Sundays = DateDim_Sundays.append(DateDim_LastYear_Dec)
DateDim_Sundays = DateDim_Sundays.sort_values(by='Date')
DateDim_Sundays.reset_index(inplace=True, drop=True)

Date_Sunday_PriorWk = DateDim_Sundays[DateDim_Sundays.Date < str(Date_Sunday_CurrWk)].Date.max()
Date_Sunday_PriorMth = DateDim_Sundays[DateDim_Sundays.Date < str(MonthAgo)].Date.max()

# STAGE FILEPATH AND FILENAME VARIABLES
MainDir = r'//ac-hq-fs01/Accounting/Finance/002 Areas/FinBond/FinBond Monthly Reporting Package/'

## CURRENT WEEK PARAMETERS
Param_Year_CurrWk = str(datetime.strptime(Date_Sunday_CurrWk,'%Y-%m-%d').year)
Param_Month_CurrWk = str(datetime.strptime(Date_Sunday_CurrWk,'%Y-%m-%d').month).rjust(2,'0')
Param_nthWeek_CurrMth_CurrWk = str(int(DateDim_Sunday_CurrWk.Weekday_nthOfMonth))
WkDir_CurrWk = str(Param_nthWeek_CurrMth_CurrWk) + ' ' + 'WK ' + Date_Sunday_CurrWk
Filepath_Reporting_CurrMth = MainDir + Param_Year_CurrWk + '/' + Param_Month_CurrWk
Filepath_Reporting_CurrWk = MainDir + Param_Year_CurrWk + '/' + Param_Month_CurrWk + '/' + WkDir_CurrWk + '/'
Filepath_Summary_CurrWk = MainDir + Param_Year_CurrWk + '/' + Param_Month_CurrWk + '/Summary/'

## PRIOR WEEK PARAMETERS
Param_Year_PriorWk = str(datetime.strptime(Date_Sunday_PriorWk,'%Y-%m-%d').year)
Param_Month_PriorWk = str(datetime.strptime(Date_Sunday_PriorWk,'%Y-%m-%d').month).rjust(2,'0')
Param_nthWeek_CurrMth_PriorWk = DateDim_Sundays[DateDim_Sundays.Date == Date_Sunday_PriorWk].Weekday_nthOfMonth.to_string(index=False).strip()
WkDir_PriorWk = str(Param_nthWeek_CurrMth_PriorWk) + ' ' + 'WK ' + Date_Sunday_PriorWk
Filepath_Reporting_PriorWk = MainDir + Param_Year_PriorWk + '/' + Param_Month_PriorWk + '/' + WkDir_PriorWk + '/'
Filepath_Summary_PriorWk = MainDir + Param_Year_PriorWk + '/' + Param_Month_PriorWk + '/Summary/'

## PRIOR MONTH SAME WEEK PARAMETERS
Param_Year_PriorMth = str(datetime.strptime(Date_Sunday_PriorMth,'%Y-%m-%d').year)
Param_Month_PriorMth = str(datetime.strptime(Date_Sunday_PriorMth,'%Y-%m-%d').month).rjust(2,'0')
Param_nthWeek_PriorMth_SameWk = DateDim_Sundays[DateDim_Sundays.Date == Date_Sunday_PriorMth].Weekday_nthOfMonth.to_string(index=False).strip()
WkDir_PriorMth = str(Param_nthWeek_PriorMth_SameWk) + ' ' + 'WK ' + Date_Sunday_PriorMth
Filepath_Reporting_PriorMth = MainDir + Param_Year_PriorMth + '/' + Param_Month_PriorMth + '/' + WkDir_PriorMth + '/'
Filepath_Summary_PriorMth = MainDir + Param_Year_PriorMth + '/' + Param_Month_PriorMth + '/Summary/'

## STAGE REPORTING FILENAMES
ReportFilenamePrefix_CurrWk = Param_Year_CurrWk + ' ' + Param_Month_CurrWk + ' WK ' + Param_nthWeek_CurrMth_CurrWk + ' FinBond Weekly Reporting '
ReportFilenamePrefix_CurrWk_gm = Param_Year_CurrWk + ' ' + Param_Month_CurrWk + ' WK ' + Param_nthWeek_CurrMth_CurrWk + '  FinBond Weekly Reporting '
ReportFilenamePrefix_CurrWk_summary =  Param_Year_CurrWk + ' ' + Param_Month_CurrWk + ' WK ' + Param_nthWeek_CurrMth_CurrWk + ' FinBond Weekly Summary.xlsx'
ReportFilenamePrefix_PriorWk = Param_Year_PriorWk + ' ' + Param_Month_PriorWk + ' WK ' + Param_nthWeek_CurrMth_PriorWk + ' FinBond Weekly Reporting '
ReportFilenamePrefix_PriorWk_gm = Param_Year_PriorWk + ' ' + Param_Month_PriorWk + ' WK ' + Param_nthWeek_CurrMth_PriorWk + '  FinBond Weekly Reporting '
ReportFilenamePrefix_PriorWk_summary = Param_Year_PriorWk + ' ' + Param_Month_PriorWk + ' WK '  + Param_nthWeek_CurrMth_PriorWk + ' FinBond Weekly Summary.xlsx'
ReportFilenamePrefix_PriorMth_summary = Param_Year_PriorMth + ' ' + Param_Month_PriorMth + ' WK '  + Param_nthWeek_PriorMth_SameWk + ' FinBond Weekly Summary.xlsx'

## STAGE REPORTING FILENAME SUFFIXES
rpt_GM_suffix = 'GM (weekly).xlsx'
rpt_CollectionsProgress_suffix = 'Collections Progress Report.xlsx'
rpt_ExpectedCollected_suffix = 'Expected Collected Report.xlsx'
rpt_NewClientFigs_suffix = 'New Client Figures.xlsx'
rpt_NewCustomerLoans_suffix = 'New Customer Loans.xlsx'
rpt_StoreSalesComparison_suffix = 'Store Sales Comparison.xlsx'
rpt_StoreSalesReport_suffix = 'Store Sales Report.xlsx'

## STAGE WORKSHEET NAMES
sheet_CollectionsProgress = 'Collections Progress Report'
sheet_ExpectedCollected = 'Collections Progress Report'
sheet_NewClientFigs = 'New Client Figures'
sheet_NewCustomerLoans = 'New Customer Loans'
sheet_StoreSalesComparison = 'Sales Comparison Report'
sheet_StoreSalesReport = 'Store Sales Report'

# STAGE DF FOR PARAMTERS OF REPORTING WORKBOOKS
ReportParams = pd.DataFrame(columns=['files_suffix','worksheet','DBtable','Data_ColEnd1','Data_ColEnd2','Data_RowStart'])
ReportParams.files_suffix = [rpt_CollectionsProgress_suffix
                             ,rpt_ExpectedCollected_suffix
                             ,rpt_NewClientFigs_suffix
                             ,rpt_NewCustomerLoans_suffix
                             ,rpt_StoreSalesComparison_suffix
                             ,rpt_StoreSalesReport_suffix]
ReportParams.worksheet = [sheet_CollectionsProgress
                          ,sheet_ExpectedCollected
                          ,sheet_NewClientFigs
                          ,sheet_NewCustomerLoans
                          ,sheet_StoreSalesComparison
                          ,sheet_StoreSalesReport]
ReportParams.DBtable = ['PaymentCollection'
                       ,'ExpectedCollectedReport'
                       ,'NewClientFigures'
                       ,'NewCustomerLoans'
                       ,'SalesComparisonReport'
                       ,'StoreSalesReport']
ReportParams.Data_ColEnd1 = [36,36,10,20,8,18]
ReportParams.Data_ColEnd2 = ['AJ','AJ','J','T','H','R']
ReportParams.Data_RowStart = [1,1,1,1,4,1]

# COPY PRIOR WEEK FILES TO CURRENT WEEK AND RENAME FILENAMES RESPECTIVELY
## CREATE CURRENT MONTH'S DIRECTORY 
if os.path.isdir(Filepath_Reporting_CurrMth) == False:
    os.mkdir(Filepath_Reporting_CurrMth)
    os.mkdir(Filepath_Reporting_CurrMth + '/Summary')
else:
    pass

## CREATE CURRENT WEEK'S DIRECTORY
if os.path.isdir(Filepath_Reporting_CurrWk) == False:
    os.mkdir(Filepath_Reporting_CurrWk)
else:
    pass

# UPDATE CURRENT WEEK WORKBOOKS
## OPEN EXCEL FROM PYTHON
xl = win32.Dispatch("Excel.Application")

## UPDATE EACH WORKBOOK EXCEPT GM (FILENAME HAS 2X SPACE. MAINTAINING FILENAME IN CASE FINBOND HAS JOB DEPENDENT ON GM FILENAME.)
for i in range(0,len(ReportParams)):
    
    PriorWeekFile = Filepath_Reporting_PriorWk + ReportFilenamePrefix_PriorWk + ReportParams.files_suffix[i]
    CurrentWeekFile = Filepath_Reporting_CurrWk + ReportFilenamePrefix_CurrWk + ReportParams.files_suffix[i]
    CurrentWeekTempWB = Filepath_Reporting_CurrWk + 'tempwb.xlsx'
   
    ### COPY OVER PRIOR WEEK WOKRBOOKS
    shutil.copy(src = PriorWeekFile, dst = CurrentWeekFile)

    ### BEGIN UPDATING DATA ON EACH CURRENT FILE
    wb = xl.Workbooks.Open(CurrentWeekFile)
    xl.Visible = False
    
    ws = wb.Worksheets('Lead')
    ws.Cells(4,2).Value = Date_Sunday_CurrWk

    wb.RefreshAll()
    xl.CalculateUntilAsyncQueriesDone()
    
    ws = wb.Worksheets(ReportParams.worksheet[i])
    ws.Delete()

    ### CREATE TEMPWB FOR VALUES OF DATA WORKSHEET
    writer = pd.ExcelWriter(CurrentWeekTempWB)
    eval(ReportParams.DBtable[i]).to_excel(writer, ReportParams.worksheet[i], index=False)
    writer.save()

    ### MOVE TEMPWB.WORKSHEET TO REPORTING
    tempwb = xl.Workbooks.Open(CurrentWeekTempWB)
    tempwb.Worksheets(ReportParams.worksheet[i]).Select
    tempwb.Worksheets(ReportParams.worksheet[i]).Move(Before = wb.Worksheets('Data'))

    ws = wb.Worksheets('Lead')
    ws.Select()
    
    wb.Save()
    wb.Close()

os.remove(CurrentWeekTempWB)

## COPY GM FILE FROM PRIOR WEEK TO CURRENT WEEK DIRECTORY
PriorWeekFileGM = Filepath_Reporting_PriorWk + ReportFilenamePrefix_PriorWk_gm + rpt_GM_suffix
CurrentWeekFileGM = Filepath_Reporting_CurrWk + ReportFilenamePrefix_CurrWk_gm + rpt_GM_suffix
shutil.copy(src = PriorWeekFileGM, dst = CurrentWeekFileGM)

## BEGIN UPDATING CURRENT WEEK GM
wb = xl.Workbooks.Open(CurrentWeekFileGM)
xl.Visible = False
ws = wb.Worksheets('Lead')
ws.Cells(4,2).Value = Date_Sunday_CurrWk
wb.RefreshAll()
xl.CalculateUntilAsyncQueriesDone()
wb.Save()
wb.Close()
del(i,tempwb,wb,writer,ws)

# UPDATE SUMMARY FILE
## COPY SUMMARY FILE FROM LAST WEEK TO CURRENT WEEK DIRECTORY
shutil.copy(src = Filepath_Summary_PriorWk + ReportFilenamePrefix_PriorWk_summary
            ,dst = Filepath_Summary_CurrWk + ReportFilenamePrefix_CurrWk_summary)




## BEGIN UPDATING CURRENT WEEK SUMMARY
### COPY LAST MONTHS SAME WEEK ONTO COMPARISON COLUMN F

#### STAGE CURRENT WORKBOOK
WB_Summary_CurrMth = xl.Workbooks.Open(Filepath_Summary_CurrWk + ReportFilenamePrefix_CurrWk_summary)
xl.Visible = False

#### UPDATE CELLS OF COMPARISON COLUMN
WB_Summary_CurrMth.Worksheets('Weekly').Cells(3,6).Value = Date_Sunday_PriorMth
wb_temp = xl.Workbooks.Open(Filepath_Summary_PriorMth + ReportFilenamePrefix_PriorMth_summary)
rows_list = [10,16,13,19,22,25,30,33,38,40,44,47,52,54,58,61,68,71,78,83,86]
for row in rows_list:
    wb_temp.Worksheets('Weekly').Cells(row,4).Copy(WB_Summary_CurrMth.Worksheets('Weekly').Cells(row,6))
wb_temp.Close()

#### UPDATE CELLS OF CURRENT COLUMN
WB_Summary_CurrMth.Worksheets('Weekly').Cells(3,4).Value = Date_Sunday_CurrWk

##### FINBOND WEEKLY REPORTING GM (WEEKLY) - SECTION
wb_temp = xl.Workbooks.Open(CurrentWeekFileGM)
row_list_from = [11,22,11,22,11,22]
row_list_to = [10,16,13,19,22,25]
col_list_from = [17,17,38,38,22,22]
for i in range(0,6):
    WB_Summary_CurrMth.Worksheets('Weekly').Cells(row_list_to[i],4).Value = wb_temp.Worksheets('GM Report').Cells(row_list_from[i],col_list_from[i])
wb_temp.Close(SaveChanges=False)

##### FinBond Weekly Reporting Collections Progress Report
row_list = [30,33,38]
ws_list = ['InstallmentAmount','AmountReceived','Outstanding']
for i in range(0,3):
    WB_Summary_CurrMth.Worksheets('Weekly').Cells(row_list[i],4).Value = sum(eval('PaymentCollection.' + ws_list[i]))
WB_Summary_CurrMth.Worksheets('Weekly').Cells(40,4).Value = len(PaymentCollection)

##### FinBond Weekly Reporting Expected Collected Report
row_list = [44,47,52]
for i in range(0,3):
    WB_Summary_CurrMth.Worksheets('Weekly').Cells(row_list[i],4).Value = sum(eval('ExpectedCollectedReport.' + ws_list[i]))
WB_Summary_CurrMth.Worksheets('Weekly').Cells(54,4).Value = len(ExpectedCollectedReport)

##### FinBond Weekly Reporting New Client Figures
WB_Summary_CurrMth.Worksheets('Weekly').Cells(58,4).Value = sum(NewClientFigures.Capital)
WB_Summary_CurrMth.Worksheets('Weekly').Cells(61,4).Value = len(NewClientFigures)

##### FinBond Weekly Reporting New Customer Loans
WB_Summary_CurrMth.Worksheets('Weekly').Cells(68,4).Value = sum(NewCustomerLoans.Capital)
WB_Summary_CurrMth.Worksheets('Weekly').Cells(71,4).Value = len(NewCustomerLoans)

##### FinBond Weekly Reporting Store Sales Comparison
WB_Summary_CurrMth.Worksheets('Weekly').Cells(78,4).Value = len(StoreSalesReport)

##### FinBond Weekly Reporting Store Sales Report
WB_Summary_CurrMth.Worksheets('Weekly').Cells(83,4).Value = sum(StoreSalesReport.Capital)
WB_Summary_CurrMth.Worksheets('Weekly').Cells(86,4).Value = len(StoreSalesReport)

WB_Summary_CurrMth.Save()
WB_Summary_CurrMth.Close()

## QUIT
xl.Quit()

# END CLOCKING OF RUNTIME FOR LOG
Runtime_end = ti.default_timer()
Runtime = round(Runtime_end - Runtime_start,2)

# LOG OUTPUT
PackageCompletion = datetime.now()
print('FinBond Weekly data updating and report generating routine completed at {}. Runtime: {}'.format(PackageCompletion, Runtime), 'seconds')