import pandas as pd
import sqlalchemy
from sqlalchemy import create_engine
import os
from dateutil import parser
from dateutil.relativedelta import relativedelta
import datetime 
import calendar
import keyring


username = 'prodSQLFinUserRO'
password = keyring.get_password('ACL.LMS',username)

today = datetime.datetime.today()
last_month = datetime.datetime.strptime(datetime.datetime.strftime(today + relativedelta(months=-1), '%Y-%m-01'),'%Y-%m-%d') + relativedelta(days=0)
last_day_of_last_month = calendar.monthrange(last_month.year,last_month.month)[-1]

last_month_eom = datetime.datetime.strptime(datetime.datetime.strftime(last_month + relativedelta(days=(last_day_of_last_month-1)  ), '%Y-%m-%d'),'%Y-%m-%d') + relativedelta(days=0)

eom = datetime.datetime.strftime(last_month_eom,'%Y-%m-%d')

engine = sqlalchemy.create_engine(f'mssql+pyodbc://{username}:{password}@secondaryacl.database.windows.net/acl.lms?driver=ODBC+Driver+17+for+SQL+Server')
connection = engine.connect()

query = f'''
SET NOCOUNT ON

declare @EOM date = eomonth('{eom}',0)
if object_id('tempdb..#Loans') is not null drop table #Loans
select	LoanID
INTO	#Loans
FROM	Loans a
WHERE	a.LoanDisbursementStatus = 2 -- funds are disbursed
and	convert(date,rpt.ToCst(a.DisbursementDate)) <= @EOM -- funds disbursed before EOM
and	isnull(rpt.ToCst(a.VoidedDate),'2099-12-31') > @EOM -- not voided or voided after EOM
and	isnull(rpt.ToCst(a.RescindedDate),'2099-12-31') > @EOM -- not rescinded or rescinded after EOM
Create	Clustered INDEX [IX_Loans] on #Loans (LoanID asc)



if object_id('tempdb..#MaxPay') is not null drop table #MaxPay
create	table #MaxPay (LoanID varchar(36), LastPayDate datetime)
insert	into #MaxPay
select	b.LoanID
	, max(TransactionDate) as LatestDate
from	#Loans a
Inner	join LoanTransactions b on a.LoanID = b.LoanID
where	1 not in (HasBeenOffset, IsAnOffset, HasBeenReversed)
and	TransactionTypeID in (6,7,8,9,10,26,36,42,57)
and posteddate <= '{eom}'
group by b.LoanID


select	@EOM [EOM]
	, h.Name [TenantName]
	, g.Abbreviation [ST]
	, c.Name [ProductName]
	, i.Name [ProductType]
	, a.LoanNo [LoanNo]
	, upper(b.Name_LastName + ', ' + b.Name_FirstName) [Customer]
	, d.Number [Loc]
	, a.APR
	, convert(date,rpt.ToCst(a.OriginationDate)) [Date]
	, convert(date,rpt.ToCst(a.DisbursementDate)) [Funded]
	, e.LoanAmount [Amount]
	, isnull(f.CurrentPrincipal,0) [CurrentPrincipal]
	, isnull(f.CurrentInterest,0) [CurrentInterest]
	, isnull(f.CurrentFees,0) [CurrentFees]
	, isnull(f.CurrentPrincipal + f.CurrentInterest + f.CurrentFees,0) [CurrentBalance]
	, f.IsPastDue [IsPastDue]
	, f.DaysPastDue [DaysPastDue]
	, f.AmountPastDue [AmountPastDue]
	, convert(date,rpt.ToCst(a.ChargedOffDate)) [ChargedOffDate]
	, convert(date,rpt.ToCst(a.PaidInFullDate)) [PaidInFullDate]
	, j.LastPayDate


from	#Loans loan
	inner join Loans a on loan.LoanID = a.LoanID
	inner join Customers b on b.CustomerID = a.CustomerID
	inner join Products c on c.ProductID = a.ProductID
	inner join Locations d on d.LocationID = a.OriginatingLocationID
	inner join LoanApplications e on e.LoanID = a.LoanID
	inner join HistoryMonthEndBalances f on f.LoanID = a.LoanID
	inner join UnitedStates g on g.StateID = d.StateOfBusinessID
	inner join Tenants h on h.TenantID = b.TenantID
	inner join AccountingClassificationTypes i on i.ID = c.AccountingClassificationTypeID
	left join #MaxPay j on j.LoanID = a.LoanID

where	a.LoanDisbursementStatus = 2 -- funds are disbursed
and	convert(date,rpt.ToCst(a.DisbursementDate)) <= @EOM -- funds disbursed before EOM
and	isnull(rpt.ToCst(a.VoidedDate),'2099-12-31') > @EOM -- not voided or voided after EOM
and	isnull(rpt.ToCst(a.RescindedDate),'2099-12-31') > @EOM -- not rescinded or rescinded after EOM
and	f.MonthEOM = month(@EOM) 
and	f.YearEOM = year(@EOM)

order by
	h.Name
	, g.Abbreviation
	, c.Name
	, a.LoanNo

--drop table #MaxPay'''

df = pd.read_sql(query, connection)  # PROD PURPOSES ONLY
connection.close()

#os.chdir(r'R:\2021\21-10')
os.chdir(r'R:\\')
main_dir = r'R:'
year_dir = fr'{main_dir}\{eom[0:4]}'
current_dir = fr'{year_dir}\{eom[2:7]}'
current_dir
try:
	os.mkdir(fr'{current_dir}\\')
except:
	pass
os.chdir(current_dir)
os.getcwd()

df.to_excel(f'EOM_Loans_{eom}.xlsx')
