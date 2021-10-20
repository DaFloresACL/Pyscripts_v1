'''------------------------------------------------------------------------------------------------------------------------
DEV NOTE:   Please search for the string 'DEV PURPOSES ONLY' and uncomment when working on this code.
            Also comment out their production counterparts which have the string 'PROD PURPOSES ONLY' associated with
            them on their code line.
------------------------------------------------------------------------------------------------------------------------'''
import pyodbc
import pandas as pd
from datetime import timedelta, datetime, date
import smtplib
import mimetypes
from email.message import EmailMessage
#import keyring
from argparse import ArgumentParser

from email import encoders
from email.message import Message
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pretty_html_table import build_table
import sys
# credentials

# connect to LMS
LMS_db_config = 'DRIVER={SQL Server};SERVER=AC-PC-097;DATABASE=FinanceTeamDB;Trusted_Connection=yes;'
cnxn = pyodbc.connect(LMS_db_config)

# set date param and run query
today = datetime.today()
startdate = today - timedelta(1)
startdate = format(startdate.strftime('%Y-%m-%d'))

#### DEV PURPOSES ONLY (No need to comment/uncomment anything here ###############################################################
sail_pif = '''
SET NOCOUNT ON;
DECLARE @ReportDate datetime = rpt.ToCST(getutcdate())
--DECLARE @ReportDate datetime = rpt.ToCST('2021-10-06')
DECLARE @DayOfWeek int = datepart(dw,@ReportDate)
DECLARE @Start datetime = dateadd(d,datediff(d,0,@ReportDate) - 1,0)
DECLARE @End datetime = dateadd(ms,-3,dateadd(d,1,@Start))

SET	@Start = dateadd(d
			, CASE @DayOfWeek
				WHEN 1 THEN - 2 -- if Sunday reporting date, start date is Friday morning
				ELSE 0
				END
			, @Start
			)

SET	@Start = rpt.ToUTC(@Start)
SET	@End = rpt.ToUTC(@End)

IF OBJECT_ID('tempdb..#Locs','U') IS NOT NULL DROP TABLE #Locs
SELECT	a.LocationID
	, 'SAIL' [Tenant]
INTO	#Locs
FROM	LMSLocations a
inner	join LMSCompanies b on b.CompanyID = a.CompanyID
WHERE	b.TenantID = 'A50F67F5-111B-45A9-BB40-2E356FAC482D' -- SAIL 

IF OBJECT_ID('tempdb..#Loans','U') IS NOT NULL DROP TABLE #Loans
SELECT	a.LoanID
INTO	#Loans
FROM	LMSLoans a
inner	join #Locs b on b.LocationID = a.OriginatingLocationID
inner	join LMSProducts c on c.ProductID = a.ProductID
--WHERE	c.Name like '%SAIL%' -- exclude WLM loan
--and	
where a.PaidInFullDate between @Start and @End
and	a.StatusID not in (11,12,14) -- converted
and	a.LoanDisbursementStatus = 2 -- funded

-- pulling in loan status table
IF OBJECT_ID('tempdb..#LnStatus','U') IS NOT NULL DROP TABLE #LnStatus
SELECT	a.ID
	, a.Name
INTO	#LnStatus
FROM	[SECONDARYACL].[ACL.LMS].[DBO].LoanStatus a

-- output
SELECT	a.LoanNo [Loan Number]
	, convert(date,rpt.ToCST(a.DisbursementDate)) [Fund Date]
	, convert(date,rpt.ToCST(a.PaidInFullDate)) [Paid-in-Full Date]
	, convert(date,rpt.ToCST(c.PostedDate)) [Last-Payment Date]
	, c.PaymentType [Payment Type]
	, d.[Name] [Loan Status]
	, (a.CurrentPrincipal + a.CurrentInterest + a.CurrentFees) [Current Balance]
FROM	LMSLoans a
inner	join #Loans b on b.LoanID = a.LoanID
inner	join	(
		SELECT	a.LoanID
			, ROW_NUMBER() OVER(PARTITION BY b.LoanID ORDER BY b.PostedDate DESC, b.TransactionDate DESC, b.RowNo DESC) [Row]
			, b.PostedDate
			, c.Name [PaymentType]
		FROM	#Loans a
		inner	join LMSLoanTransactions b on b.LoanID = a.LoanID
		inner	join LMSTransactionTypes c on c.ID = b.TransactionTypeID
		WHERE	b.HasBeenOffset = 0
		and	b.HasBeenReversed = 0
		and	b.IsAnOffset = 0
		and	c.IncludeInTotalPaid = 1
		) c on c.LoanID = a.LoanID
		and	c.Row = 1
inner	join #LnStatus d on d.ID = a.StatusID

DROP TABLE
	#Locs
	, #Loans
	, #LnStatus
'''
##################################################################################################################################
#df = pd.read_sql(legal_check, cnxn)  # DEV PURPOSES ONLY
df = pd.read_sql(sail_pif, cnxn)  # PROD PURPOSES ONLY

df['Current Balance'] = df['Current Balance'].apply(lambda x: "${:.2f}".format((x/1)))

body = 'This is a list of SAIL tenant loans that were Paid in Full yesterday.' +"""
<html>
<head>
</head>

<body>
        {0}
</body>

</html>
""".format(build_table(df, 'blue_light', width = '130px')) + '\n' + '\n Please do not respond to this email. Any concerns over the email, contact Daniel Flores, Ralph Gozun, or Sam Lingeman.' + '\n Thank you.'
# Set filename and filepath
#filename = 'SAIL_PIF_test' + str(date.today().strftime('%Y-%m-%d')) + '.xlsx'
#filepath = fr'\\ac-hq-fs01\accounting\Finance\003 Reporting\001 Daily Reporting\SAIL Daily Reporting\SAIL_PIF\{filename}'


# set email subject
def result(df):
    res = ''
    res_body = ''
    res_body = 'Please see the attached report for SAIL_PIF loans.  For any questions, please contact Ralph Gozun (847) 827-9740 Ext239 or Daniel Flores (847) 827-9740 Ext237'
        
    # Write to CSV if errors exist
    df.to_excel( filepath, index=False)
        
    # Guess the content type based on the file's extension then attach report
    ctype, encoding = mimetypes.guess_type(filepath)
    if ctype is None or encoding is not None:
        ctype = 'application/octet-stream'
    maintype, subtype = ctype.split('/', 1)
    #with open(filepath, 'rb') as fp:
    #    full_email.add_attachment(fp.read(), maintype=maintype, subtype=subtype, filename=filename)
    
    return res, res_body
# Stage SMTP message body
#msg = f'''
#Good morning,
    
#Attached is a summary on SAIL loans and their status that were paid in full yesterday.
    
#For any questions about this report, please contact Ralph Gozun, Daniel Flores, or Sam Lingeman.
   

#'''

# Assign SMTP message body with email container

#### Create email container ####
#full_email = EmailMessage()
##full_email['To'] = 'Ralph Gozun <rgozun@americashloans.net>'  # DEV PURPOSES ONLY
#full_email['To'] = 'DaFlores@americashloans.net'  # PROD PURPOSES ONLY
#full_email['From'] = 'DaFlores@americashloans.net'
#full_email['CC'] = 'DaFlores@americashloans.net'


#full_email.set_content(msg)
#full_email.add_attachment(filepath)
## Define email subject
#full_email['Subject'] = 'SAIL PIF Summary: Please see attachment.'

## Compile and Send the Message
#s = smtplib.SMTP('mail.americashloans.net', port=25)
#s.send_message( from_addr=full_email['From']
#               ,to_addrs=full_email['To']
#               ,msg=full_email
#               )
#s.quit()

# log output
#print('sail_pif:  {} for {}, executed at {}'.format(result(df)[0],startdate, today.strftime('%Y-%m-%d %H:%M:%S')))




#recipients = ['DaFlores@americashloans.net','SLingeman@americashloans.net'] 
#emaillist = [elem.strip().split(',') for elem in recipients]
msg2 = MIMEMultipart()
part1 = MIMEText(body, 'html')
#part0 = MIMEText(msg)
#msg2.attach(part0)
msg2.attach(part1)
msg2['Subject'] = "Yesterday's Paid in Full SAIL Loans"
msg2['From'] = 'DaFlores@americashloans.net'
#msg2['To'] = 'SLingeman@AmeriCashLoans.net'
#msg2['To'] = 'DaFlores@americashloans.net; SLingeman@AmeriCashLoans.net; mguenther@AmeriCashLoans.net; RHiatt@AmeriCashLoans.net; HSong@AmeriCashLoans.net'
msg2['To'] = 'DaFlores@americashloans.net'
#msg2['To'] = 'RHiatt@AmeriCashLoans.net; SLingeman@AmeriCashLoans.net; mguenther@AmeriCashLoans.net; DaFlores@americashloans.net; HSong@AmeriCashLoans.net;'

server = smtplib.SMTP('email.americashloans.net', port=25)

server.sendmail(msg2['From'],msg2['To'].split(";"), msg2.as_string())

server.quit()
