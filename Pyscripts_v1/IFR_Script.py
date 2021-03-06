import pandas as pd
import numpy as np
import pyodbc

#period_end day after EOM
period_end='2022-03-01'
#period_start does not end
period_start='2016-11-01'

LGCY = pd.read_csv(r'\\ac-hq-fs01\users$\daflores\My Documents\python\Impairment LEGACY.csv')
LMS = pd.read_csv(r'\\ac-hq-fs01\users$\daflores\My Documents\python\Impairment LMS.csv',names=LGCY.columns,header=0)
LMS.info()
LGCY.info()
for i in range(len(LMS.columns)):
    if LMS.columns[i]==LGCY.columns[i]:
        pass  # print(LMS.columns[i],LGCY.columns[i],0)
    else:
        print('Check the following column in both sources: ', LMS.columns[i])
df=pd.concat([LGCY,LMS])
del LMS, LGCY
df.info()
df.LocNo=df.LocNo.apply(lambda x:str(x).zfill(5))

'''
Server = 'AC-PC-097'
DatabaseName = 'FinanceTeamDB'

conn_str = "DRIVER={SQL Server};Trusted_Connection=yes;SERVER=%s;Database=%s;" % (
    Server, DatabaseName)
cnxn = pyodbc.connect(conn_str)
'''

Locations=pd.read_csv(r'\\AC-PC-097\imports\Audit Impairement Imports\IFRS9\District Mapping.csv')
Locations.LocNo=Locations.LocNo.apply(lambda x:str(x).zfill(5))
Locations.loc[~Locations.Mapped.isna(),'Mapped']=Locations.loc[~Locations.Mapped.isna(),'Mapped'].apply(lambda x:str(x).zfill(5))
Locations

Loc=Locations.loc[(Locations.TenantID.isin([1,2,4,5]))&(Locations.LocType).isin([1,2]),['LocNo','TenantID','Region','ST','District']]
Loc.rename(columns={'ST':'State','LocNo':'BranchNo','Location':'Branch'},inplace=True)
Loc.Region=Loc.Region.astype('Int64')
Loc

df=pd.merge(df, Loc, left_on='LocNo', right_on='BranchNo')
df.ChargeOffDate=pd.to_datetime(df.ChargeOffDate)
df.Period=pd.to_datetime(df.Period)
df.shape
sorted(df.Location.unique())

# sanity check, this number should not change
# #df.loc[df.Period=='2021-01-31','RecoveryMTD'].sum()

# change dates per execution
period_end='2022-03-01'
period_start='2016-11-01'

df_output=df.loc[((df.ChargeOffDate<period_end)&(df.ChargeOffDate>=period_start))|((df.Period<period_end)&(df.Period>=period_start)&(df.RecoveryMTD!=0))].copy()

# sanity check, this number should not change
#df_output.loc[df_output.Period=='2021-01-31','RecoveryMTD'].sum()

df_output['Funding']=df_output.FundingMTD+df_output.RefinanceMTD
df_output['Principal']=df_output.PrincipalMTD+df_output.PrincipalRefi
df_output['Interest']=df_output.InterestMTD+df_output.InterestRefi


df_output.rename(columns={'ProductType':'Product','Location':'Branch'},inplace=True)
df_output.shape
df_output.columns
df_output=df_output[['Region','State','District','BranchNo','Branch','Product','FundDate','LoanNo','ChargeOffDate','PaidInFullDate', 'Period','CustomerPrincipal','Funding','Principal','Interest','WOAmount','RecoveryMTD']]
sorted(df_output['Branch'].unique())

df_output[df_output.ChargeOffDate.dt.to_period(freq ='M')==df_output.Period.dt.to_period(freq ='M')].groupby(by='Period')[['WOAmount','RecoveryMTD']].sum()

df_SAIL = df_output[df_output.Branch.isin(['SAIL IL'])]
df_SAIL[df_SAIL.ChargeOffDate.dt.to_period(freq ='M')==df_SAIL.Period.dt.to_period(freq ='M')].groupby(by='Period')[['WOAmount','RecoveryMTD']].sum()

df_output[df_output.Branch=='NiceLoans! OK'][df_output.ChargeOffDate.dt.to_period(freq ='M')==df_output.Period.dt.to_period(freq ='M')].groupby(by='Period')[['WOAmount','RecoveryMTD']].sum()

df_output[df_output.Branch!='SAIL IL'][df_output.Branch!='WLM SAIL'][df_output.Branch!='NiceLoans! OK'][df_output.ChargeOffDate.dt.to_period(freq ='M')==df_output.Period.dt.to_period(freq ='M')].groupby(by='Period')[['WOAmount','RecoveryMTD']].sum()


save_path = r'\\ac-hq-fs01\accounting\Finance\002 Areas\FinBond\FinBond Ad Hoc Projects\IFRS9 impairments\2022'
df_output.to_csv(save_path + '\\2022-02-28 IFRS9 impairments.csv',index=0)

