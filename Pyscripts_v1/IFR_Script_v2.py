
import pandas as pd
import numpy as np
import pyodbc
#a = pd.read_csv(r'\\ac-hq-fs01\accounting\Finance\002 Areas\FinBond\FinBond Ad Hoc Projects\IFRS9 impairments\2022\2022-02-28 IFRS9 impairments old.csv',nrows = 20)
#a.Branch
#period_end day after EOM
period_end='2022-03-01'
#period_start does not change
period_start='2016-11-01'

Server = 'AC-PC-097'
DatabaseName = 'FinanceTeamDB'

conn_str = "DRIVER={SQL Server};Trusted_Connection=yes;SERVER=%s;Database=%s;" % (
    Server, DatabaseName)
cnxn = pyodbc.connect(conn_str)


LGCY = pd.read_csv(r'\\ac-hq-fs01\users$\daflores\My Documents\python\Impairment LEGACY.csv')
LMS = pd.read_csv(r'\\ac-hq-fs01\users$\daflores\My Documents\python\Impairment LMS.csv',names=LGCY.columns,header=0)

Locations = pd.read_sql('''SELECT *, left(CASE WHEN Tenant like 'Best%' then 'NiceLoans!' else Tenant END,4) as Tenant_Adj FROM v_Location_Hierarchy WHERE GP_DB not in ('NLADV') and LocType in (1,2) and Tenant in ('SAIL','AmeriCash Loans','CreditBox','NiceLoans!','Best Loans')''',con =cnxn)

LMS.info()
LGCY.info()
#for i in range(len(LMS.columns)):
#    if LMS.columns[i]==LGCY.columns[i]:
#        pass  # print(LMS.columns[i],LGCY.columns[i],0)
#    else:
#        print('Check the following column in both sources: ', LMS.columns[i])
#df=pd.concat([LGCY,LMS])
#del LMS, LGCY
#df.info()
#df.LocNo=df.LocNo.apply(lambda x:str(x).zfill(5))

LMS['Location_Adj'] = LMS.Location
LMS['Location_Adj'] = LMS.Location_Adj.replace({'Online IL': 'Americash Loans','Online SC': 'Americash Loans','Online WI': 'Americash Loans','Online MO': 'Americash Loans'})
LMS['Location_Adj'] = LMS.Location_Adj.str[0:4]
LMS.Location_Adj.unique()

LGCY['Location_Adj'] = LGCY.Location
LGCY['Location_Adj'] = LGCY.Location_Adj.replace({'Online IL': 'Americash Loans','Online SC': 'Americash Loans','Online WI': 'Americash Loans','Online MO': 'Americash Loans'})
LGCY['Location_Adj'] = LGCY.Location_Adj.str[0:4]
LGCY.Location_Adj.unique()
Locations.LMS_LocNo = Locations.LMS_LocNo.fillna(0)
Locations.LEG_LocNo = Locations.LEG_LocNo.fillna(0)
Locations.LMS_LocNo = Locations.LMS_LocNo.apply(lambda x:str(x).zfill(5))
Locations.LEG_LocNo = Locations.LEG_LocNo.apply(lambda x:str(x).zfill(5))
Locations.LMS_LocNo=Locations.LMS_LocNo.astype(str).astype(float)
Locations.LEG_LocNo=Locations.LEG_LocNo.astype(str).astype(float)
Locations.LMS_LocNo.unique()
Locations.LEG_LocNo.unique()

#Locations=pd.read_csv(r'\\AC-PC-097\imports\Audit Impairement Imports\IFRS9\District Mapping.csv')
Locations.info()


Locations.rename(columns={'ST':'State','LMS_LocNo':'BranchNo_LMS','LEG_LocNo':'BranchNo_LEG','LocName':'Branch'},inplace=True)
#Locations.Region=Locations.Region.astype('Int64')

Locations.drop(['Branch'], axis = 1, inplace=True)

stg_df_lms=pd.merge(LMS, Locations, how='left', left_on=['LocNo','Location_Adj'], right_on=['BranchNo_LMS','Tenant_Adj'])
stg_df_lms.drop(['BranchNo_LEG'], axis = 1, inplace=True)
stg_df_lgcy=pd.merge(LGCY, Locations, how='left', left_on=['LocNo'], right_on=['BranchNo_LEG'])
stg_df_lgcy.drop(['BranchNo_LMS'], axis = 1, inplace=True)
stg_df_lgcy.rename(columns={'BranchNo_LEG':'BranchNo_LMS'},inplace=True)
del LMS, LGCY
df_lms=pd.concat([stg_df_lms,stg_df_lgcy])

df_lms.rename(columns={'BranchNo_LMS':'BranchNo'},inplace=True)
df_lms.info()

df_lms.ChargeOffDate=pd.to_datetime(df_lms.ChargeOffDate)
df_lms.Period=pd.to_datetime(df_lms.Period)
df_lms.shape
sorted(df_lms.Location.unique())
#df.ChargeOffDate=pd.to_datetime(df.ChargeOffDate)
#df.Period=pd.to_datetime(df.Period)
#df.shape
#sorted(df.Location.unique())

# sanity check, this number should not change
# #df.loc[df.Period=='2021-01-31','RecoveryMTD'].sum()

# change dates per execution
#period_end='2022-03-01'
#period_start='2016-11-01'
#df_output=df.loc[((df.ChargeOffDate<period_end)&(df.ChargeOffDate>=period_start))|((df.Period<period_end)&(df.Period>=period_start)&(df.RecoveryMTD!=0))].copy()

df_output=df_lms.loc[((df_lms.ChargeOffDate<period_end)&(df_lms.ChargeOffDate>=period_start))|((df_lms.Period<period_end)&(df_lms.Period>=period_start)&(df_lms.RecoveryMTD!=0))].copy()

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
df_output.to_csv(save_path + '\\2022-02-28 IFRS9 impairments v3.csv',index=0)

#a = pd.read_csv(save_path + '\\2022-02-28 IFRS9 impairments v3.csv', nrows=20)