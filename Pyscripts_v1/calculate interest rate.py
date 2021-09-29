import pandas as pd
import openpyxl
import pyodbc
import pypyodbc
import datetime
import sqlalchemy
from sqlalchemy import create_engine
import os
import urllib
from dateutil import parser
import math

os.chdir(r'\\ac-hq-fs01\users$\daflores\My Documents\\')
os.getcwd()
def rounddown(number:float, decimals:int=2):
    """
    Returns a value rounded down to a specific number of decimal places.
    """
    if not isinstance(decimals, int):
        raise TypeError("decimal places must be an integer")
    elif decimals < 0:
        raise ValueError("decimal places has to be 0 or more")
    elif decimals == 0:
        return math.floor(number)

    factor = 10 ** decimals
    return math.floor(number * factor) / factor

cnxn2 = pypyodbc.connect("Driver={SQL Server Native Client 11.0};"
                        "Server=AC-PC-097;"
                        "Database=FinanceTeamDB;"
                        "Trusted_Connection=yes;")

def create_NPER(disburse,df):
    l = ''
    z = 0
    daysbetween = [disburse]
    perbet = []

    for index,row in df.iterrows():
        daysbetween.append(row['datescheduled'])

    for i in daysbetween:
        if z == 0:
            l = disburse
        else:
            perbet.append((i-l).days)
            l = i
        z += 1
    return perbet

def create_paymentSch(df,column):
    paynumber = 1
    paymentsch = {}

    for index,row in df.iterrows():
        paymentsch[paynumber] = row[f'{column}']
        paynumber += 1
    
    return paymentsch

def create_amort(amount,payment,PCycles,APR,sch):
    apr = APR/365.0/100.0

    BegP = amount
    BegI = 0
    CurI = 0
    PMT = payment
    IPMT = 0
    PPMT = 0
    endP = 0
    endI = 0
    paymentsch = create_paymentSch(sch,'amount')


    n = 0
    dAMORT = {}
    for seq in PCycles:
        n += 1
        PMT = paymentsch[1]
        #print(n)
        BegP = BegP - PPMT
        BegI = BegP - IPMT

        CurI = BegP * seq * apr
        if PMT > BegI + CurI + BegP:
            PMT = PMT - (BegI + CurI + BegP)
            PPMT = BegP
            IPMT = BegI + CurI
            endP = 0
            endI = 0
        
        elif CurI < PMT:
            endI = 0
            IPMT = CurI
            PPMT = PMT - IPMT
            endP = BegP - PPMT
        else:
            IPMT = CurI
            endI = CurI - PMT
            endP = BegP
    
        dAMORT[n] = {'BegP': BegP,  'BegI': BegI, 'CurI': CurI, 'PMT': PMT, 'IPMT': IPMT, 'PPMT' : PPMT, 'endP': endP, 'endI': endI}
        if PMT > BegI + CurI + BegP:
            break
    return dAMORT

def dAMORT_TotalI(APRNew):
    dAMORT_cust = create_amort(amount,payment,PCycles,APRNew,sch)
    TotalIR = 0.00
    for CurI in dAMORT_cust:
        TotalIR += dAMORT_cust[CurI]['CurI']

    return TotalIR

def FindIR(i,NPER,P,A,dAMORT):
    E = rounddown(P * NPER - A,2)
    I0 = i
    E1 = dAMORT_TotalI(I0)
    step = 1
    
    # step 2
    I0 += step
    E2 = dAMORT_TotalI(I0)

    while abs(E - E2) > 1:

        x1 = E1 - E
        x2 = E2 - E
        
        if abs(x2) < abs(x1):
            pass
        elif abs(x2) > abs(x1):
            I0 -= step
            step = step * 0.1

        if x2 > 0:
            step = - abs(step)
        else:
            step = abs(step)
        
        I0 += step
        E1 = E2
        E2 = dAMORT_TotalI(I0)

        if E1 == E2:
            break
        
        i = I0

    return i



loans = pd.read_sql("select * from v_PosAPR_NegCalcIR",cnxn2)

Findings = []

for index,row in loans.iterrows():
    sch = pd.read_sql("select * from LMSScheduledPayments sch where LoanID = " + str(row['loanid'])[1:] + " order by DateScheduled",cnxn2)
    #second query here to make dictionary with key being LoanID and values list/dict of supplemental script
    disburse = row['loandisbursementdatetimecst']
    payments = pd.read_sql("select distinct amount from LMSScheduledPayments sch where LoanID = " + str(row['loanid'])[1:] ,cnxn2)
    APR = row['apr']
    amount = row['loanamount']
    payment = payments.iloc[-1][0]

    PCycles = create_NPER(disburse,sch)
    dAMORT = create_amort(amount,payment,PCycles,APR,sch)
    APRendP = dAMORT[len(dAMORT)]['endP']


    i = APR
    P = payment
    A = amount
    NPER = len(PCycles)
    IR = round(FindIR(i,NPER,P,A,dAMORT),15)
    IR

    test_dAMORT = create_amort(amount,payment,PCycles,IR,sch)
    IRendP = test_dAMORT[len(test_dAMORT)]['endP']
    os.chdir(r'\\ac-hq-fs01\users$\daflores\My Documents\Testing\\')
    v = pd.DataFrame.from_dict(test_dAMORT,orient='index')
    v.to_excel(f'{row.loanno}_{IR}_CalcIRSch.xlsx')
    os.chdir(r'\\ac-hq-fs01\users$\daflores\My Documents\\')
    #for i in range(len(test_dAMORT)):
    #pd.DataFrame.from_dict(test_dAMORT = create_amort(amount,payment,PCycles,IR,sch), orient='index')
    summary = [str(row.loanid)[2:-1],row.loanno,i,IR,APRendP,IRendP]
    Findings.append(summary)

df = pd.DataFrame(Findings, columns = ['LoanID','LoanNo','APR','CalculatedInterestRate','APR_EndingPrincipal','IR_EndingPrincipal'])
df.to_excel('Findings4.xlsx')

#sch2 = pd.read_sql("select * from LMSScheduledPayments sch where LoanID = '8D23BF5B-58E8-4F35-A90F-8D8B741150C2' order by DateScheduled",cnxn2)
#pay2 = pd.read_sql("select distinct amount from LMSScheduledPayments sch where LoanID = '8D23BF5B-58E8-4F35-A90F-8D8B741150C2' ",cnxn2)
#disburse2 = parser.parse('2021-01-13 00:00:00.000')
#APR2 = -694.2501111
#amount2 = 8000.00
#payment2 = 666.67
#PCycles2 = create_NPER(disburse2,sch2)
#dAMORT2 = create_amort(amount2,payment2,PCycles2,APR2,sch2)
#v = pd.DataFrame.from_dict(dAMORT2 , orient='index')

#l = ''
#z = 0
#daysbetween = [disburse2]
#perbet = []

#for index,row in sch2.iterrows():
#    row['datescheduled']
#    daysbetween.append(row['datescheduled'])

#for i in daysbetween:
#    if z == 0:
#        l = disburse2
#    else:
#        (i-l).days
#        perbet.append((i-l).days)
#        l = i
#    z += 1

#perbet
