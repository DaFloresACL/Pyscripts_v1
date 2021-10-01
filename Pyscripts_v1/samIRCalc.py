# import pandas as pd
import datetime as dt
import dateutil as du
import timeit as it
import math

from dateutil.relativedelta import relativedelta
from dateutil import parser
import math

#build schedule of dates
def SchDtRange(FundDt,FirstDt,NPER,PaymentFrequency):
    D1 = FundDt
    D2 = FirstDt
    Freq = PaymentFrequency

    # earliest period after fund date (D1)
    # number of initial units peirods
    results = FindStartDate(D1,D2,Freq)
    FirstIntitalDt = results[0]
    IntialPerCnt = results[1]

    # build scheduled payments range
    rng = {}
    for x in range(NPER):
        rng[x] = {}

        if x == 0:
            StartDate = D1
            EndDate = D2
        else:
            EndDate = FindNextDate(StartDate,1,Freq)

        rng[x]['Start'] = StartDate.strftime('%Y-%m-%d')
        rng[x]['End'] = EndDate.strftime('%Y-%m-%d')
        rng[x]['Days'] = (EndDate - StartDate).days

        # current end date / due date become start of next period
        StartDate = EndDate

    results = {}
    results['FirstIntitalDt'] = FirstIntitalDt
    results['IntialPerCnt'] = IntialPerCnt
    results['DateRange'] = rng

    return results

def RndUp(value:float,decimals:int):
    factor = 10 ** decimals
    return math.ceil(value * factor) / factor

# used to find initial periods, along with the earliest date on or after the fund date
def FindStartDate(D1,D2,Freq):
    D0 = D2
    t = 0

    # find the starting date
    while D0 >= D1:
        t += 1
        D0 = FindNextDate (D0,-1,Freq)
        # print(D0)

    # step back one period
    D0 = FindNextDate (D0,1,Freq)
    t -= 1

    return D0, t

# return the beginning of the month
def FindBoM(kDate,nMths):
    Mth = kDate.month + nMths
    Yr = kDate.year

    # adjust year if needed
    if Mth == 13:
        Mth = 1
        Yr += 1
    elif Mth == 0:
        Mth = 12
        Yr -= 1

    return dt.datetime(Yr,Mth,1)

# calculates the number of days between the kDate and nDate
def n_to_k(nDate,kDate,kType):
    if kType == 1:
        d = (nDate - kDate).days
    else:
        d = (kDate - nDate).days
    return d

# functions determines the next date for a payment frequency
def FindNextDate (kDate,kType,Freq):
    Mth = kDate.month
    Yr = kDate.year
    DoM = kDate.day # day of month
    DoW = kDate.weekday() # day of week
    WoM = RndUp(DoM / 7, 0)

    if Freq == 'Weekly':
        d = 7

    if Freq == 'Bi-Weekly':
        d = 14

    if Freq == 'Semi-Monthly':
        if DoM >= 15 and kType == 1:
            nDate = kDate + relativedelta(months=1) - dt.timedelta(14)
            d = n_to_k(nDate,kDate,kType)
        elif DoM < 15 and kType != 1:
            nDate = (kDate + dt.timedelta(14)) - relativedelta(months=1)
            d = n_to_k(nDate,kDate,kType)
        else:
            d = 14

    if Freq == 'Semi-Monthly EOM':
        if DoM > 15 and kType == 1:
            dlt = 14
            nMths = 1
        elif DoM <= 15 and kType == 1:
            dlt = -1
            nMths = 1
        elif DoM > 15 and kType != 1:
            dlt = 14
            nMths = 0
        else:
            dlt = -1
            nMths = 0

        nDate = FindBoM(kDate,nMths) + dt.timedelta(dlt)
        d = n_to_k(nDate,kDate,kType)

    if Freq == 'Semi-Monthly Variable':
        if DoM < 15 and kType != 1:
            nMths = -1
        elif DoM >= 15 and kType == 1:
            nMths = 1
        else:
            nMths = 0

        # beginning of month statistics and related adjustments
        BoM = FindBoM(kDate,nMths) # beginning of next month
        BDoW = BoM.weekday() # day of week for BoM
        Adj = DoW - BDoW

        if BDoW <= DoW:
            Adj -= 7 # adjustment for how the BoM falls during the week

        if DoM >= 15 and kType == 1:
            Adj += 7 * (WoM - 2) # adjust for the number of weeks
            nDate = BoM + dt.timedelta(Adj)
            d = n_to_k(nDate,kDate,kType)
        elif DoM < 15 and kType != 1:
            Adj += 7 * (WoM + 2) # adjust for the number of weeks
            nDate = BoM + dt.timedelta(Adj)
            d = n_to_k(nDate,kDate,kType)
        else:
            d = 14

    if Freq == 'Monthly':
        nDate = FindBoM(kDate,kType) + dt.timedelta(DoM - 1) 
        d = n_to_k(nDate,kDate,kType)

    if Freq == 'Monthly EOM':
        nDate = FindBoM(kDate,kType) + relativedelta(months=1) - dt.timedelta(1)
        d = n_to_k(nDate,kDate,kType)

    if Freq == 'Monthly Variable':
        # beginning of month statistics and related adjustments
        BoM = FindBoM(kDate,kType) # beginning of next month
        BDoW = BoM.weekday() # day of week for BoM
        Adj = DoW - BDoW

        if BDoW <= DoW:
            Adj -= 7 # adjustment for how the BoM falls during the week

        Adj += 7 * WoM # adjust for the number of weeks
        nDate = BoM + dt.timedelta(Adj)
        d = n_to_k(nDate,kDate,kType)

    if kType != 1:
        d = d * -1

    DN = kDate + dt.timedelta(d)
    return DN

# list of variables
    # A = loan amount
    # I = annual rate
    # i = period rate (I/n)
    # NPER = number of payments
    # NPY = number of payments per year
    # P = payment amount
    # F = fractional period adjustment
    # freq = payment frequency
    # u = unit period days
    # t = unit periods before first payment
    # t0 = first periods after fund date
    # D1 = fund date
    # D2 = next payment date

PRCSN = 15

# annual payments dictionary
# freq description: n, freq type, unit period days
PaymentFrequency = {
    'Weekly': [52,'7D',7]
    , 'Bi-Weekly': [26,'14D',14]
    , 'Semi-Monthly': [24,'SMS',15]
    , 'Semi-Monthly EOM': [24,'SM',15]
    , 'Semi-Monthly Variable': [24,'2W',15]
    , 'Monthly': [12,'MS',30]
    , 'Monthly EOM': [12,'M',30]
    , 'Monthly Variable': [12,'MS',30]   
}

# function to round down to the nears decimal point
# https://kodify.net/python/math/round-decimals/
def MyRndDwn(value:float,decimals:int):
    factor = 10 ** decimals
    return math.floor(value * factor) / factor

# calcualte i (period rate)
def PeriodRate (I,NPY):
    return I / NPY

# guess the loan's payment using the standard payment formula
# adjusting for fractional periods
def Payment (A,I,NPER,NPY,F):
    i = PeriodRate(I,NPY)
    return round(A * (i * (1 + i) ** NPER) / ((1 + F * i) * (1 + i) ** NPER - 1),2)

# calculate the PV of an ordinary annuity of payments with fractional first periods
def RegZPV(I,NPER,NPY,P,F,t):
    i = PeriodRate(I,NPY)
    a = P * ((1 - (1 / (1 + i) ** -NPER)) / i) # PV of an ordinary annuity
    b = (1 + F * i) * (1 + i) ** t * (1 + i) ** (NPER) # general equation
    return (- a / b)

# calculate needed fraction days for first unit period
def FractionDays(D1,D2,u,freq):
    U1 = (D2 - D1).days

    # find the number of months and initial period date
    m = (D2.year - D1.year) * 12 + (D2.month - D1.month)
    D0 = (D2 + dt.timedelta(1)) - relativedelta(months = m) - dt.timedelta(1)
    if D0 < D1:
        D0 = (D0 + dt.timedelta(1)) + relativedelta(months = 1) - dt.timedelta(1)
        m -= 1
    
    # actual days (e.g., weekly and bi-weekly)
    if freq == 'Weekly' or freq == 'Bi-Weekly':
        t = MyRndDwn(U1 / u,0)
        f = U1 - u * t 

    if freq == 'Semi-Monthly' or freq == 'Semi-Monthly Variable' or freq == 'Semi-Monthly EOM':
        if m == 0:
            t = MyRndDwn(U1 / u,0)
            f = U1 - u * t
        else:
            f0 = (D0 - D1).days + m * 30
            t = MyRndDwn(f0 / u,0)
            f = f0 - t * u

    if freq == 'Monthly' or freq == 'Monthly Variable'or freq == 'Monthly EOM':
        f = (D0 - D1).days
        t = m

    t -= 1

    results = {}
    results['F'] = f / u
    results['t'] = t
    
    return results

# build amortization table
def BuildAmort (rng):
    # build nested dictionary for amortization table
    NPER = len(rng)
    dct = {}

    for a in range(NPER):
        # create the variable
        dct[a] = {}

        dct[a]['Payment Date'] = rng[a]['End']
        dct[a]['Days'] = rng[a]['Days']
        dct[a]['BegP'] = 0.00
        dct[a]['BegI'] = 0.00
        dct[a]['CurI'] = 0.00
        dct[a]['PMT'] = 0.00
        dct[a]['IPMT'] = 0.00
        dct[a]['PPMT'] = 0.00
        dct[a]['EndP'] = 0.00
        dct[a]['EndI'] = 0.00
    
    return dct

# find the payment for the current APR
def FindPayment(I,NPER,NPY,P,F,A,t):
    A0 = RegZPV(I,NPER,NPY,P,F,t)
    P0 = P

    while A0 != A:
        # step 1
        A1 = RegZPV(I,NPER,NPY,P,F,t)

        # step 2
        if A1 == A:
            pass
        else:
            P += 0.01

        A2 = RegZPV(I,NPER,NPY,P,F,t)    

        # step 3
        if A2 == A:
            pass
        else:
            P0 += 0.01 * ((A - A1) / (A2 - A1))

        P = P0
        A0 = round(RegZPV(I,NPER,NPY,P,F,t),2)

    return round(P,2)

# find the APR for the current payment 
def FindAPR(I,NPER,NPY,P,F,A,t):
    I0 = I
    A0 = RegZPV(I,NPER,NPY,P,F,t)

    while A0 != A:
        # step 1
        A1 = RegZPV(I,NPER,NPY,P,F,t)

        # step 2
        if A1 == A:
            pass
        else:
            I += 0.001

        A2 = RegZPV(I,NPER,NPY,P,F,t)  

        # step 3
        if A2 == A:
            pass
        else:
            I0 += 0.1 * ((A - A1) / (A2 - A1)) / 100

        I = I0
        A0 = round(RegZPV(I,NPER,NPY,P,F,t),2)

    return round(I,6)

# interest rate amortization function
def AmortizeIR(P,I,A,dAMORT):
    i = PeriodRate(I,365)

    for x in range(len(dAMORT)):
        dct = dAMORT[x]

        if x == 0:
            BegP = A
            BegI = 0
        else:
            BegP = EndP
            BegI = EndI
        
        dct['PerRate'] = i
        dct['BegP'] = BegP
        dct['BegI'] = BegI
        dct['CurI'] = round(dct['BegP'] * dct['PerRate'],PRCSN) * dct['Days']
        dct['PMT'] = P
        dct['IPMT'] = round(min(dct['PMT'], dct['BegI'] + dct['CurI']),2)
        dct['PPMT'] = round(min(dct['BegP'], dct['PMT'] - dct['IPMT']),2)
        dct['EndP'] = EndP = round(dct['BegP'] - dct['PPMT'],2)
        dct['EndI'] = EndI = round(dct['BegI'] + dct['CurI'] - dct['IPMT'],PRCSN)

    return dAMORT

# return dAMORT ending balance
def dAMORT_EndP(P,I,A,dAMORT):
    dct = AmortizeIR(P,I,A,dAMORT)
    val = 0
    for x in range(len(dct)):
        val += dct[x]['CurI']
    return val

# find the IR for the current payment 
def FindIR(I,NPER,P,F,A,t,dAMORT):
    E = MyRndDwn(P * NPER - A,2)
    I0 = I
    E1 = dAMORT_EndP(P,I0,A,dAMORT)
    step = 1
    
    # step 2
    I0 += step
    E2 = dAMORT_EndP(P,I0,A,dAMORT)

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
        E2 = dAMORT_EndP(P,I0,A,dAMORT)

        if E1 == E2:
            break
        
        I = I0
        # print(I0,E,E1,E1 == E2)

    return round(I,6), E2

def Main(LoanAmount,MaxTargetAPR,MaxTargetTerm,freq,FundDt,NextPayDt):
    # input variables
    A = LoanAmount
    I = MaxTargetAPR
    NPER = MaxTargetTerm
    freq = freq
    D1 = FundDt
    D2 = NextPayDt
    D1 = du.parser.parse(D1)
    D2 = du.parser.parse(D2)

    # payment schedule inputs
    results = SchDtRange(D1,D2,NPER,freq)
    rng = results['DateRange']

    # find fraction period
    # t = results['IntialPerCnt']
    # t0 = results['FirstIntitalDt']
    u = PaymentFrequency[freq][2]
    FD = FractionDays(D1,D2,u,freq)
    F = FD['F']
    t = FD['t']

    # find the payment for a given APR
    NPY = PaymentFrequency[freq][0]
    P = Payment(A,I,NPER,NPY,F) # starting payment
    P = round(FindPayment(I,NPER,NPY,P,F,A,t),2)

    APR_Test = 0
    while APR_Test == 0:
        # find the APR for a given payment
        APR = FindAPR(I,NPER,NPY,P,F,A,t)

        if APR <= MaxTargetAPR:
            APR_Test = 1
        else:
            # reduce the payment if APR violates target
            P = round(P - 0.01,2)

    # find the IR for a given payment
    dAMORT = BuildAmort(rng)
    FIR = FindIR(I,NPER,P,F,A,t,dAMORT)
    IR = FIR[0]
    IR_FCHRG = FIR[1]

    # add'l outputs
    P1 = dAMORT[0]['Payment Date'] # first payment
    FCHRG  = P * NPER - A # finance charge
    Lst_EndP = dAMORT[len(dAMORT) - 1]['EndP']

    results = {}
    results['P'] = P
    results['APR'] = APR
    results['IR'] = IR
    results['P1'] = P1
    results['FCHRG'] = FCHRG
    results['NPER'] = NPER
    results['t'] = t
    results['F'] = F
    results['IR_FCHRG'] = IR_FCHRG
    results['Lst_EndP'] = Lst_EndP

    return results

def MyTest():
    LoanAmount = 8000
    MaxTargetAPR = 0.0010
    MaxTargetTerm = 12
    freq = 'Monthly'
    FundDt = '2021-01-13'
    NextPayDt = '2021-02-01'

    results = Main(LoanAmount,MaxTargetAPR,MaxTargetTerm,freq,FundDt,NextPayDt)

    print('Loan amount: ${:,.2f}'.format(LoanAmount))
    P = results['P']
    P = 666.67
    print('Payment amount: ${:,.2f}'.format(P))
    FCHRG = results['FCHRG']
    print('Finance charge: ${:,.2f}'.format(FCHRG))
    APR = results['APR']
    print('APR: {:,.4%}'.format(APR))
    IR = results['IR']
    print('IR: {:,.4%}'.format(IR))
    t = results['t']
    print('t: {:,.0f}'.format(t))
    F = results['F']
    print('F: {:,.4%}'.format(F))

    print('')
    D1 = du.parser.parse(FundDt)
    D2 = du.parser.parse(NextPayDt)
    #PrintAmortSch(D1,D2,MaxTargetTerm,freq,P,IR,LoanAmount)

def PrintAmortSch(D1,D2,NPER,freq,P,I,A):
    results = SchDtRange(D1,D2,NPER,freq)
    rng = results['DateRange']
    dAMORT = BuildAmort(rng)
    dAMORT = AmortizeIR(P,I,A,dAMORT)
    print(NPER)
    for x in range((NPER)):
        print(x, dAMORT[x])
    print(I)

def MyTest2():
    LoanAmount = 175
    MaxTargetAPR = 0.352298
    MaxTargetTerm = 52
    freq = 'Weekly'
    FundDt = '2021-06-10'
    NextPayDt = '2021-06-18'

    results = Main(LoanAmount,MaxTargetAPR,MaxTargetTerm,freq,FundDt,NextPayDt)

    print('Loan amount: ${:,.2f}'.format(LoanAmount))
    P = results['P']
    P = 4.01
    print('Payment amount: ${:,.2f}'.format(P))
    FCHRG = results['FCHRG']
    print('Finance charge: ${:,.2f}'.format(FCHRG))
    APR = results['APR']
    print('APR: {:,.4%}'.format(APR))
    IR = results['IR']
    print('IR: {:,.4%}'.format(IR))
    t = results['t']
    print('t: {:,.0f}'.format(t))
    F = results['F']
    print('F: {:,.4%}'.format(F))

    print('')
    D1 = du.parser.parse(FundDt)
    D2 = du.parser.parse(NextPayDt)
    PrintAmortSch(D1,D2,MaxTargetTerm,freq,P,IR,LoanAmount)


MyTest()
#MyTest2()