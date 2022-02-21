import os
import datetime


dir = r'\\ac-hq-fs01\accounting\Finance\002 Areas\FinBond\FinBond Monthly Reporting Package\2022\Monthlies'
os.chdir(dir)
list = os.listdir(dir)

for i in list:
    os.rename(i,'2022 01' + i[7:])

