import pandas as pd
import openpyxl
import pyodbc
import pypyodbc

import sqlalchemy
from sqlalchemy import create_engine
import os
import urllib
from bs4 import BeautifulSoup

folder = r'\\ac-hq-fs01\users$\daflores\My Documents\XMLDataResponseRequests\\'
os.chdir(folder)
os.getcwd()



sub_dir = os.listdir(folder)

for dir in sub_dir:
#    print(dir)
    for files in os.listdir(dir):
        if files == 'ReadMe.txt':
            pass
        elif files[-3:] != 'txt':
            for file in os.listdir(folder + dir + r'\\' + files):
                
                with open(folder[:-1] + dir + r'\\' + files + r'\\' + file) as fd:
                    xml = fd.readlines()
                    i = 0
                    sent = ''
                    fields = []
                    example = []

                    for line in xml:
                        #print(line.lstrip().split('>')[0])
                        if len(line.lstrip().split('>')) == 1 :
                            sent = line.lstrip().split('>')[0]
                        else:
                            fields.append(line.lstrip().split('>')[0] + '>')
                            example.append(line.lstrip().split('>')[1][1:])
        else:
            with open(folder[:-1] + dir + r'\\' + files) as fd:
                xml = fd.readlines()
                i = 0
                sent = ''
                fields = []
                example = []
                for line in xml:
                    #print(line.lstrip().split('>')[0])
                    if len(line.lstrip().split('>')) == 1:
                        sent = line.lstrip().split('>')[0]
                    else:
                        fields.append(line.lstrip().split('>')[0] + '>')
                        example.append(line.lstrip().split('>')[1][1:])
                

#for j in soup:
#    n2 = j.split('>')
#    print(n2)
#    for i in n2:
#        print(i + '>')
#    print(type(j))
#    c.append(j)


#c

