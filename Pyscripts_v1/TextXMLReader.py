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

writer = pd.ExcelWriter(r'\\ac-hq-fs01\users$\daflores\My Documents\\DataDictionary.xlsx', engine='xlsxwriter')

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
                            example.append(line.lstrip().split('>')[1].split('<')[0])
                    df['Folder'] = dir
                    df['File'] = files + '\\' + file
                    df['SentFor'] = sent
                    df = df[['Folder','File','SentFor','Fields','Example']]
                    sheetname = files + '\\' + file
                    df.to_excel(writer,sheet_name = sheetname[0:31])
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
                        example.append(line.lstrip().split('>')[1].split('<')[0])
                df = pd.DataFrame(list(zip(fields,example)),columns = ['Fields','Example'])
                df['Folder'] = dir
                df['File'] = files
                df['SentFor'] = sent
                df = df[['Folder','File','SentFor','Fields','Example']]
                sheetname = files
                df.to_excel(writer,sheet_name = sheetname[0:31])

writer.save()
writer.close()