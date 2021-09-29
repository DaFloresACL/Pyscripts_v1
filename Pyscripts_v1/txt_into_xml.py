import os
import re
folder = r'\\ac-hq-fs01\users$\daflores\My Documents\\'
os.chdir(folder)
os.getcwd()

file = folder + r'txt file parse into xml example3' + '.txt'
with open(file) as f:
    lines = f.readlines()
newlines = ''
for i in lines:
    newlines = newlines + i

while re.search('> <',newlines):
    newlines = newlines.replace('> <','><')

splited = newlines.split('>')
for index, item in enumerate(splited):
    splited[index] = splited[index] + '>'

splited.pop()
splited.pop(0)

splited
lastline = ''
start = 0
tab = 0
undent = 0
sameline = 1
tabdict = {}
testing = open('tst_xmlwrite_stage3.txt','w+')
for line in splited:
    if line[0] == '<':
        linekey = line
        linekey = linekey.replace('</','')
        linekey = linekey.replace('<','')
        linekey = linekey.replace('>','')
        addedkey = linekey.split(' ')[0]
        if addedkey in tabdict:
            tab = tabdict[addedkey] + 1
            undent = 1
        else:
            tabdict[addedkey] = tab
    
    if re.search('</',lastline) and  re.search('</',line):
        tab = tab - 1
        #testing.write('\n')
        sameline = 0
    if start == 0:
        testing.write(line)
        tab += 1
        start = start +1
    elif re.search('</',line) and sameline != 0:
        testing.write(line)
        tab = tab - 1
        sameline = 0
    else:
        testing.write( '\n' + ('\t' * tab) + line )
        tab += 1
        sameline = 1
    lastline = line
    if undent == 1:
        tab = tab - 1
        undent = 0
tabdict
testing.close()

slip = open('list_xml.txt','w+')
for i in splited:
    slip.write(i + '\n')
slip.close()


    #print(('\t' * tab) + line)
    #tab += 1
    #print(tab)
    #if re.search('</',line):
    #    tab = tab - 1
    #print(tab)


