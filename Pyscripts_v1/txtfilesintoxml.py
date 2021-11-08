
import os
import re
folder = r'\\ac-hq-fs01\users$\daflores\My Documents\XMLData2\\'
os.chdir(folder)
os.getcwd()

#file = folder + r'txt file parse into xml example3' + '.txt'
folders = os.listdir(folder)
for dir in folders:
    os.chdir(folder + dir + '\\')
    files = os.listdir(folder + dir + '\\')
    for file in files:
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
        testing = open('modded' + file,'w+')
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

        #slip = open('list_xml_' + file,'w+')
        #for i in splited:
        #    slip.write(i + '\n')
        #slip.close()


            #print(('\t' * tab) + line)
            #tab += 1
            #print(tab)
            #if re.search('</',line):
            #    tab = tab - 1
            #print(tab)



'''
Sub LoanFundDate()
Dim WeekDay As Variant
Dim NumberofDays As Variant
Dim LastNumberofWeekDay As Variant

WeekDay = "'" & Range("H4").Value & "'"
NumberofDays = "'" & Range("H1").Value & "'"
LastNumberofWeekDay = "'" & Range("H7").Value & "'"


WeekDay = IIf(WeekDay = "'0'", "NULL", WeekDay)
NumberofDays = IIf(NumberofDays = "'0'", "NULL", NumberofDays)
LastNumberofWeekDay = IIf(LastNumberofWeekDay = "'0'", "NULL", LastNumberofWeekDay)


Dim conString As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim myFile, SQLScript As String
Dim LocNo As Integer
Dim fso As New FileSystemObject

LocNo = 601
SQLScript = "EXEC SAIL_Funding_Summary_by_day @weekday = " & WeekDay & ", @numberofdays = " & NumberofDays & ", @LastNumberofWeekday = " & LastNumberofWeekDay


' Create the connection string.
conString = "Provider=SQLOLEDB;" & _
            "Server=AC-PC-097;" & _
            "Initial Catalog=FinanceTeamDB;" & _
            "Trusted_Connection=Yes;"
            

' Create the Connection and Recordset objects.
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset

' Open the connection and execute.
conn.Open conString
Set rs = conn.Execute(SQLScript)

ActiveSheet.Unprotect Password:="Sazza"
ActiveSheet.Cells.Locked = False

Debug.Print SQLScript

Dim conServer As ADODB.Connection
Dim rstResult As ADODB.Recordset
Dim strDatabase As String
Dim strServer As String
Dim strSQL As String

strServer = "AC-PC-097"
strDatabase = "FinanceTeamDB"

Set conServer = New ADODB.Connection
conServer.ConnectionString = "PROVIDER=SQLOLEDB; " _
    & "DATA SOURCE=" & strServer & "; " _
    & "INITIAL CATALOG=" & strDatabase & "; " _
    & "Trusted_Connection=Yes;"
On Error GoTo SQL_ConnectionError
conServer.Open
On Error GoTo 0

Set rstResult = New ADODB.Recordset
strSQL = "set nocount on; "
strSQL = strSQL & SQLScript
rstResult.ActiveConnection = conServer
On Error GoTo SQL_StatementError
rstResult.Open strSQL
On Error GoTo 0


' Get headers

Range("A2:D100000").Select
   
Selection.ClearContents


For i = 0 To rs.Fields.Count - 1
        Sheets("FundingSummary").Cells(2, i + 1).Value = rs.Fields(i).Name
    Next i

    ' Transfer result.
    Sheets("FundingSummary").Range("A2").CopyFromRecordset rs
' Close the recordset
    rs.Close
Sheets("FundingSummary").UsedRange.Columns.AutoFit
Exit Sub

SQL_ConnectionError:
MsgBox "Problems connecting to the server." & Chr(10) & "Aborting..."
Exit Sub

SQL_StatementError:
MsgBox "Connection established. Yet, there is a problem with the SQL syntax." & Chr(10) & "Aborting..."
Exit Sub



End Sub
'''