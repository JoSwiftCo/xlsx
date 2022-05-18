#!/home/pi/software/bin/python3
import cgi, cgitb
cgitb.enable()
import xlsxwriter
from xlsxwriter import Workbook
import re
from datetime import date
from datetime import datetime
import copy
from functools import reduce
import random
import math 
import string
import time  

form = cgi.FieldStorage( )

quiz           = form.getvalue('quiz')
lab            = form.getvalue('lab')
assignment     = form.getvalue('assignment')
test           = form.getvalue('test')
quizzesval     = form.getvalue('quizval')
labsval        = form.getvalue('labval')
assignmentsval = form.getvalue('assignmentval')
testsval       = form.getvalue('testval')
Courses        = form.getvalue('course')
go             = form.getvalue('go')
color          = form.getvalue('bgcolor')
pattern        = r'\b[A-Z]{3}\d{3}\b'
today          = date.today()
tDate          = today.strftime("%Y.%m.%d")
tDate2         = today.strftime("%B %d, %Y")
now            = datetime.now()
nTime          = now.strftime("%H:%M %p")
excel          = "." + tDate + ".xlsx"

quizN    = int(quiz)
quizval  = int(quizzesval)
labN     = int(lab)
labval   = int(labsval)
asN      = int(assignment)
asval    = int(assignmentsval)
testN    = int(test)
testval  = int(testsval)

total = quizN*quizval + labN*labval + asN*asval + testN*testval
print("<html>\n") 
#print("<body bgcolor='%s'><pre>" %color)

quizRow    = ['Q1']
labRow     = ['L1']
asRow      = ['A1']
testRow    = ['T1']

for i in range(quizN-1):
   result1 = re.sub(r'[0-9]', lambda x: f"{str(int(x.group())+1).zfill(len(x.group()))}",quizRow[i])
   quizRow.append(str(result1))
for i in range(labN-1):
   result2 = re.sub(r'[0-9]', lambda x: f"{str(int(x.group())+1).zfill(len(x.group()))}",labRow[i])
   labRow.append(str(result2))
for i in range(asN-1):
   result3 = re.sub(r'[0-9]', lambda x: f"{str(int(x.group())+1).zfill(len(x.group()))}",asRow[i])
   asRow.append(str(result3))
for i in range(testN-1):
   result = re.sub(r'[0-9]', lambda x: f"{str(int(x.group())+1).zfill(len(x.group()))}",testRow[i])
   testRow.append(str(result))

if re.findall(pattern, Courses):
   if(total == 100):
      workbook = xlsxwriter.Workbook('A2.xlsx') 
      print("Spreadsheet", Courses.lower() + excel, "created on", tDate2, "at", nTime)
      print("\nSpreadsheet contains: ")
      print(" ")

      print(quizN, "quizzes @", quizval ,"% each")
      print(labN, "labs @", labval , "% each")
      print(asN, "assignments @", asval , "% each")
      print(testN, "tests @", testval , "% each")
      print("TOTAL", total, "%")
   else:
      print("ERROR, Please go back to enter new values")
else:
   print("The course code has to be 3 characters being letters and the last 3 numbers\nExample: PRG550 or ETY155")


# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('A2.xlsx')
worksheet = workbook.add_worksheet( )
cell_format = workbook.add_format()
cell_format.set_align('left') 
cell_format.set_pattern(1)
cell_format.set_bg_color(color)

# Insert an image.
worksheet.insert_image('A7', 'python.png')
worksheet.set_column('A1:A1', 15)
worksheet.set_column('B1:B1', 10)
worksheet.set_column('C1:C1', 25)
worksheet.set_column('D1:D1', 15)

alpha = list(string.ascii_uppercase)
start = alpha[4]
for i in range(quizN):
    worksheet.write(alpha[4+i] + '1', quizRow[i], cell_format)
for i in range(labN):
    worksheet.write(alpha[4+i+quizN] + '1', labRow[i], cell_format)
for i in range(asN):
    worksheet.write(alpha[4+i+quizN+labN] + '1', asRow[i], cell_format)
for i in range(testN):
    worksheet.write(alpha[4+i+quizN+labN+asN] + '1', testRow[i], cell_format)
    score = alpha[4+i+quizN+labN+asN+testN] + '1'
    grade = alpha[4+1+quizN+labN+asN+testN] + '1'

worksheet.write(score, "Score", cell_format)
worksheet.write(grade, "Grade", cell_format)

def addRowToXlsFile(workbook, stuInfo, row = 0):
    for colNum, value in enumerate(stuInfo):
        workbook.write(row, colNum, value)

xlsxwriter.worksheet.Worksheet.addRow = addRowToXlsFile

worksheet.addRow(stuInfo = ["Student #, Last, First, Login"])
rowC = 0
with open("data.dat", 'r') as file:
    lines = file.readlines()
    for line in lines:
        worksheet.addRow(stuInfo=line.split(';'), row = rowC)
#        stuInfo[3] = stuInfo[3].rstrip("\n")
        rowC  = rowC + 1
#        worksheet.write(rowC,0, stuInfo[2])
#        worksheet.write(rowC,1, stuInfo[0])
#        worksheet.write(rowC,2, stuInfo[1])
#        worksheet.write(rowC,3, stuInfo[3])
worksheet.write('A1','Student #', cell_format)
worksheet.write('B1','Last', cell_format)
worksheet.write('C1','First', cell_format)
worksheet.write('D1','Login', cell_format)
workbook.close( )




