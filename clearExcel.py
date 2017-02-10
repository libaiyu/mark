# _*_ coding: utf_8  _*_
# Attention: this will damage all xlsx files' data
 
import openpyxl
import os
import re

for k in range(3):
    input("Attention  %d : this will damage all xlsx files' data. Are you sure." % (k + 1))

# find the file that include '学生名单' in filename
ChineseReg = re.compile(r'学生名单')
excelReg = re.compile(r'.xlsx')

wb = openpyxl.Workbook()
sheet = wb.get_active_sheet()

count = 0
for fileName in os.listdir('d:\\_PythonWorks\\excelOperate\\pscj161702'):
    if excelReg.search(fileName):
        if ChineseReg.search(fileName) == None:
            wb.save(fileName)
            count += 1
            print('file %s   is clear!' % (fileName))
print('total %d files is clear!' % (count))
