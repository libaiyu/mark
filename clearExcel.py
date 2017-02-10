# _*_ coding: utf_8  _*_
 
import openpyxl
import os
import re

# find the file that include '学生名单' in filename
ChineseReg = re.compile(r'学生名单')
excelReg = re.compile(r'.xlsx')

wb = openpyxl.Workbook()
sheet = wb.get_active_sheet()


for fileName in os.listdir('d:\\_PythonWorks\\execlOperate\\pscj161702'):
    if excelReg.search(fileName):
#        if ChineseReg.search(fileName) == None:
            wb.save(fileName)
            print('file %s is clear!',fileName)
