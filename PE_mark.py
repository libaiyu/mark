#! python 3
# _*_ coding: utf_8  _*_

'''

'''

import openpyxl


import logging

# logging.basicConfig( level = logging.DEBUG, format = ' %(asctime)s - %(levelname)s - %(message)s' )
logging.basicConfig( level = logging.ERROR, format = ' %(asctime)s - %(levelname)s - %(message)s' )
logging.critical('--------Start of program---------')

PE_items = set(['100米','跳远','铅球','８００米','男子成绩','女子成绩'])
mark_man = {}
mark_woman = {}

tagdict = {}


wb = openpyxl.load_workbook('高考体育成绩对照表.xlsx')

sheet = wb.get_active_sheet()

for col in range(1,sheet.max_column+1):
    for row in range(1,sheet.max_row+1):
        c_v = sheet.cell(row = row,column = col).value
        if c_v in PE_items:
            print(c_v)
        elif c_v:
            print(float(c_v))




logging.critical('-------End--------')






