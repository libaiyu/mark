#! python 3
# _*_ coding: utf_8  _*_

'''
During the course, note the performance for the students.
select the "课程"，select "加减分的项",input"学号","分值"
it will write the "分值"
Record can not be written. prompt is not good enough.   2017-2-11-10:50
Student's number must be 2 digits. When it is smaller than 10, It should be 0X.  2017-2-11-17:23
tidy the value. add display the 总分。display the rank.         2017-2-20
debug the course regular.      2017-3-4.
fix the IndexError.     by evan              2017-3-11
add save error capture. add some Tags.       2017-3-12
'''

import openpyxl
import os
import re

import getdir
from getfull import *
from filesele import *
from getdigits import *       # 2017-3-12

import logging

# logging.basicConfig( level = logging.DEBUG, format = ' %(asctime)s - %(levelname)s - %(message)s' )
logging.basicConfig( level = logging.ERROR, format = ' %(asctime)s - %(levelname)s - %(message)s' )
logging.critical('--------Start of program---------')

# Chinese words '学生名单' in filename
Chinese_reg = re.compile(r'学生名单')
# class
class_reg = re.compile(r'\d{7}')
# course
course_reg = re.compile(r'-([a-z]{3,11})-')     # 2017-3-4 debug.

performance_tag = [
    ['旷课',],
    ['迟到',],
    ['早退',],
    ['提出问题',],
    ['回答问题',],
    ['课堂作业',],
    ['作业1',],['作业2',],['作业3',],['作业4',],['作业5',],['作业6',],['作业7',],
    ]
lab_tag = [
    ['旷课',],
    ['迟到',],
    ['早退',],
    ['操作1',],['操作2',],['操作3',],['操作4',],['操作5',],['操作6',],['操作7',],['操作8',],
    ['数据1',],['数据2',],['数据3',],['数据4',],['数据5',],['数据6',],['数据7',],['数据8',],   
    ['报告1',],['报告2',],['报告3',],['报告4',],
    ]
design_tag = [
    ['旷课',],
    ['迟到',],
    ['早退',],
    ['设计1',],['设计2',],['设计3',],['设计4',],
    ['数据1',],['数据2',],['数据3',],['数据4',],  
    ['报告1',],['报告2',],
    ]
practice_tag = [
    ['旷课',],
    ['迟到',],
    ['早退',],
    ['操作1',],['操作2',],['操作3',],['操作4',],
    ['数据1',],['数据2',],['数据3',],['数据4',],   
    ['报告1',],['报告2',],    
    ]

tagdict = {'performance':performance_tag,
           'lab':lab_tag,
           'design':design_tag,
           'practice':practice_tag}

# Get the directory name.
DIRNAME = getdir.getdir()
# Get the filename list.
FILELIST = os.listdir( DIRNAME)
# Get the filename list include coursetype.
filelist = filesele( FILELIST, course_reg)
# Prompt the number and filename to select.
k = 0
for line in filelist:
    print(k, line)
    k += 1
# File full name list.
fulllist = getfull( DIRNAME, filelist)
st = 'please input a number for select course:'
conum = getdigits( st, 0, k)  # input a digit, it is smaller than k.
coursenum = int( conum)
coursetype = course_reg.search(filelist[coursenum]).group(1)
logging.critical( filelist[coursenum])
print(coursetype)
item = {}
k = 0
for val in tagdict[coursetype]:
    print(str(k) + ' ' + str(val) + ' ')
    item[k+6] = str(val)
    k += 1
st = 'please input a number for select item:'
inum = getdigits( st, 0, k)
itnum = int( inum)
itemnum = 6 + itnum

wb = openpyxl.load_workbook(fulllist[coursenum])

##try:
###    import pdb;pdb.set_trace()
##    wb = openpyxl.load_workbook(fulllist[coursenum])
##except IndexError:
##    input('Please open the workbook, save it and close it.')
##    wb = openpyxl.load_workbook(fulllist[coursenum])

sheet = wb.get_active_sheet()

finish = '0'
while finish.isdigit():   # when finish is digit, do loop.
    studnum = input("\n 输入学号则继续,其他则退出: 205 ")
    for row in range(3,sheet.max_row + 1):
        logging.debug(str(sheet['b'+str(row)].value)[-3:])           #  学号在B列
        if str(sheet['b'+str(row)].value)[-3:] == studnum:                   #  学号在B列
            # Write
            logging.critical(sheet.cell(row = row,column = itemnum).value)   # 写之前，cell的值
            mark = input('\n please input the mark: ')
            sheet.cell(row = row,column = itemnum).value += int(mark)        # 加上要加减的分数
            sheet.cell(row = row,column = 4).value += int(mark)              # 总分也加上该分数
            logging.critical(sheet.cell(row = row,column = itemnum).value)   # 写之后，cell的值
            print(studnum,' 总分是：', sheet.cell(row = row,column = 4).value)
            break
        
    marks = []
    for row in range(3,sheet.max_row + 1):
        marks.append((sheet.cell(row = row,column = 4).value, sheet['b'+str(row)].value, sheet['c'+str(row)].value))
    marks.sort(reverse=True)
    print('前8名为：')
    for k in marks[:8]:
        print(k)
    logging.critical('---------------')
    print( studnum.lower()) 
    finish = studnum   # when studnum is not digit, then finish is not digit, means end.

while True:
    try:    
        wb.save(fulllist[coursenum])
    except PermissionError:
        input('Please close the workbook.')
    else:
        break
    
print('后8名为：')
for k in marks[-8:]:
    print(k)

logging.critical('-------End--------')






