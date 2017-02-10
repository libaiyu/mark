# _*_ coding: utf_8  _*_
 
import openpyxl
import os
import re
import logging


logging.basicConfig( level = logging.ERROR, format = ' %(asctime)s - %(levelname)s - %(message)s' )
logging.critical('--------Start of program---------')

# find the file that include '学生名单' in filename
ChineseReg = re.compile(r'学生名单')
# find the class
classReg = re.compile(r'\d{7}')
# find the course
courseReg = re.compile(r'-(\w{3,11})-')

students = [
    [ ],
    [ ],
    ['初始分', ],
    ]
# read the nomber and name 
for fileName in os.listdir('d:\\_PythonWorks\\execlOperate\\pscj161702'):
    if ChineseReg.search(fileName): 
        logging.info(ChineseReg.search(fileName).group())        
        className = classReg.search(fileName).group()
        logging.info(className)
        wb = openpyxl.load_workbook( fileName )
        sheet = wb.get_active_sheet()
        for row in range(1,60):
            logging.debug(sheet['b'+str(row)].value)           #  学号在B列
            if sheet['b'+str(row)].value:                      #  学号在B列
                logging.debug(sheet['d'+str(row)].value)                   #  姓名在d列
                students[0].append(str(sheet['B'+str(row)].value))    #  学号在B列 
                students[1].append(sheet['d'+str(row)].value)         #  姓名在D列
                students[2].append(85)         #  初始分设定为85

# write the nomber and name
        for fileName in os.listdir('d:\\_PythonWorks\\execlOperate\\working'):
            classReg2 = re.compile(className)
            if classReg2.search(fileName):
                if ChineseReg.search(fileName) == None:
                    wb = openpyxl.load_workbook(fileName)
                    sheet = wb.get_active_sheet()
                    # Write
                    col = 2       
                    for val in students:
                        logging.info(val)
                        for n in range(len(students[0])):         
                            logging.info(val[n])
                            sheet.cell(row = 2 + n,column = col).value = val[n]
                        col +=1
                    wb.save(fileName)
                    print('one file is written!')

logging.critical('-------End--------')




