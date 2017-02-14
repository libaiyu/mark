# _*_ coding: utf_8  _*_
# copy marks in many workbooks and put them togather for every student.

import openpyxl
import os
import re
import logging

logging.basicConfig( level = logging.ERROR, format = ' %(asctime)s - %(levelname)s - %(message)s' )
# logging.basicConfig( level = logging.DEBUG, format = ' %(asctime)s - %(levelname)s - %(message)s' )

class collectmarks():
    """    collect marks of the students in several courses.    
    """

    def __init__(self):
        self.classlist = []       

    def classfind(self, srcdir, relist):
        for file in os.listdir(srcdir):
            logging.info(file)
            CLASSREG = re.compile(relist)
            if CLASSREG.search(file): 
                self.classlist.append((CLASSREG.search(file).group(1),file))    # tuple as a element of the list 
        return self.classlist
    
    # copy all marks in one sheet. add the course name .    
    def allmark(self, srcdir, dstdir):
        studentMarks = []        
        for t in self.classlist:
            logging.info(t);logging.info(t[0]);logging.info(t[1])
            classN = t[0]            
            file = t[1]            
            fullname = srcdir + '\\' + file
            logging.info(fullname)
            # copy marks ,add course name            
            wb = openpyxl.load_workbook( fullname)
            sheet = wb.get_active_sheet()
            for row in range(1,sheet.max_row):
                rowmark = []
                rowmark.append(file)
                for col in range(1,sheet.max_column):
                    rowmark.append(sheet.cell(row = row, column = col).value)
                studentMarks.append(rowmark)
                logging.info(rowmark)
#                input('for debug')
            logging.info(studentMarks)
        # Write marks    
        wb = openpyxl.Workbook()
        sheet = wb.get_active_sheet()
        for row in range(1,len(studentMarks)):
            for col in range(1,len(studentMarks[0])):
                sheet.cell(row = row, column = col + 1).value = studentMarks[row - 1][col - 1]        
        newfullname = dstdir + '\\cj-total2' + '.xlsx'
        wb.save(newfullname)

    # collectmarks for students one by one .
    def tidymark(self):
        pass


def main():
    SRCDIR = 'd:\\_PythonWorks\\excelOperate\\cj-2016201701'
    DSTDIR = 'd:\\_PythonWorks\\excelOperate\\cjcl-2016201701'
    RELIST = '-(\d{7}z?)\.' # CLASSREG = re.compile(r'\d{7}[zZ]?')

    collmark = collectmarks()
    CLASSLIST = collmark.classfind(SRCDIR,RELIST)
    collmark.allmark(SRCDIR,DSTDIR)
    collmark.tidymark()

if __name__ == '__main__':
    logging.critical('--------Start of program---------')
    main()
    logging.critical('--------End---------')
    

##        COLUMNS_MAP = {
##                '课堂平时成绩': {'column': 'J', 'index': 3},
##                '课堂期末成绩': {'column': 'M', 'index': 4},
##                '课堂总成绩': {'column': 'O', 'index': 5},
##                '实践成绩': {'column': 'Q', 'index': 6},
##                '实验成绩': {'column': 'R', 'index': 7},
##                '总成绩': {'column': 'S', 'index': 8},
##                }
