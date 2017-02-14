# _*_ coding: utf_8  _*_
# copy marks in many workbooks and put them togather for every student.

import openpyxl
import os
import re
import logging

logging.basicConfig( level = logging.ERROR, format = ' %(asctime)s - %(levelname)s - %(message)s' )
# logging.basicConfig( level = logging.DEBUG, format = ' %(asctime)s - %(levelname)s - %(message)s' )

class collectmarks():
    """
    collect marks of the students in several courses.    
    """

    def __init__(self):
        self.classlist = []       

    def classfind(self, srcdir, relist):
        for file in os.listdir(srcdir):
            logging.info(file)
            CLASSREG = re.compile(relist)
            if CLASSREG.search(file): 
                self.classlist.append((CLASSREG.search(file).group(1),file))
##        for t in self.classlist:
##            logging.info(t)
##        logging.info(t[0] + t[1])
        return self.classlist
    
    # copy all marks in one sheet. add the course name .    
    def allmark(self, srcdir, dstdir):

        COLUMNS_MAP = {
                '课堂平时成绩': {'column': 'J', 'index': 3},
                '课堂期末成绩': {'column': 'M', 'index': 4},
                '课堂总成绩': {'column': 'O', 'index': 5},
                '实践成绩': {'column': 'Q', 'index': 6},
                '实验成绩': {'column': 'R', 'index': 7},
                '总成绩': {'column': 'S', 'index': 8},
                }
        studentMarks = [
            ['课程',],
            ['学号', ],
            ['姓名', ],
            ['课堂平时成绩',],
            ['课堂期末成绩',],
            ['课堂总成绩',],
            ['实践成绩',],
            ['实验成绩',],
            ['总成绩',],
            ]        
        
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
                studentMarks[0].append(file)
                studentMarks[1].append(sheet['B'+str(row)].value)            #  学号在B列
                studentMarks[2].append(sheet['D'+str(row)].value)            #  姓名在D列
                for (k, v) in COLUMNS_MAP.items():
                    logging.debug(k + str( sheet[v['column']+str(row)].value ) )  #  课堂平时成绩在J列
                    studentMarks[v['index']].append(sheet[v['column']+str(row)].value)
        # Write marks    
        wb = openpyxl.Workbook()
        sheet = wb.get_active_sheet()
        col = 2       
        for val in studentMarks:
            logging.info(val)
            for n in range(len(studentMarks[0])):         
                logging.info(val[n])
                sheet.cell(row = 2 + n,column = col).value = val[n]
            col +=1
        newfullname = dstdir + '\\cj-total' + '.xlsx'
        wb.save(newfullname)

    # collectmarks for students one by one .
    def tidymark(self):
        pass


def main():
    STUDENT_COUNT = 38
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
    
