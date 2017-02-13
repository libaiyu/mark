# _*_ coding: utf_8  _*_
# read marks of many workbooks then write the marks togather for every student.

import openpyxl
import os
import re
import logging

# logging.basicConfig( level = logging.ERROR, format = ' %(asctime)s - %(levelname)s - %(message)s' )
logging.basicConfig( level = logging.DEBUG, format = ' %(asctime)s - %(levelname)s - %(message)s' )


class collectmarks():
    """
    collect marks of the students in several courses.    
    """

    def __init__(self):
        self.classlist = []
#        self.classdict = {}

##    def set_regular(self, list):
##        self.regexList.append(re.compile(s))        

    def classfind(self, dirname, relist):
        for file in os.listdir(dirname):
            logging.info(file)
            CLASSREG = re.compile(relist)
            if CLASSREG.search(file): 
                self.classlist.append((CLASSREG.search(file).group(1),file))
            #    self.classdict.setdefault(CLASSREG.search(file).group(1), file)
                
            #    list(self.classdict[CLASSREG.search(file).group()].value).append(file)
        logging.info(self.classlist)
        
        # logging.info(self.classdict)
        input('any key to continue.')
        return self.classlist
        
    def colclass(self, dirname, classname):
        pass

    def colstud(self, classname, studnumb):
        pass


def main():
    STUDENT_COUNT = 38
    DIRNAME = 'd:\\_PythonWorks\\excelOperate\\cj-2016201701'
    RELIST = '-(\d{7}z?)\.' # CLASSREG = re.compile(r'\d{7}[zZ]?')
    COLUMNS_MAP = {
            '课堂平时成绩': {'column': 'J', 'index': 3},
            '课堂期末成绩': {'column': 'M', 'index': 4},
            '课堂总成绩': {'column': 'O', 'index': 5},
            '实践成绩': {'column': 'Q', 'index': 6},
            '实验成绩': {'column': 'R', 'index': 7},
            '总成绩': {'column': 'S', 'index': 8},
            }
    collmark = collectmarks()
    collmark.classfind(DIRNAME,RELIST)
##    for classname in collmark.classfind(DIRNAME,RELIST):
##        collmark.collclass(classname)
#    collmark.collstud(studnumb)

if __name__ == '__main__':
    logging.critical('--------Start of program---------')
    main()
    logging.critical('--------End---------')
