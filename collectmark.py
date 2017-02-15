# _*_ coding: utf_8  _*_
# copy marks in many workbooks and put them togather for every student.

import openpyxl
import os
import re
import logging

# logging.basicConfig( level = logging.ERROR, format = ' %(asctime)s - %(levelname)s - %(message)s' )
logging.basicConfig( level = logging.DEBUG, format = ' %(asctime)s - %(levelname)s - %(message)s' )

class collectmarks():
    """    collect marks of the students in several courses.    
    """

    def __init__(self):
        self.classlist = []
        self.titlerow = []
        self.allrow = []
        self.tagindex = {}
        self.newtitle = []
        self.newallrow = []
        self.markdict = {}
        
        self.coursemark = []
        self.classmark = {}
        self.studentmarks = []        

    def findclass(self, srcdir, relist):
        for file in os.listdir(srcdir):
            logging.info(file)
            CLASSREG = re.compile(relist)
            if CLASSREG.search(file):
                # tuple as a element of the list
                self.classlist.append((CLASSREG.search(file).group(1),file))     
        print(self.classlist)

    def findtitle(self, srcdir):
        twoline = 0
        file = self.classlist[0][1]            
        fullname = srcdir + '\\' + file
        logging.info(fullname)         
        wb = openpyxl.load_workbook( fullname)
        sheet = wb.get_active_sheet()
        for row in range(1,sheet.max_row):
            # rowmark: store marks of one row
            # the first element of every row is file name
            rowmark = [file]       
            for col in range(1,sheet.max_column):
                rowmark.append(sheet.cell(row = row, column = col).value)
            # get the title line. it has two lines.
            while(twoline == 1):
                self.titlerow.append(rowmark)
                twoline += 1
            if rowmark[2] == '学    号' and twoline == 0:
                twoline += 1
                self.titlerow.append(rowmark)
        print(self.titlerow)
        
    def copyrows(self, srcdir):
        self.allrow = []   # for store all row marks        
        for t in self.classlist:
            file = t[1]            
            fullname = srcdir + '\\' + file
            logging.info(fullname)
            # copy row marks ,add course name            
            wb = openpyxl.load_workbook( fullname)
            sheet = wb.get_active_sheet()
            for row in range(1,sheet.max_row):
                # rowmark: store marks of one row
                # the first element of every row is file name
                rowmark = [file]       
                for col in range(1,sheet.max_column):
                    rowmark.append(sheet.cell(row = row, column = col).value)
                # add every row marks to form all row marks
                self.allrow.append(rowmark)

    def findindex(self):
        firstline = ['课程名', '序号', '学    号', None, '姓  名', None, '班  级', None, None, None,   '课堂', None,    '课堂',      '课堂',   None,  '课堂',   None, '实践成绩', '实验成绩', '总成绩', '特殊原因', None, '录入状态', '备  注', None]
        secondline = ['课程名', None,    None,      None,    None,  None,   None,   None, None, None, '平时成绩', None,'期中成绩',  '期末成绩', None, '总成绩', None,   None,         None,     None,     None,     None,    None,      None,   None]
        self.tagindex = {}
        self.tagindex[0] = firstline[0]
        for pos in range(1,len(firstline)):
#            print(pos, firstline[pos], secondline[pos])
            if firstline[pos] != None and secondline[pos] == None:
                self.tagindex[pos] = firstline[pos]
            if firstline[pos] != None and secondline[pos] != None:
                self.tagindex[pos] = firstline[pos] + secondline[pos]
        for k,v in self.tagindex.items():
            print(k, v)
            
    def newtitle(self,  ):
        newtitle = []
        index = [0, 2, 4, 6, 10, 12, 13, 15, 17, 18, 19, 20]            
        for n in index:
            newtitle.append(self.tagindex[n])
        logging.info(self.newtitle)
            
    def newrows(self,  ):
        self.newallrow = []
        self.newallrow.append(newtitle)
        for row in self.allrow:
            if re.compile(r'\d{12}').search(str(row[2])):
                newrow = []
                index = [0, 2, 4, 6, 10, 12, 13, 15, 17, 18, 19, 20]            
                for n in index:
                    newrow.append(row[n])
                self.newallrow.append(newrow)
        logging.info(self.newallrow)
            
    def markdic(self,  ):
        newtitled = []
        index = [10, 12, 13, 15, 17, 18, 19, 20]            
        for n in index:
            newtitled.append(self.tagindex[n])
        self.markdict.setdefault((self.tagindex[2],self.tagindex[4]), [(self.tagindex[0],newtitled)])
        for row in self.allrow:
            if re.compile(r'\d{12}').search(str(row[2])):
                newrowd = []
                index = [10, 12, 13, 15, 17, 18, 19, 20]            
                for n in index:
                    newrowd.append(row[n])
                student = (row[2],row[4])
                course = row[0]
                self.markdict.setdefault(student,[])
                self.markdict[student].append([(row[0],newrowd)])
        logging.info(self.markdict)

    def copyall(self, srcdir):
        twoline = 0
        self.classmark = {}    # for store marks of all courses, class as key, marks as the values
        self.allrow = []   # for store all row marks
        for t in self.classlist:
            self.coursemark = []   # for store marks of one course
            classname = t[0]
            file = t[1]            
            fullname = srcdir + '\\' + file
            logging.info(fullname)
            # copy marks ,add course name            
            wb = openpyxl.load_workbook( fullname)
            sheet = wb.get_active_sheet()
            for row in range(1,sheet.max_row):
                # the first element of every row is file name(include course name)
                rowmark = [file]               # for store marks of one row
                for col in range(1,sheet.max_column):
                    rowmark.append(sheet.cell(row = row, column = col).value)
                self.allrow.append(rowmark)      # add one row marks to form all row marks                
                self.coursemark.append(rowmark)      # add one row marks to form a course marks
                # get the title line. it has two lines.
##                print(rowmark[2])                
##                print(twoline)
##                input('any key')
                while(twoline == 1):
                    self.titlerow.append(rowmark)
                    twoline += 1
                if rowmark[2] == '学    号' and twoline == 0:
##                    print('in')
                    twoline += 1
                    self.titlerow.append(rowmark)
##            print(self.titlerow)                
#            logging.info(self.allrow)
#            input('any key')
            self.classmark.setdefault(classname, [])
            self.classmark[classname].append(self.coursemark)

    # copy one class' marks in one sheet. add the course name .     
    def writeclass(self, dstdir):        
        for classname,mark in self.classmark.items():
            wb = openpyxl.Workbook()            
            sheet = wb.get_active_sheet()
            # Write marks 
            row = 0
            logging.info(len(mark))
            logging.info(len(mark[0]))
            logging.info(len(mark[0][0]))
            for k in range(len(mark)):     # k is begin from 0 to max
                for r in range(len(mark[k])):
                    row += 1
                    for col in range(len(mark[k][r])):
                        sheet.cell(row = row, column = col + 2).value = mark[k][r][col]
                        logging.info(str(row) + ' ' + str(r) + ' ' + str(col) + ': ' + str(mark[k][r][col]))
            newfullname = dstdir + '\\cj-total-' + classname + '.xlsx'
            wb.save(newfullname)
            print('Marks of %s have been written.' % (classname))
#            input('any key') 
    
    def writestudent(self, studentnum):
        print(self.titlerow)
        STUDREG = re.compile(studentnum)
        for rowmark in self.allrow:
            if STUDREG.search(str(rowmark[2])):
                self.studentmarks.append(rowmark)                
                print (rowmark)

    def writeall(self, srcdir, dstdir):
        pass

def main():
    SRCDIR = 'd:\\_PythonWorks\\excelOperate\\cj-2016201701'
    DSTDIR = 'd:\\_PythonWorks\\excelOperate\\cjcl-2016201701'
    CLASSRE = '-(\d{7}z?)\.' # CLASSREG = re.compile(r'\d{7}[zZ]?')
    
    collmark = collectmarks()
    CLASSLIST = collmark.findclass(SRCDIR,CLASSRE)
    input('finish findclass')
    collmark.findtitle(SRCDIR)
    input('finish findtitle')
    collmark.copyrows(SRCDIR)
    input('finish copyrows')
    collmark.findindex()
    input('finish findindex')
    collmark.newrows()
    input('finish newrows')
    collmark.markdic()
    input('finish markdic')
    collmark.copyall(SRCDIR)
    input('finish copyall')
    collmark.writeclass(DSTDIR)
    input('finish writeclass')
    
    STUDNUM = input("please input student's number: ")
    collmark.writestudent(STUDNUM)
    input('finish writestudent')

##    collmark.writeall(SRCDIR,DSTDIR)

if __name__ == '__main__':
    logging.critical('--------Start of program---------')
    main()
    logging.critical('--------End---------')

  

##    # copy all marks in one sheet. add the course name .    
##    def allmark(self, srcdir, dstdir):
##        self.studentmarks = []        
##        for t in self.classlist:
##            logging.info(t);logging.info(t[0]);logging.info(t[1])
##            classN = t[0]            
##            file = t[1]            
##            fullname = srcdir + '\\' + file
##            logging.info(fullname)
##            # copy marks ,add course name            
##            wb = openpyxl.load_workbook( fullname)
##            sheet = wb.get_active_sheet()
##            for row in range(1,sheet.max_row):
##                rowmark = []
##                rowmark.append(file)
##                for col in range(1,sheet.max_column):
##                    rowmark.append(sheet.cell(row = row, column = col).value)
##                self.studentmarks.append(rowmark)
##                logging.info(rowmark)
###                input('for debug')
##            logging.info(self.studentmarks)
##        # Write marks    
##        wb = openpyxl.Workbook()
##        sheet = wb.get_active_sheet()
##        for row in range(1,len(self.studentmarks) + 1):
##            for col in range(1,len(self.studentmarks[0]) + 1):
##                sheet.cell(row = row, column = col + 1).value = self.studentmarks[row][col]        
##        newfullname = dstdir + '\\cj-total' + '.xlsx'
##        wb.save(newfullname)   

    


