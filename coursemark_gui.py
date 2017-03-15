#! python3
# _*_ coding: utf_8  _*_
'''list the ahead 8 students' mark.       2017-3-5
try to list the courses. then you can choose the course by click.  2017-3-6
try to arrange the partials.    2017-3-7
try to add a Listbox.  it can work. but not strong. minus can not work. 2017-3-15

''' 

import openpyxl
import os
import re
from tkinter import *

from getdir import *

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

class App(Frame):
    '''GUI application that manage the course.
    '''

    def __init__( self, master):
        'Initialize the Frame.'

        super( App, self).__init__(master)
        self.grid()
        self.create_widgets()

    def create_widgets(self):
        'Create all kinds of widgets.'
        
        # Create"列出课程" button .
        Button( self,
                text = "列出课程",
                command=self.listfile
                ).grid( column=0, row=0, sticky=(W))

        # Create listbox.
        self.course =  Listbox(self,
                               width=85, height = 6
                               )
        self.course.grid( column=0, row=1, sticky=(W))

        # Create"选择课程" button .
        Button( self,
                text = "选择课程",
                command=self.sele_course
                ).grid( column=0, row=2, sticky=(W))
        
        # label.
        Label( self,
               text = '所选课程'
               ).grid(  column=0, row=3, sticky=(W))

        # Entry.
        self.contents = StringVar()
        # set it to some value
        self.contents.set('test.')
        self.course_ent = Entry( self,
                                 width=85,
                                 textvariable = self.contents
                                 )
        self.course_ent.grid( column=0, row=4, sticky=(W))
        # tell the entry widget to watch this variable
        self.course_ent["textvariable"] = self.contents
        # when the user hits return
        self.course_ent.bind('<Key-Return>', self.markin)
           
        # "前8名:" button .
        Button(self,
               text = "前8名:",
               command = self.ahead
               ).grid( column=0, row=5, sticky=(W))

        # text.
        self.ranktext = Text(self,
                             height=10, width=30, wrap='word'
                             )
        self.ranktext.grid( column=0, row=6, sticky=(W))

        # "加减分项目:" button .
        Button(self,
               text = "加减分项目:",
               command = self.list_item
               ).grid( column=0, row=7, sticky=(W))

        # "退出"button.
        Button(self,
               text="退出",
               fg="red",
               command=root.destroy
               ).grid(  column=0, row=11, sticky=(W))

    def sele_course(self):
        
        global fulllist, NUM
        # every click, NUM increase 1. to select the next course.
        NUM -= 1
        self.contents.set( fulllist[NUM])
        if NUM == 0:
            NUM = len( fulllist)
        pass

    def listfile(self):
        
        global fulllist, NUM 
        # Get the directory name.
        DIRNAME = getdir()
        # Get the filename list.
        FILELIST = os.listdir( DIRNAME)
        # Get the filename list include coursetype.
        filelist = filesele( FILELIST, course_reg)
        # sort the filelist. so the index of the file is nochange.
        filelist.sort()       
        # File full name list.
        fulllist = getfull( DIRNAME, filelist)
        NUM = len( fulllist)
        self.course.delete( 0, END)
        for coursen in fulllist:
            self.course.insert( END, coursen)

    def list_item(self):
        # list the item according to the coursetpye.
        self.course.delete( 0, END) # self.ranktext.delete(0.0, END)
        coursetype = course_reg.search( fulllist[NUM]).group(1)
        k = 0
        for val in tagdict[coursetype]:
            self.course.insert( END, str(k)+','+str(val)+'\n')
            k += 1
        self.contents.set('请选择项目代号，然后输入学号及分数')
        pass

    def markin( self, event):    #    event.??
        # modify the mark.
        
        global fulllist, NUM

        st = self.course_ent.get()
        print( st)
        if st.isdigit():
            itemnum = 6 + int( st[:2])

            wb = openpyxl.load_workbook(fulllist[NUM])
            sheet = wb.get_active_sheet()

            studnum = st[2:5]
            for row in range(3,sheet.max_row + 1):
                if str(sheet['b'+str(row)].value)[-3:] == studnum:                   #  学号在B列
                    # Write
                    mark = st[5:]
                    dp = studnum +':'+ str( sheet.cell(row = row,column = itemnum).value)
                    self.ranktext.insert( END, dp+' ')
                    sheet.cell(row = row,column = itemnum).value += int(mark)        # 加上要加减的分数
                    dp = str( sheet.cell( row = row,column = itemnum).value)
                    self.ranktext.insert( END, dp+' ')
                    sheet.cell(row = row,column = 4).value += int(mark)              # 总分也加上该分数
                    break
            while True:
                try:    
                    wb.save(fulllist[NUM])
                except PermissionError:
                    input('Please close the workbook.')
                else:
                    break
            pass
    
    def ahead(self):
        
        global NUM, fulllist
        # Read the marks.
        if NUM == 3:
            NUM = 0
        wb = openpyxl.load_workbook( fulllist[NUM])
        sheet = wb.get_active_sheet()
        marks = []
        for row in range(3,sheet.max_row + 1):
            marks.append((sheet.cell(row = row,column = 4).value, sheet['b'+str(row)].value, sheet['c'+str(row)].value))
        wb.save( fulllist[NUM])
        
        self.ranktext.delete(1.0, END)
        self.ranktext.insert(END, '前8名为：\n')
        # rank the marks.
        marks.sort(reverse=True)
        for k in marks[:8]:
            # insert the ahead marks to text.
            self.ranktext.insert(END, str(k)+'\n')
        if NUM == 0:
            NUM = 3
        pass

def test():
    global root, FILENAME, NUM  #  root used in QUIT Button( command=root.destroy).
    NUM = 0
    FILENAME = ''
    root = Tk()
    root.title("课程平时成绩")
    root.geometry('580x380')
    app = App(master=root)
    app.mainloop()

if __name__ == '__main__':
    
    test()
    print('End')


