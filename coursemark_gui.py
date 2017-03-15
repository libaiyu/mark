#! python3
# _*_ coding: utf_8  _*_
'''list the ahead 8 students' mark.       2017-3-5
try to list the courses. then you can choose the course by click.  2017-3-6
try to arrange the partials.    2017-3-7
try to add a Listbox.    2017-3-15

''' 

import openpyxl
import os
import re
from tkinter import *

from getdir import *


class App(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.course_button1()
        self.course_button2()
        self.create_listbox()
        self.create_entry()
        self.rank_button()
        self.create_text()
        self.quit_button()
        
    def course_button1(self):

        # "列出课程" button .
        self.courseL = Button(self)
        self.courseL["text"] = "列出课程"
        self.courseL.config( command=self.listfileB)
        self.courseL.grid( column=0, row=0, sticky=(W))

    def course_button2(self):

        # "选择课程" button .
        self.courseS = Button(self)
        self.courseS["text"] = "选择课程"
        self.courseS.config( command=self.sele_course)
        self.courseS.grid( column=1, row=0, sticky=(W))
        
    def create_entry(self):

        # label.
        self.ind = Label()
        self.ind['text'] = '先列出课程'
        self.ind.grid(  column=0, row=0, sticky=(W))
        # Entry.
        self.entercourse = Entry()
        # here is the application variable
        self.contents = StringVar()
        # set it to some value
        self.contents.set(FILENAME)
        self.entercourse.config( width=90)
        # tell the entry widget to watch this variable
        self.entercourse["textvariable"] = self.contents
        # when the user hits return
        self.entercourse.bind('<Key-Return>', self.listfileE)        
        self.entercourse.grid( column=0, row=1, sticky=(W))
        
    def rank_button(self):
        
        # "前8名:" button .
        self.rank = Button(self)
        self.rank["text"] = "前8名:"
        self.rank.config( command=self.ahead)
        self.rank.grid( column=2, row=0, sticky=(W))

        
    def create_text(self):

        # text.
        self.ranktext = Text()
        self.ranktext.config( height=10, width=30, wrap='word')
        self.ranktext.grid( column=0, row=6, sticky=(W))
        
    def create_listbox(self):
        
        # try to use listbox.
        self.course = Listbox(self)
        self.course.config( width=80)
        self.course.grid( column=2, row=4, sticky=(W))
        pass

    def quit_button(self):
        
        # "退出"button.
        self.QUIT = Button(self, text="退出", fg="red", command=root.destroy)
        self.QUIT.grid(  column=1, row=8, sticky=(W))
        self.grid( column=0, row=9, sticky=(W))


    def sele_course(self):
        
        global fulllist, NUM
        # every click, NUM increase 1. to select the next course.
        NUM -= 1
        self.contents.set( fulllist[NUM])
        if NUM == 0:
            NUM = len( fulllist)
        pass

    def listfileB(self):
        
        global fulllist, NUM 
        # Get the directory name.
        DIRNAME = getdir()
        # Get the filename list.
        FILELIST = os.listdir( DIRNAME)
        # Get the filename list include coursetype.
        course_reg = re.compile(r'-([a-z]{3,11})-')
        filelist = filesele( FILELIST, course_reg)
        # File full name list.
        fulllist = getfull( DIRNAME, filelist)
        NUM = len( fulllist)
        self.course.delete(0, END)
        for coursen in fulllist:
            self.course.insert(END,  coursen)

    def listfileE(self, event):
        
        global fulllist, NUM 
        # Get the directory name.
        DIRNAME = getdir()
        # Get the filename list.
        FILELIST = os.listdir( DIRNAME)
        # Get the filename list include coursetype.
        course_reg = re.compile(r'-([a-z]{3,11})-')
        filelist = filesele( FILELIST, course_reg)
        # File full name list.
        fulllist = getfull( DIRNAME, filelist)
        NUM = len( fulllist)
        self.ranktext.delete(0.0, END)
        self.course.delete(0, END)
        for coursen in fulllist:
#            self.ranktext.insert(END, coursen + '\n\n')
            self.course.insert(END,  coursen)
        pass

    def ahead(self):
        
        self.rank["activeforeground"] = "red"
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

def prepare():
    
##    import openpyxl
##    import os
##    import re
##
##    # from getdir import *  # 2017-3-12


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
    DIRNAME = getdir()
    # Get the filename list.
    FILELIST = os.listdir( DIRNAME)
    # Get the select list include coursetype.
    filelist = filesele( FILELIST, course_reg)
    # sort the filelist. so the index of the file is nochange.
    filelist.sort()
    # Prompt the number and filename to select.
    k = 0
    for line in filelist:
        print(k, line)
        k += 1
    # File full name list.
    fulllist = getfull( DIRNAME, filelist)
    st = 'select course或q:'
    cour_num = getdigits( st, 0, k)  # input a digit, it is smaller than k.
    if cour_num is 'q':
        pass
    else:
        coursenum = int( cour_num)
        pfrank( fulllist[coursenum], 3)    # print rank.
        coursetype = course_reg.search(filelist[coursenum]).group(1)
        need_check = tagdict[coursetype]
        # need_check = ['提出问题', '课堂作业', '作业1']
        need_check = [ ['作业1',],]
        for each in need_check:
            item_mark( fulllist[coursenum], each[0], 3)  # 分数为0的同学.
        print(coursetype)
        item = {}
        k = 0
        for val in tagdict[coursetype]:
            print(str(k) + ' ' + str(val) + ' ')
            item[k+6] = str(val)
            k += 1

        st = 'select item或q:'
        itnum = getdigits( st, 0, k)
        if itnum is 'q':
            pass
        else:
            itemnum = 6 + int( itnum)

            wb = openpyxl.load_workbook(fulllist[coursenum])
            sheet = wb.get_active_sheet()

            finish = itnum
            while finish.isdigit():   # when finish is digit, do loop.
                st = "\n 输入学号或q: 205, 401:"
                studnum = getdigits( st, 100, 900)
                if not studnum.isdigit():
                    finish = studnum   # when studnum is not digit, then finish is not digit, means end.
                    break
                for row in range(3,sheet.max_row + 1):
                    if str(sheet['b'+str(row)].value)[-3:] == studnum:                   #  学号在B列
                        # Write
                        mark = input('\n please input the mark: ')
                        sheet.cell(row = row,column = itemnum).value += int(mark)        # 加上要加减的分数
                        sheet.cell(row = row,column = 4).value += int(mark)              # 总分也加上该分数
                        print(studnum,' 总分是：', sheet.cell(row = row,column = 4).value)
                        break
            while True:
                try:    
                    wb.save(fulllist[coursenum])
                except PermissionError:
                    input('Please close the workbook.')
                else:
                    break
            pfrank( fulllist[coursenum], 8)    # print rank.

def test():
    global root, FILENAME, NUM  #  root used in QUIT Button( command=root.destroy).
    NUM = 0
    FILENAME = ''
#    prepare()
    root = Tk()
    root.title("课程平时成绩")
    root.geometry('600x380')
    app = App(master=root)
    app.mainloop()

if __name__ == '__main__':
    
    test()
    print('End')


##
##================ RESTART: D:\_PythonWorks\mark\coursemark2.py ================
##hi. contents of entry is now ----> hi.
##Exception in Tkinter callback
##Traceback (most recent call last):
##  File "C:\Python34\lib\tkinter\__init__.py", line 1538, in __call__
##    return self.func(*args)
##TypeError: turn_red() missing 1 required positional argument: 'event'
##End

##button .b -text "Hello, World!" -command exit
##pack .b
##button .b1 -text Hello -underline 0
##button .b2 -text World -underline 0
##bind . <Key-h> {.b1 flash; .b1 invoke}
##bind . <Key-w> {.b2 flash; .b2 invoke}
##pack .b1 .b2

##from tkinter import *
##class App(Frame):
##    def __init__(self, master=None):
##        Frame.__init__(self, master)
##        self.grid()
##
##
### create the application
##myapp = App()
##
###
### here are method calls to the window manager class
###
##myapp.master.title("My Do-Nothing Application")
##myapp.master.maxsize(1000, 400)
##
### start the program
##myapp.mainloop()



##
##import tkinter
##from tkinter.constants import *
##tk = tkinter.Tk()
##frame = tkinter.Frame(tk, relief=RIDGE, borderwidth=2)
##frame.grid(fill=BOTH,expand=1)
##label = tkinter.Label(frame, text="Hello, World")
##label.grid(fill=X, expand=1)
##button = tkinter.Button(frame,text="Exit",command=tk.destroy)
##button.grid(side=BOTTOM)
##tk.mainloop()
##

##from tkinter import *
##
##class Application(Frame):
##    def __init__(self, master=None):
##        Frame.__init__(self, master)
##        self.grid()
##        self.createWidgets()
##
##    def createWidgets(self):
##        self.hi_there = Button(self)
##        self.hi_there["text"] = "Hello World\n(click me)"
##        self.hi_there["command"] = self.say_hi
##        self.hi_there.grid(side="top")
##
##        self.course = OptionMenu(self, text="课程", variable='', value='', fg="red",)
##        self.course.grid()
##        
##        self.QUIT = Button(self, text="QUIT", fg="red",
##                                            command=root.destroy)
##        self.QUIT.grid(side="bottom")
##
##    def say_hi(self):
##        print("hi there, everyone!")
##
##
##def test():
##    global root       #  root used in QUIT Button( command=root.destroy).
##    root = Tk()
##    app = Application(master=root)
####    widgs = app.createWidgets()
##    app.mainloop()
##
##if __name__ == '__main__':
##    test()
##    print('End')

 
##
### try to use GUI to rewrite the program that can record the mark during the course.
##
##from tkinter import *
##
##class Application(Frame):
##    """ A GUI application with three buttons. """
##    
##
##    def __init__( self, master):
##        'Initialize the Frame.'
##
##        super( Application, self).__init__(master)
##        self.grid()
##        self.create_widgets()
##
##    def create_widgets(self):
##        'Create three buttons that do nothing.'
##        
##        # Create first button.
##        self.bttn1 = Listbox( self, text="I do nothing!")
##        self.bttn1.grid()
##
##        # Create second button.
##        self.bttn2 = Button( self)
##        self.bttn2.grid()
##        self.bttn2.configure(text="lazzy button")
##
##        # Create third button.
##        self.bttn3 = Button( self)
##        self.bttn3.grid()
##        self.bttn3['text'] = 'same here'
##
##def test():
##    # test for the function.
##    root = Tk()
##    root.title('平时成绩记录')
##    root.geometry('600x300')
##
##    app = Application(root)
##    root.mainloop()
##
##if __name__ == '__main__':
##    test()
##    print('End')
##
##    


