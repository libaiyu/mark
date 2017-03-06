#! python3
# _*_ coding: utf_8  _*_

import openpyxl
import os
import re
from tkinter import *

import getdir
from getfull import *
from filesele import *

class App(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.course_button()
        self.create_entry()
        self.create_listbox()
        self.rank_button()
        self.create_text()
        self.quit_button()

    def course_button(self):

        # "选择课程" button .
        self.course = Button(self)
        self.course["text"] = "选择课程"
        self.course.config( command=self.sele_course)
        self.course.pack()
        
    def create_entry(self):
        
        # Entry.
        self.entercourse = Entry()
        # here is the application variable
        self.contents = StringVar()
        # set it to some value
        self.contents.set(FILENAME)
        self.entercourse.config( width=80)
        # tell the entry widget to watch this variable
        self.entercourse["textvariable"] = self.contents
        # when the user hits return
        self.entercourse.bind('<Key-Return>', self.listfile)        
        self.entercourse.pack()
        
    def rank_button(self):
        
        # "前8名:" button .
        self.rank = Button(self)
        self.rank["text"] = "前8名:"
        self.rank.config( command=self.ahead)
        self.rank.pack()

        
    def create_text(self):

        # text.
        self.ranktext = Text()
        self.ranktext.config( height=10, width=100, wrap='word')
        self.ranktext.pack()

    def quit_button(self):
        
        # "退出"button.
        self.QUIT = Button(self, text="退出", fg="red", command=root.destroy)
        self.QUIT.pack()
        self.pack()

    def create_listbox(self):

##        # try to use listbox.
##        self.course = Listbox(self)
##        self.course.config( width=40)
##        self.course.pack()
        pass

    def sele_course(self):
        
        global fulllist, NUM
        # every click, num increase 1. to select the next course.
        NUM -= 1
        self.contents.set( fulllist[NUM])
        if NUM == 0:
            NUM = len( fulllist)
        pass

    def listfile(self, event):
        
        global fulllist, NUM 
        # Get the directory name.
        DIRNAME = getdir.getdir()
        # Get the filename list.
        FILELIST = os.listdir( DIRNAME)
        # Get the filename list include coursetype.
        course_reg = re.compile(r'-([a-z]{3,11})-')
        filelist = filesele( FILELIST, course_reg)
        # File full name list.
        fulllist = getfull( DIRNAME, filelist)
        NUM = len( fulllist)
        self.ranktext.delete(0.0, END)
        for coursen in fulllist:
            self.ranktext.insert(END, coursen + '\n\n')
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


def test():
    global root, FILENAME   #  root used in QUIT Button( command=root.destroy).
    FILENAME = "d:\_PythonWorks\excelOperate\pscj161702\模拟电子技术-performance-1523701.xlsx"
    root = Tk()
    root.title("课程平时成绩")
    root.geometry('600x360')
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
##        self.pack()
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
##frame.pack(fill=BOTH,expand=1)
##label = tkinter.Label(frame, text="Hello, World")
##label.pack(fill=X, expand=1)
##button = tkinter.Button(frame,text="Exit",command=tk.destroy)
##button.pack(side=BOTTOM)
##tk.mainloop()
##

##from tkinter import *
##
##class Application(Frame):
##    def __init__(self, master=None):
##        Frame.__init__(self, master)
##        self.pack()
##        self.createWidgets()
##
##    def createWidgets(self):
##        self.hi_there = Button(self)
##        self.hi_there["text"] = "Hello World\n(click me)"
##        self.hi_there["command"] = self.say_hi
##        self.hi_there.pack(side="top")
##
##        self.course = OptionMenu(self, text="课程", variable='', value='', fg="red",)
##        self.course.pack()
##        
##        self.QUIT = Button(self, text="QUIT", fg="red",
##                                            command=root.destroy)
##        self.QUIT.pack(side="bottom")
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


