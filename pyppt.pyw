''' It has an error need to fix.
'''

from time import sleep

from tkinter import Tk
from tkinter.messagebox import showwarning
import win32com.client as win32

warn = lambda app: showwarning(app, 'Exit?')
RANGE = range( 3, 8)

def ppoint():
    app = 'powerpoint'
    # ppoint = win32.gencache.EnsureDispatch('%s,Application' % app)  # ',' I really hate you. 
    # ppoint = win32.gencache.EnsureDispatch('%s.Application' % app)  # static
    ppoint = win32.Dispatch('%s.Application' % app)                   # dynamic
    pres = ppoint.Presentations.Add()
    ppoint.Visible = True

    s1 = pres.Slides.Add( 1, win32.constants.ppLayoutText)
    # sleep( 0.1)

    sla = s1.Shapes[0].TextFrame.TextRange
    sla.Text = 'Python-to-%s Demo' % app
    # sleep( 0.1)
    slb = s1.Shapes[1].TextFrame.TextRange
    for i in RANGE:
        slb.InsertAfter( "Line %d\r\n" % i )  # set the cell value.
        # sleep( 0.1)
    s1b.InsertAfter( "\r\nTh-th-th-that's all folks!")

    warn( app)
    pres.Close()
    ppoint.Quit()

if __name__=='__main__':
    Tk().withdraw()
    ppoint()

##
##
##复制代码 代码如下:
###-*- coding:utf-8 -*-
##from win32com.client import Dispatch
##import time
##def start_office_application(app_name):
### 在这里获取到app后，其它的操作和通过VBA操作办公软件类似
##app = Dispatch(app_name)
##app.Visible = True
##time.sleep(0.5)
##app.Quit()
##if __name__ == '__main__':
##'''''
##通过python启动办公软件的应用进程，
##其中wpp、et、wpp对应的是金山文件、表格和演示
##word、excel、powerpoint对应的是微软的文字、表格和演示
##'''
##lst_app_name = [
##"wps.Application",
##'et.Application',
##'wpp.Application',
##'word.Application',
##'excel.Application',
##'powerpoint.Application'
##]
##for app_name in lst_app_name:
##print "app_name:%s" % app_name
##start_office_application(app_name)
##
##本文开发（python）相关术语:python基础教程 python多线程 web开发工程师 软件开发工程师 软件开发流程 
