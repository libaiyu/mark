# it cannot work .   2017-3-24
# It can work now. just a ',' mistake. it is '.'. 2017-3-25

from time import sleep

from tkinter import Tk
from tkinter.messagebox import showwarning
import win32com.client as win32

warn = lambda app: showwarning(app, 'Exit?')
RANGE = range( 3, 8)

def excel():
    app = 'Excel'
    # xl = win32.gencache.EnsureDispatch('%s,Application' % app)  # ',' I really hate you. 
    xl = win32.gencache.EnsureDispatch('%s.Application' % app)  # static
    # xl = win32.Dispatch('%s.Application' % app)                   # dynamic
    ss = xl.Workbooks.Add()
    sh = ss.ActiveSheet
    xl.Visible = True
    # sleep( 0.1)

    sh.Cells( 1,1).Value = 'Python-to-%s Demo' % app
    # sleep( 0.1)
    for i in RANGE:
        sh.Cells( i,1).Value = 'Line %d' % i   # set the cell value.
        # sleep( 0.1)
    sh.Cells( i+2,1).Value = "Th-th-th-that's all folks!"

    warn( app)
    ss.Close( False)
    xl.Application.Quit()

if __name__=='__main__':
    Tk().withdraw()
    excel()
