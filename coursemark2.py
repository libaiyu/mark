#! python3

from tkinter import *

class Application(Frame):
    """ A GUI application with three buttons. """
    

    def __init__( self, master):
        'Initialize the Frame.'

        super( Application, self).__init__(master)
        self.grid()
        self.create_widgets()

    def create_widgets(self):
        'Create three buttons that do nothing.'
        
        # Create first button.
        self.bttn1 = Listbox( self, text="I do nothing!")
        self.bttn1.grid()

        # Create second button.
        self.bttn2 = Button( self)
        self.bttn2.grid()
        self.bttn2.configure(text="lazzy button")

        # Create third button.
        self.bttn3 = Button( self)
        self.bttn3.grid()
        self.bttn3['text'] = 'same here'

def test():
    # test for the function.
    root = Tk()
    root.title('平时成绩记录')
    root.geometry('600x300')

    app = Application(root)
    root.mainloop()

if __name__ == '__main__':
    test()
    print('End')

    
