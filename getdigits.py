
def getdigits( s, min_, max_):
    '''Make the input is the digit. s is prompt.
    '''
    while True:
        inum = input( '\n'+s)
        while not inum.isdigit():
            inum = input('\n 输入的必须是数字！')
        itnum = int( inum)
        
        if itnum in range( min_, max_):
            return itnum
        else:
            print('\n 注意输入数字的范围：')
            pass

def test():
    st = 'please input a number for select course:'
    getdigits( st, -100, 300)

if __name__ == '__main__':
    test()
