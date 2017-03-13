
def getdigits( s, min_, max_):
    '''Make the input is the digit. s is prompt.
    '''
    
    while True:
        print( '\n'+s+'between: '+str( min_)+', '+str( max_))
        inum = input()
        while not inum.isdigit():
            inum = input('\n 输入的必须是数字！')
        itnum = int( inum)
        
        if itnum in range( min_, max_):
            return itnum
        else:
            print('\n 注意输入数字的范围：')
            pass

def phead( marks, num):
    '''print the students' mark according to given number.
    '''

    marks.sort(reverse=True)
    print('前8名为：')
    for k in marks[:num]:
        print(k)

def prear( marks, num):
    '''print the students' mark according to given number.
    '''

    marks.sort()
    print('后8名为：')
    for k in marks[:num]:
        print(k)

def test():
    st = 'please input a number '
    getdigits( st, -100, 300)

if __name__ == '__main__':
    test()
