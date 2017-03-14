#! python3
# Get the directory name in this module.

def getdir():
    ''' '''
    FILE = 'd:\\_PythonWorks\\mark\\temp\\directoryfile.txt'
    try:
        dirname = open(FILE).read()
    except IOError:
        print('No such filename.')
        dirname = input('Please input the directory name:')
        f = open( FILE, 'w')
        f.write(dirname)
        f.close()
        print('Directory name has been written in ', FILE)
        pass
    else:
        print('We have got directory name', dirname, 'in', FILE, '\n')
        pass
    return dirname

def filesele( filelist, regex):
    '''select the file name that according to the regular express.
    '''
    
    fileselect = []
    for file in filelist:
        # slect the files that not fit regex.    
        if regex.search( file):
            fileselect.append( file)
    return fileselect

def getfull( dirname, filelist):
    ''' '''
    full_list = []
    for file in filelist:
        fullname = dirname + '\\' + file
        full_list.append( fullname)
    return full_list

def getdigits( s, min_, max_):
    '''Make the input is the digit or 'q'. s is prompt.
    '''
    
    while True:
        inum = input('\n'+s+'between: '+str( min_)+', '+str( max_)+':')
        while not inum.isdigit():
            if inum == 'q':
                return inum
            else:
                inum = input('\n 输入的必须是数字或q！')
        itnum = int( inum)
        
        if itnum in range( min_, max_):
            return inum
        else:
            print('\n 注意输入数字的范围：')
            pass

def prank( marks, num):
    '''print the students' mark according to given number.
    '''

    marks.sort(reverse=True)
    print('前'+str(num)+'名为：')
    for k in marks[ :num]:
        print(k)
    print('后'+str(num)+'名为：')
    for k in marks[ -num: ]:
        print(k)

def pfrank( file, num):
    '''print the students' mark according to given number.
    '''
    import openpyxl
    wb = openpyxl.load_workbook( file)
    sheet = wb.get_active_sheet()
    marks = []
    for row in range(3,sheet.max_row + 1):
        marks.append((sheet.cell(row = row,column = 4).value, sheet['b'+str(row)].value, sheet['c'+str(row)].value))
    while True:
        try:    
            wb.save( file)
        except PermissionError:
            input('Please close the workbook.')
        else:
            break
    prank( marks, num)




def main():
    # Get the directory name.
    DIRNAME = getdir()
    print( DIRNAME)
    import os
    # Get the filename list.
    FILELIST = os.listdir( DIRNAME)
    k = 0
    for file in FILELIST:
        print(k, file)
        k += 1
    print('\n')
    import re
    course_reg = re.compile(r'-([a-z]{3,11})-')
    filese = filesele( FILELIST, course_reg)
    k = 0
    for file in filese:
        print(k, file)
        k += 1
    print('\n')
    fulllist = getfull( DIRNAME, filese)
    k = 0
    for line in fulllist:
        print(k, line, end='')
        k += 1
    st = 'please input a number '
    getdigits( st, -100, 300)

if __name__ == '__main__':
    main()
