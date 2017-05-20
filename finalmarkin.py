#! python3
# Get the directory name in this module.
import os
import datetime
import openpyxl
from getdir import getdir, filesele, getfull, getbackup, getfile

global BACKUPFILE

    
def markin( ):    #  mark write in.

    excel_file = getfile() # input( 'Please input the excel file that need write marks.')
    x = 0
    for e in excel_file:
        print( str(x), e)
        x += 1
    filenum = int( input('select file.'))
    # get the multi marks from input. 11189 means the last three digits of students' number is 111, mark is 89.
    st = input('Please input the marks split by ",":')   #  multi marks split by ",".
    backup = 'y'  #  input('Is this need backup?')
    if backup.lower() == 'y':
##        import datetime
##        global BACKUPFILE
        BACKUPFILE = getbackup()
        t = datetime.datetime.now()
        memory_file = open( BACKUPFILE,'a')
        memory_file.write( str( t.year)+'-'+str( t.month)+'-'+str( t.day)+','+str( t.hour)+':'+str( t.minute)+':'+str( t.second)+'\n')
        memory_file.write( excel_file[filenum]+'\n')
        memory_file.write( st+'\n')
        memory_file.close()

    # split each mark.
    mm = st.split(',')
    # get the studnum, mark list.
    itemnum = 22    # Final mark in column 22
    studnum = []
    mark = []
    for e in range(len(mm)):
        if not mm[e].isdigit():    # each mark should be digit. if so, it will write in.
            break                  # till mark is not digit, break.

        if len(mm[e]) == 5:          # student number, mark.
##                print(mm[e],type(mm[e]))
            studnum.append( mm[e][:3])
            marktemp = mm[e][3:]
            mark.append( marktemp)
        elif len(mm[e]) == 3:        # student number change, mark no change.
            studnum.append( mm[e])
            mark.append( marktemp)   #  mark no change.
        else:
            print('数字位数不对。')
    pass
    print( itemnum)
    print( studnum)
    print( mark)

    # open the excel file.
    wb = openpyxl.load_workbook(excel_file[filenum])
    sheet = wb.get_active_sheet()

    # write the marks.
    for e in range(len(studnum)):
        for row in range(3,sheet.max_row + 1):
            if str( sheet[ 'b'+str( row)].value)[-3:] == studnum[e]:           #  学号在B列
                # Write
                sheet.cell(row = row,column = itemnum).value = int(mark[e])        # 加上要加的分数
                break
    pass

    # save the excel file.
    while True:
        try:    
            wb.save(excel_file[filenum])
        except PermissionError:
            input('Please close the workbook.')
        else:
            break
    pass


if __name__ == '__main__':
    markin()


