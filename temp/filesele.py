#! python3
# _*_ coding: utf_8  _*_

def filesele( filelist, regex):
    'select the file name that according to the regular express.'
    
    fileselect = []
    for file in filelist:
        # slect the files that not fit regex.    
        if regex.search( file):
            fileselect.append( file)
    return fileselect


import re
import os
import getdir

def test():
    # Get the directory name.
    DIRNAME = getdir.getdir()
    # Get the filename list.
    FILELIST = os.listdir( DIRNAME)
    k = 0
    for file in FILELIST:
        print(k, file)
        k += 1
    print('\n')
    course_reg = re.compile(r'-([a-z]{3,11})-')
    filese = filesele( FILELIST, course_reg)
    k = 0
    for file in filese:
        print(k, file)
        k += 1
        
if __name__ == '__main__':
    test()
