#! python3
# _*_ coding: utf_8  _*_

def filesele( filelist, regex):
    fileselect = []
    for file in filelist:
        # slect the files that not fit regex.    
        if not regex.search( file):
            fileselect.append( file)
    return fileselect


import re

import getdir
from getfile import *

def test():
    ChineseReg = re.compile(r'学生名单')
    # Get the directory name.
    DIRNAME = getdir.getdir()
    # Get the filename list.
    FILELIST = os.listdir( DIRNAME)
    k = 0
    for line in FILELIST:
        print(k, line)
        k += 1
    filese = filesele( FILELIST, ChineseReg)
    k = 0
    for line in filese:
        print(k, line)
        k += 1

if __name__ == '__main__':
    test()
