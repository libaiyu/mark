#! python3

import getdir
from getfile import *

def getfull( dirname, filelist):
    full_list = []
    for file in filelist:
        fullname = dirname + '\\' + file
        full_list.append( fullname)
    return full_list

def test():

    DIRNAME = getdir.getdir()
    FILELIST = getfile( DIRNAME)
    fulllist = getfull( DIRNAME, FILELIST)
    k = 0
    for line in fulllist:
        print(k, line, end='')
        k += 1

if __name__ == '__main__':
    test()
