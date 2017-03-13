#! python3

import os
import getdir

def getfull( dirname, filelist):
    full_list = []
    for file in filelist:
        fullname = dirname + '\\' + file
        full_list.append( fullname)
    return full_list

def test():

    DIRNAME = getdir.getdir()
    # Get the filename list.
    FILELIST = os.listdir( DIRNAME)
    fulllist = getfull( DIRNAME, FILELIST)
    k = 0
    for line in fulllist:
        print(k, line, end='')
        k += 1

if __name__ == '__main__':
    test()
