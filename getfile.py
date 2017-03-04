#! python3

import os
import getdir

def getfile(dirname):
    filelist = []
    for file in os.listdir(dirname):
        filelist.append( file)
    return filelist
##
##def getfile(dirname):
##    try:
##        filelist = open('d:\\_PythonWorks\\mark\\filelist.txt').readlines()
##    except IOError:
##        filelist = []
##        for file in os.listdir(dirname):
##            filelist.append( file)
##            # write the filename in the a text file.
##        f = open( 'filelist.txt', 'w')
##        f.writelines( filelist)
##        f.close()
##    else:
##        print('We have got  filelist in d:\\_PythonWorks\\mark\\filelist.txt.')
##    return filelist

def test():

    # Get the directory name.
    DIRNAME = getdir.getdir()
    filelist = getfile( DIRNAME)
    k = 0
    for line in filelist:
        print(k, line, end='')
        k += 1

if __name__ == '__main__':
    test()
