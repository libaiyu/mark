# _*_ coding: utf_8  _*_
import os
import logging
# find the file by name in the given folder.

# logging.basicConfig( level = logging.DEBUG, format = ' %(asctime)s - %(levelname)s - %(message)s' )
logging.basicConfig( level = logging.ERROR, format = ' %(asctime)s - %(levelname)s - %(message)s' )

name2find = input("please input the file name you try to find:")
dirname = input("please input the dir:")
# dirname = 'd:\\_PythonWorks\\workDir\\2电路分析方法-test'
for folderName,subFolders,filenames in os.walk(dirname):
    logging.info(' the current folder is  ' + folderName )

    for subFolder in subFolders:
        logging.info(' subFolder of ' + folderName + ':  ' + subFolder )

    for filename in filenames:
        logging.info(' file inside ' + folderName + ':  ' + filename )
        if filename == name2find:
            print(' file inside ' + folderName + ':  ' + filename )
    
print ('---Done---')
