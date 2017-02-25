
try:
    DIRNAME = open('d:\\directoryname.txt').read()
except IOError:
    print('No such filename.')
    DIRNAME = input('Please input the directory name:')
    f = open('d:\\directoryname.txt', 'w')
    f.write(DIRNAME)
    f.close()
    print('Directory name has been written in d:\\directory.txt.')
    pass
else:
    print('We have got directory name in d:\\directory.txt.')
    pass
