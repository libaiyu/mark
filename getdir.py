#! python3
# Get the directory name in this module.

def getdir():
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
        print('We have got directory name', dirname, 'in', FILE)
        pass
    return dirname

def main():
    DIRNAME = getdir()
    print( DIRNAME)

if __name__ == '__main__':
    main()
