class MyClass():
    """    I try to write some modules.
    """
    

    def __init__(self):
        'constuctor'
        pass        

    def getdir(self):
        try:
            dirname = open('d:\\directoryname.txt').read()
        except IOError:
            print('No such filename.')
            dirname = input('Please input the directory name:')
            f = open('d:\\directoryname.txt', 'w')
            f.write(dirname)
            f.close()
            print('Directory name has been written in d:\\directory.txt.')
            pass
        else:
            print('We have got directory name in d:\\directory.txt.')
            pass
        return dirname
    
def main():
    getdir()

if __name__ == '__main__':
    main()
