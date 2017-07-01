
st = '0411102,11203,21112,0512301,0521301,222'
stri = st.split(',')
print (stri)

itemnum = []
studnum = []
mark = []
for e in range(len(stri)):
    if len(stri[e]) == 7:
        print(stri[e],type(stri[e]))
        itemnum.append( 6 + int( stri[e][:2]))
        studnum.append( stri[e][2:5])
        mark.append( stri[e][5:])
    elif len(stri[e]) == 5:
        itemnum.append( 'nochange')
        studnum.append( stri[e][:3])
        mark.append( stri[e][3:])
    elif len(stri[e]) == 3:
        itemnum.append( 'nochange')
        studnum.append( stri[e])
        mark.append( 'nochange')
    else:
        print('数字位数不对。')
        pass
print( itemnum, studnum, mark)
        
##['0411102', '11203', '21112', '0512301', '0521301', '222']
##0411102 <class 'str'>
##0512301 <class 'str'>
##0521301 <class 'str'>
