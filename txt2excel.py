import xlwt


def collect():
    global tl, line, fh, c, wb, sh, position, n
    tl.append(line[5:-1])
    tc = dict()
    line = fh.readline()
    while not (line.startswith('2') or line == ''):
        if line != '\n':
            a = line.split(': ')
            a[1] = a[1][0:4]
            if position.get(a[0]):
                R = position.get(a[0])
            else:
                R = len(position) + 1
                position[a[0]] = R
                sh.write(R, 0, a[0])
            tc[R] = a[1]
        line = fh.readline()
    for k, v in tc.items():
        sh.write(k, c, float(v))
    if c == 144:
        newSheet()
    c += 1
    if line != '':
        collect()


def checkError(tl):
    global lastindex, checklist
    initindex = lastindex + 1
    num = initindex % 12
    for j in tl:
        check=j[-3:-1]
        if check != checklist[num]:
            print('Error occoured at time:\n'+j+' where the time should be '+ checklist[num]+'分.\n')
        lastindex = checklist.index(check)
        num = (lastindex+1)%12   


def newSheet():
    global wb, sh, c, tl, n
    for i in range(len(tl)):
        sh.write(0, i + 1, tl[i][6:8]+':'+tl[i][9:11])
    try:
        sh = wb.add_sheet(tl[0][:-6])
        n = 2
    except:
        sh = wb.add_sheet(tl[0][:-6]+' ('+str(n)+')')
        n += 1
    for k, v in position.items():
        sh.write(v, 0, k)
    checkError(tl)
    tl.clear()
    c = 0


file = 'data_right.txt'
fh = open(file, 'r', encoding='utf-8')
wb = xlwt.Workbook()
tc = dict()
position = dict()
tl = list()
r = 0
line = fh.readline()
checklist = ['00','05','10','15','20','25','30','35','40','45','50','55']
init = 0
while True:
    if line != '\n':
        if line.startswith('2'):
            if init == 1:
                break
            tl.append(line[5:-1])
            init += 1
        else:
            r += 1
            a = line.split(': ')
            a[1] = a[1][0:4]
            position[a[0]] = r
            tc[r] = a[1]
    line = fh.readline()

sh=wb.add_sheet(tl[0][:-6])
for k, v in position.items():
    sh.write(v, 0, k)
for k, v in tc.items():
    sh.write(k, 1, float(v))
c = 2
n = 2
lastindex = checklist.index(tl[0][-3:-1]) - 1
name=tl[0][-6:-4]+tl[0][-3:-1]
while line != '':
    collect()
for i in range(len(tl)):
    sh.write(0, i + 1, tl[i][6:8]+':'+tl[i][9:11])
checkError(tl)
wb.save('平均速度07'+name+'-07'+tl[-1][-6:-4]+tl[-1][-3:-1]+'.xls')
fh.close()
