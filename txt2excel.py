import xlwt


def collect():
    global tl, line, fh, c, wb, sh, position, n
    tl.append(line[5:])
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
        sh.write(k, c, v)
    if c == 143:
        n += 1
        newSheet(n)
    c += 1
    if line != '':
        collect()

def checkError(tl):
    global lastindex
    checklist = ['00','05','10','15','20','25','30','35','40','45','50','55']
    initindex = lastindex + 1
    num = initindex % 12
    for j in tl:
        check=j[-4:-2]
        if check != checklist[num]:
            print('Error occoured at time:\n'+j)
        lastindex = checklist.index(check)
        num = (lastindex+1)%12

    
def newSheet(n):
    global wb, sh, c, tl
    for i in range(len(tl)):
        sh.write(0, i + 1, tl[i])
    sh = wb.add_sheet(str(n))
    for k, v in position.items():
        sh.write(v, 0, k)
    checkError(tl)
    tl.clear()
    c = 0


file = 'data.txt'
fh = open(file, 'r', encoding='utf-8')
wb = xlwt.Workbook()
sh = wb.add_sheet('1')
tc = dict()
position = dict()
tl = list()
r = 0
line = fh.readline()
init = 0
while True:
    if line != '\n':
        if line.startswith('2'):
            if init == 1:
                break
            tl.append(line[5:])
            init += 1
        else:
            r += 1
            a = line.split(': ')
            a[1] = a[1][0:4]
            position[a[0]] = r
            tc[r] = a[1]
    line = fh.readline()

for k, v in position.items():
    sh.write(v, 0, k)
for k, v in tc.items():
    sh.write(k, 1, v)

c = 2
n = 1
lastindex = 2
while line != '':
    collect()
for i in range(len(tl)):
    sh.write(0, i + 1, tl[i])
checkError(tl)
wb.save('平均速度总表.xls')
fh.close()
