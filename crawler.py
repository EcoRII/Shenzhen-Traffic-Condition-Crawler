import urllib.request
import re
import time


def getHtml(url):
    page = urllib.request.urlopen(url)
    html = page.read()
    return html


def getData():
    for i in range(1, 14):
        html = getHtml("http://sztocc.sztb.gov.cn/roadcongmore.aspx?page=" + str(i))
        H = html.decode("utf8")

        f = open("page" + str(i) + ".txt", "w", encoding='utf8')
        f.write(H)
        f.close()

    file = open("data.txt", "a", encoding='utf8')

    for i in range(1, 14):
        f = open("page" + str(i) + ".txt", encoding='utf8')
        tmp = f.read()
        f.close()
        tmp = tmp.replace(" ", "")
        regt = r'(?<=：).*分'
        regd = r'(?<=align="center">\n).*(?=\n)'
        tre = re.compile(regt)
        dre = re.compile(regd, re.M)
        t = re.findall(tre, tmp)[0]
        data = re.findall(dre, tmp)
        if i == 1:
            file.write(t + "\n\n")
        for j in range(int(len(data) / 4) - 1):
            road = data[4 * j]
            speed = data[4 * j + 2]
            file.write(road + ": " + speed + "\n")
    file.close()
    print(t)
    return 0


k = 1
while k > 0:
    getData()
    time.sleep(300)
