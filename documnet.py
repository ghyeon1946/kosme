import pdftotext
import pandas as pd
import os.path

def file(filename):
    file = open(filename, 'rb')
    fileReader = pdftotext.PDF(file)

    f = open("./script.txt", 'w+')
    f.write(fileReader[0])
    f.close()

    line = 0

    str = []

    f1 = open("./script.txt", 'r+')
    for x in range(12):
        line += 1
        lines = f1.readline()
        str = lines.split()

        if line == 12:
            for i in range(len(str)):
                f2 = open("./success.txt", 'a+')
                f2.write(str[i])
                if i != len(str)-1:
                    f2.write("/")
            f2.write("\n")
            f2.close()
    f1.close()

def toExel(filename):
    df = pd.DataFrame(pd.read_csv(filename, sep='/'))
    df.to_excel('출장 및 수당.xlsx', index=False)

f2 = open("./success.txt", 'a+')
f2.write("진단자번호(위촉)/진단자/수당/출장여비(A + B)/합계" + "\n")
f2.close()

n = int(input())

for i in range(n):
    fileName = input()
    if not os.path.isfile(fileName):
        print("파일이름 오류 다시 입력 :")
        fileName = input()
    file(fileName)

toExel("./success.txt")

f2 = open("./success.txt", 'a+')
f2.seek(0)
f2.truncate()
f2.close()