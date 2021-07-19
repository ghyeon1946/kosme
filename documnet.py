import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog
import pdftotext
import pandas as pd
import os.path

class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        btn = QPushButton('변환 할 파일 열기', self)
        btn.move(130,20)
        btn.resize(100,40)
        btn.clicked.connect(self.Button_click)
        self.resize(400, 300)
        self.show()

    def Button_click(self):
        file_names = QFileDialog.getOpenFileNames(self)
        f2 = open("./success.txt", 'a+')
        f2.write("업체번호/진단자번호(위촉)/진단자/수당/출장여비(A + B)/합계" + "\n")
        f2.close()

        for file in file_names[0]:
            self.File(file)

    def File(self, filename):
        file = open(filename, 'rb')
        fileReader = pdftotext.PDF(file)

        self.f = open("./script.txt", 'w+')
        self.f.write(fileReader[0])
        self.f.close()

        self.line = 0

        str = []

        self.f1 = open("./script.txt", 'r+')
        for x in range(12):
            self.line += 1
            self.lines = self.f1.readline()
            self.str = self.lines.split()

            if self.line == 4:
                for i in range(len(self.str)):
                    f2 = open("./success.txt", 'a+')
                    if i == 1 and 2:
                        f2.write(self.str[i])
                f2.write("/")

            if self.line == 12:
                for i in range(len(self.str)):
                    #f2 = open("./success.txt", 'a+')
                    f2.write(self.str[i])
                    if i != len(self.str)-1:
                        f2.write("/")
                f2.write("\n")
                f2.close()
        self.f1.close()
        self.toExel("./success.txt")

    def toExel(self,filename):
        df = pd.DataFrame(pd.read_csv(filename, sep='/'))
        df.to_excel('출장 및 수당.xlsx', index=False)
        
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()

    f2 = open("./success.txt", 'a+')
    f2.seek(0)
    f2.truncate()
    f2.close()
    
    sys.exit(app.exec_())

# f2 = open("./success.txt", 'a+')
# f2.write("업체번호/진단자번호(위촉)/진단자/수당/출장여비(A + B)/합계" + "\n")
# f2.close()

# print("입력 파일 개수 : ")
# n = int(input())

# for i in range(n):
#     fileName = input()
#     if not os.path.isfile(fileName):
#         print("파일이름 오류 다시 입력 :")
#         fileName = input()
#     file(fileName)

# toExel("./success.txt")

# f2 = open("./success.txt", 'a+')
# f2.seek(0)
# f2.truncate()
# f2.close()