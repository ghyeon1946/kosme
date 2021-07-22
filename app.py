import sys
import pdftotext
import pandas
from PyQt5 import QtWidgets
from PyQt5 import QtCore
from PyQt5.QtWidgets import QFileDialog, QWidget, QApplication, QPushButton


class MyApp(QWidget):
    filename = '출장 및 수당.xlsx'
    def __init__(self):
        super().__init__()
        self.setupUi()
        self.pushButton.clicked.connect(self.Button_click)
        self.currentValue = 0
        self.progressBar.setValue(0)

    def setupUi(self):
        self.setObjectName("Form")
        self.resize(486, 61)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.sizePolicy().hasHeightForWidth())
        self.setSizePolicy(sizePolicy)
        self.setMinimumSize(QtCore.QSize(486, 61))
        self.setMaximumSize(QtCore.QSize(486, 61))
        self.pushButton = QtWidgets.QPushButton(self)
        self.pushButton.setGeometry(QtCore.QRect(10, 10, 101, 41))
        self.pushButton.setObjectName("pushButton")
        self.progressBar = QtWidgets.QProgressBar(self)
        self.progressBar.setGeometry(QtCore.QRect(120, 10, 361, 41))
        self.progressBar.setProperty("value", 24)
        self.progressBar.setOrientation(QtCore.Qt.Horizontal)
        self.progressBar.setTextDirection(QtWidgets.QProgressBar.TopToBottom)
        self.progressBar.setObjectName("progressBar")
        
        self.retranslateUi(self)
        QtCore.QMetaObject.connectSlotsByName(self)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Pdf to Excel (ver 1.0)"))
        self.pushButton.setText(_translate("Form", "변환 파일 선택"))


    def Button_click(self):
        pdflist, _ = QFileDialog.getOpenFileNames(self)
        count = len(pdflist)
        self.currentValue = 0
        self.step = 100 // count

        keys = '업체번호/업체명/진단자번호(위촉)/진단자/수당/출장여비(A + B)/합계'.split('/')
        data = {key:[] for key in keys}

        for pdffile in pdflist:
            values = self.parsePdf(pdffile)
            for key, value in zip(keys, values):
                data[key].append(value)
            self.currentValue += self.step
            if self.currentValue >= 100:
                self.currentValue = 95
            self.progressBar.setValue(self.currentValue)
        
        dataframe = pandas.DataFrame(data)
        dataframe.to_excel(self.filename)
        self.progressBar.setValue(100)


    @staticmethod
    def parsePdf(pdffile):
        with open(pdffile, 'rb') as f:
            pdfreader = pdftotext.PDF(f)
        
        data = []
        lines = pdfreader[0].split('\n')
        # 업체번호
        words = [word.strip() for word in lines[4].split(' ') if word != '']
        data.append(words[1][1:-1])
        data.append(words[2])
        # 진단자 번호, 진단자, 수당 출장여비(A + B)/합계
        words = [word.strip().replace(',', '') for word in lines[14].split(' ') if word != '']
        data += words[:5]

        return data 




if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    ex.show()
    sys.exit(app.exec_())