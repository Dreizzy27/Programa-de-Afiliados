import os
import sys

import openpyxl
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMessageBox

import design


class Main(QtWidgets.QMainWindow, design.Ui_Dialog):
    def __init__(self):
        super(Main, self).__init__()

        self.setupUi(self)
        self.pushButton.clicked.connect(self.load)


    def load(self):
        os.chdir(r"C:\Users\henri\Desktop")
        wb = openpyxl.load_workbook("try.xlsx")
        sheet = wb["Sheet"]


        try:
            a1 = str(self.lineEdit.text().lower())
            a2 = int(self.lineEdit_2.text())
            a3 = float(self.lineEdit_3.text())
            a4 = int(self.lineEdit_4.text())

            b1 = sheet["A2"].value
            b2 = sheet["A3"].value

            v1 = int(sheet["B2"].value + a2)
            v2 = int(sheet["B3"].value + a2)
            v3 = float(sheet["D2"].value + a3)
            v4 = float(sheet["D3"].value + a3)
            v5 = int(sheet["E2"].value + a4)
            v6 = int(sheet["E3"].value + a4)

            sheet["G2"].value = v1
            x1 = int(sheet["G2"].value)

            sheet["G3"].value = v2
            x2 = int(sheet["G3"].value)

            if a1 == b1:
                if x1 >= 15:
                    d = int(x1 - 15)
                    cc = int(sheet["F2"].value + 1)
                    sheet["B2"].value = v1
                    sheet["D2"].value = v3
                    sheet["E2"].value = v5
                    sheet["G2"].value = d
                    sheet["F2"].value = cc
                    wb.save("try.xlsx")
                    if d == 15:
                        cc = int(sheet["F2"].value + 2)
                        sheet["G2"].value = 0
                        sheet["F2"].value = cc
                        wb.save("try.xlsx")

                else:
                    sheet["B2"].value = v1
                    sheet["D2"].value = v3
                    sheet["E2"].value = v5
                    wb.save("try.xlsx")


            elif a1 == b2:
                if x2 >= 15:
                    d = x2 - 15
                    cc = sheet["F3"].value + 1
                    sheet["B3"].value = v2
                    sheet["D3"].value = v4
                    sheet["E3"].value = v6
                    sheet["G3"].value = d
                    sheet["F3"].value = cc
                    wb.save("try.xlsx")
                    if d == 15:
                        cc = sheet["F3"].value + 2
                        sheet["G3"].value = 0
                        sheet["F3"].value = cc
                        wb.save("try.xlsx")

                else:
                    sheet["B3"].value = v2
                    sheet["D3"].value = v4
                    sheet["E3"].value = v6
                    wb.save("try.xlsx")
        except Exception as h:
            self.msg_error(h)

    def msg_error(self, text):
        msg = QMessageBox()
        msg.setWindowTitle("ERRO")
        msg.setIcon(QMessageBox.Critical)
        msg.setText(str(text))
        msg.exec_()






if __name__ == "__main__":
    a = QtWidgets.QApplication(sys.argv)
    app = Main()
    app.show()
    a.exec()
