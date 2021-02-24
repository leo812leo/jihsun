import sys
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.uic import loadUi
def go():
    w.label.setText( "答案：" + str(float(w.lineEdit.text()) + float(w.lineEdit_2.text())))
app = QApplication(sys.argv)
w = loadUi('frist.ui')
w.pushButton.clicked.connect(go)
w.show()
app.exec_()