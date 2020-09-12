import sys
# from PyQt5 import QtGui, QtCore
from PyQt5.uic import loadUiType
from PyQt5 import QtCore, QtGui, QtWidgets
# import PyWhatsapp_edit as Pyit

Ui_MainWindow, QMainWindow = loadUiType('mainGui.ui')

class Main(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(Main, self).__init__()
        self.setupUi(self)
        self.initUI()

    def initUI(self):  
    	self.pushButton_2.clicked.connect(self.datebase_textedit) # open database file
    	self.pushButton_3.clicked.connect(self.message_textedit) # open message file

    def datebase_textedit(self):
        file = QtWidgets.QFileDialog.getOpenFileName(self,'Open database file','*.xlsx','*.xlsx') 
        print(file[0])
        # listName = Pyit.readContacts4(file)
        # for i in listName:
        #     self.textEdit.append(i)

    def message_textedit(self):
        file = QtWidgets.QFileDialog.getOpenFileName(self,'Open message file','*.txt','*.txt') 
        listName = Pyit.read_message(file)
        for i in listName:
            self.plainTextEdit.append(i)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    main = Main()
    main.show()
    sys.exit(app.exec_())