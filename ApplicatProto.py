import inspect
import sys
import traceback
import os

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import QApplication, QPushButton

from PUI import Ui_Proto
from SettingsXLSX import XLSX
from Config import CreateConfig


class ProcApp(QtWidgets.QMainWindow):
    def __init__(self):
        super(ProcApp, self).__init__()
        self.ui = Ui_Proto()
        self.ui.setupUi(self)
        self.showMaximized()
        self.rul = 0
        self.ui.pushButton.clicked.connect(self.inc)
        self.ui.pushButton_2.clicked.connect(self.inc)
        self.ui.Add_Button.clicked.connect(self.AddProto)
        self.ui.Save_Button.clicked.connect(self.SaveProto)
        self.ui.Save_Button.enteredSignal.connect(self.handel_enter)
        self.ui.Save_Button.leavedSignal.connect(self.handel_leave)
        self.ui.Add_Button.enteredSignal.connect(self.handel_enter)
        self.ui.Add_Button.leavedSignal.connect(self.handel_leave)
        # self.ui.stackedWidget.setCurrentIndex(self.rul) # руль ЁБАНЫХ stackedWidget-оф 
        self.gender = self.ui.gender.currentData()
        print(f'gender  {self.gender}')

    def handel_enter(self):
        if self.sender() == self.ui.Save_Button:
            self.ui.Save_icon.setHidden(False)
        elif self.sender() == self.ui.Add_Button:
            self.ui.Add_icon.setHidden(False)
    
    def handel_leave(self):
        if self.sender() == self.ui.Save_Button:
            self.ui.Save_icon.setHidden(True)
        elif self.sender() == self.ui.Add_Button:
            self.ui.Add_icon.setHidden(True)

    def inc(self):
        #print(inspect.stack()[1][2])
        '''print('======inc======')
        print(self.ui.pushButton)
        print(self.sender())'''
        #print(inspect.getmembers(ProcApp))
        if self.sender() == self.ui.pushButton:
            self.ui.pushButton.clearFocus()
            self.rul += 1
            if self.rul > 1:
                self.rul = 0
            self.ui.stackedWidget.setCurrentIndex(self.rul)
        elif self.sender() == self.ui.pushButton_2: 
            self.ui.pushButton_2.clearFocus()   
            self.rul -= 1
            if self.rul < 0:
                self.rul = 1
            self.ui.stackedWidget.setCurrentIndex(self.rul)
    
    def SaveProto(self):
        self.listdir = os.listdir(os.chdir('C:/Users/master.BASE/AppData/Roaming/Proto'))

        if self.listdir == []:
            CreateConfig.Create(CreateConfig, 'c:/Users/master.BASE/Desktop/Trust/')
        #elif self.listdir[0] == 'config.ini':
        #    CreateConfig.Update(CreateConfig)
        #    CreateConfig.Rewrite(CreateConfig)
        #else:
        #    print('False')

        columns = ['']
        XLSX.createXLSX(XLSX, columns, f'Пр №{XLSX.Get_Value(XLSX, XLSX.Read_Config(XLSX))} от {XLSX.CurentDate(XLSX)[0]}.{XLSX.CurentDate(XLSX)[1]}.xlsx')
        XLSX.Print_Setting(XLSX)
        XLSX.WriteXLSX(XLSX)
        #XLSX.SaveXLSX(XLSX)

    def AddProto(self):
        self.listdir = os.listdir(os.chdir('C:/Users/master.BASE/AppData/Roaming/Proto'))

        if self.listdir == []:
            CreateConfig.Create(CreateConfig, 'c:/Users/master.BASE/Desktop/Trust/')
        self.xls = XLSX()
        if self.xls.temp == []:
            self.xls.row = 2
        elif self.xls.temp != []:
            self.xls.row += 1

        columns = ['']
        XLSX.createXLSX(XLSX, columns, f'Пр №{XLSX.Get_Value(XLSX, XLSX.Read_Config(XLSX))[0]} от {XLSX.CurentDate(XLSX)[0]}.{XLSX.CurentDate(XLSX)[1]}.xlsx')
        XLSX.Print_Setting(XLSX)
        XLSX.WriteXLSX(XLSX, XLSX.Get_Value(XLSX, XLSX.Read_Config(XLSX))[1], self.xls.row)
        XLSX.Update(XLSX)
        XLSX.SaveXLSX(XLSX)

        if self.sender() == self.ui.Add_Button:
            self.ui.Add_Button.clearFocus()
        

Protoapp = QApplication([])
Application = ProcApp()
Application.show()
sys.exit(Protoapp.exec())
