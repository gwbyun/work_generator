# -*- coding: utf-8 -*-
"""
Created on Tue Jun 20 16:45:12 2023

@author: bgw60
"""

# -*- coding: utf-8 -*-
"""
Created on Tue Jun 13 13:06:35 2023

@author: bgw60
"""
from PyQt5 import QtCore, QtGui, QtWidgets, uic
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidget, QTableWidgetItem,QLabel
import sys
import pandas as pd
import data_load
import google_mi

# UI 파일을 불러옴
Ui_Dialog, _ = uic.loadUiType("input_v2.ui")

class MainWindow(QMainWindow, Ui_Dialog):
    def __init__(self):
        QMainWindow.__init__(self)
        Ui_Dialog.__init__(self)
        self.setupUi(self)
       

        # 버튼 클릭 시 이벤트 처리
        self.load_pushbutton.clicked.connect(self.load_pushbutton_clicked)
        self.google_pushbutton.clicked.connect(self.google_pushbutton_clicked)
        #self.excel_pushbutton.clicked.connect(self.excel_button)
     
    def load_pushbutton_clicked(self):
        self.year_input =  self.year_input_ui.text()
        self.sheet_input =  self.month_input_ui.text()
        self.name_input =  self.name_input_ui.text()
        print(self.year_input,self.sheet_input,self.name_input)
        
        file_name = '수의사근무표_2023(최신2).xlsx'
        sheet_name = self.sheet_input
        keyword = self.name_input
        self.month, self.date, self.day, self.work = data_load.load_date(file_name,sheet_name,keyword)
        
        
    def google_pushbutton_clicked(self):
        if len(self.date) == len(self.work):
            for i in range(len(self.date)):
                print(self.work[i])
                if self.work != '':
                    google_mi.event_select(int(self.year_input),7,int(self.date[i]),self.work[i])
                
        else:
            raise ValueError("Length of 'date' and 'work' should be the same.")

        

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())