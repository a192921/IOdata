# -*- coding: utf-8 -*-

import sys
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *

class DateTimeEditDemo(QWidget):
    def __init__(self):
        super(DateTimeEditDemo, self).__init__()
        self.initUI()
    def initUI(self):
        self.setWindowTitle("QDateTimeEdit")
        self.resize(300,90)
        
        vlayout = QVBoxLayout()
        self.dateEdit = QDateTimeEdit(QDateTime.currentDateTime(), self)
        # 设置显示的格式
        self.dateEdit.setDisplayFormat("yyyy-MM-dd")
        # 设置最小日期
        self.dateEdit.setMinimumDate(QDate.currentDate().addDays(-3650))
        # 设置最大日期
        self.dateEdit.setMaximumDate(QDate.currentDate().addDays(3650))
        # 单击下拉箭头就会弹出日历控件，不在范围内的日期是无法选择的
        self.dateEdit.setCalendarPopup(True)


        self.btn = QPushButton("获得日期")
        self.btn.clicked.connect(self.onButtonClick)

        vlayout.addWidget(self.dateEdit)
        vlayout.addWidget(self.btn)
        self.setLayout(vlayout)



    def onButtonClick(self):
        date_year = self.dateEdit.date().year()
        date_month = self.dateEdit.date().month()
        date_day = self.dateEdit.date().day()

        print(date_year)
        print(date_month)
        print(date_day)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    demo = DateTimeEditDemo()
    demo.show()
    sys.exit(app.exec_())


