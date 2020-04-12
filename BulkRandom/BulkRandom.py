import sys
import os
import win32api,win32con

if hasattr(sys, 'frozen'):
    os.environ['PATH'] = sys._MEIPASS + ";" + os.environ['PATH']
from PyQt5 import QtGui
from PyQt5.QtWidgets import *
from people import RandomPeople

departs = []

class Client(QWidget):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        '''读取部门选项'''
        try:
            with open('部门名称.ini', 'r',encoding='utf-8-sig') as file:
                all_depart = file.readlines()
        except FileNotFoundError:
            win32api.MessageBox(0, "找不到部门名称文件", "提醒", win32con.MB_OK)
        for depart in all_depart:
            temp = depart.strip()
            departs.append(temp)

        '''窗口绘制'''
        #self.setFixedSize(390, 500)
        self.setMinimumSize(490,500)
        self.setMaximumSize(800,1000)
        self.setWindowTitle('专家库随机抽取程序')
        self.setWindowIcon(QtGui.QIcon('./picture/年年有鱼.png'))
        self.client_grid = QGridLayout(self)
        self.label1 = QLabel('选择从')
        self.label2 = QLabel('内抽取')
        self.label3 = QLabel('名专家')
        self.spin_box = QSpinBox()
        self.spin_box.setValue(1)
        self.combo_box = QComboBox()
        self.combo_box.addItems(departs)
        self.create_button = QPushButton('抽取')
        self.create_button.clicked.connect(self.create_people_list)
        self.text_browser = QTextBrowser()
        font = QtGui.QFont()
        font.setPointSize(20)
        self.text_browser.setFont(font)
        # 布局
        self.client_grid.addWidget(self.label1, 0, 1, 1, 1)
        self.client_grid.addWidget(self.combo_box, 0, 2)
        self.client_grid.addWidget(self.label2, 0, 3, 1, 1)
        self.client_grid.addWidget(self.spin_box, 0, 4)
        self.client_grid.addWidget(self.label3, 0, 5, 1, 1)
        self.client_grid.addWidget(self.create_button, 0, 6, 1, 1)
        self.client_grid.addWidget(self.text_browser, 1, 1, 1, 6)
    def create_people_list(self):
        people_num = self.spin_box.value()
        department_name = self.combo_box.currentText()
        if department_name == '':
            win32api.MessageBox(0, "所有部门都已抽取", "提醒", win32con.MB_OK)
            return()
      #  self.text_browser.clear()
        self.random_people = RandomPeople(department_name,people_num)
        people_list = self.random_people.create_people()
        self.text_browser.append('%s抽取出的专家：' % department_name)
        for people in people_list:
                        self.text_browser.append(people)
        self.combo_box.clear()
        departs.remove(department_name)
        self.combo_box.addItems(departs)




def main():
    app = QApplication(sys.argv)
    # app.setStyle(QStyleFactory.create('Fusion'))
    client = Client()
    client.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
