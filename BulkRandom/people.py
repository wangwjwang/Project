import random
import win32api,win32con
from PyQt5 import QtCore




class RandomPeople(QtCore.QThread):
    text = QtCore.pyqtSignal(list)

    def __init__(self,department_name,people_num=1):
        super().__init__()
        self.department_name= department_name
        self.people_num = people_num

    def create_people(self):
        new_people = []
        s1 = self.department_name
        try:
            with open('%s.txt'%self.department_name, 'r',encoding='utf-8-sig') as file:
                all_people = file.readlines()
        except FileNotFoundError:
            win32api.MessageBox(0, "找不到该部门专家名单,", "提醒", win32con.MB_OK)
        committee_num = 0
        for people in all_people:
            temp = people.strip()
            new_people.append(temp)
            committee_num += 1
        if committee_num < self.people_num:
            self.people_num = committee_num
        people_list = random.sample(new_people, self.people_num)
        self.text.emit(people_list)
        return people_list

    def run(self):
        self.create_people()


if __name__ == '__main__':
    rp = RandomPeople(1)
    rp.create_people()