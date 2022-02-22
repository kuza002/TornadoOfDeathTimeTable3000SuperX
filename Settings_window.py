import os
import pickle
import re
from tkinter import Toplevel, Listbox, END, messagebox as mbox
from tkinter.ttk import Button, Entry


class Worker_set(Toplevel):
    index=0
    def __init__(self, parent):
        super().__init__(parent)
        self.initUI()


    def initUI(self):
        self.lb = Listbox(self)
        if os.path.exists('some_files/workers.pickle'):
            with open('some_files/workers.pickle', 'rb') as f:
                self.workers = pickle.load(f)
                for i in self.workers:
                    self.lb.insert(self.index, i[0])
                    self.index += 1
        else:
            self.workers = []

        self.lb.place(x=0, y=0)
        self.lb.bind("<<ListboxSelect>>", self.onSelect)
        self.entry_list = []
        for i in range(2):
            self.entry_list.append(Entry(self))
            self.entry_list[-1].place(x=150, y=30 * (i + 1))
        Button(self, text='Сохранить', command=self.save_data).place(x=150, y=30 * 3)
        self.e_add = Entry(self)
        self.e_add.place(x=0, y=180)
        Button(self, text="Добавить сотрудника", command=self.addItem).place(x=0, y=200)

    def addItem(self):
        self.lb.insert(self.index, self.e_add.get())
        self.index+=1
        self.workers.append([self.e_add.get(),'',''])


    def onSelect(self,event):
        worker=self.workers[event.widget.curselection()[0]][1:]
        print(event.widget.curselection()[0])
        for i in range(len(self.entry_list)):
            self.entry_list[i].delete(0,END)
            self.entry_list[i].insert(0,worker[i])


    def save_data(self):
        pattern_group=r'^09-[0-9][0-6][1-5]$'
        pattern_subgroup = r'^\([1-2]\)$'
        group=self.entry_list[0].get()
        subgroup=self.entry_list[1].get()
        if re.search(pattern_group,group.strip()):
            self.workers[self.lb.curselection()[0]][1]=f'{group} '
            print(self.workers[self.lb.curselection()[0]])
        else:
            mbox.showerror('Ошибка!', 'Некорректные данные!')
        if re.search(pattern_subgroup,subgroup.strip()):
            self.workers[self.lb.curselection()[0]][2] = subgroup
            print(self.workers[self.lb.curselection()[0]])
        else:
            mbox.showwarning('Предупреждение','Не указана подгруппа!\n Автоматический будет установлена первая')
            self.workers[self.lb.curselection()[0]][2] = '(1)'
        with open('some_files/workers.pickle', 'wb') as f:
            pickle.dump(self.workers, f)


