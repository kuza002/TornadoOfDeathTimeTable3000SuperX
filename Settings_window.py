import os
import pickle
import re
from tkinter import Toplevel, Listbox, END, Label, messagebox as mbox
from tkinter.ttk import Button, Entry


class Worker_set(Toplevel):

    def __init__(self, parent):
        super().__init__(parent)
        self.initUI()

    def initUI(self):
        self.lb = Listbox(self)
        if os.path.exists('some_files/workers.pickle'):
            with open('some_files/workers.pickle', 'rb') as f:
                self.workers = pickle.load(f)
                for i in self.workers:
                    self.lb.insert(END, i[0])

        else:
            self.workers = []

        self.lb.place(x=0, y=0)
        self.lb.bind("<<ListboxSelect>>", self.onSelect)
        self.entry_list = []
        for i in range(2):
            x_pos, y_pos= 150, 50 * (i + 1)
            self.entry_list.append(Entry(self))
            self.entry_list[-1].place(x=x_pos, y=y_pos)
        Label(self,text='Группа',font=('roboto',10)).place(x=150,y=20)
        # Label(self,text='Подгруппа').place(x=150,y=75)
        Button(self, text='Сохранить', command=self.save_data).place(x=170, y=40 * 3)
        self.e_add = Entry(self)
        self.e_add.place(x=0, y=180)
        Button(self, text="Добавить сотрудника", command=self.addItem).place(x=0, y=200)
        Button(self, text='Удалить сотрудника', command=self.delItem).place(x=0, y=250)

    def addItem(self):
        if self.e_add.get().strip()!='':
            self.lb.insert(END, self.e_add.get())
            self.workers.append([self.e_add.get(),''])


    def onSelect(self,event):
        worker=self.workers[event.widget.curselection()[0]][1]
        for i in range(len(self.entry_list)-1):
            self.entry_list[i].delete(0,END)
            self.entry_list[i].insert(0,worker)


    def save_data(self):
        index = self.lb.curselection()[0]
        group=self.entry_list[0].get()
        if group.strip()!='':
            self.workers[index][1]=group.strip().lower()
            print(self.workers[index])
        else:
            mbox.showerror('Ошибка!', 'Некорректные данные!')

        with open('some_files/workers.pickle', 'wb') as f:
            pickle.dump(self.workers, f)

    def delItem(self):
        try:
            sel = self.lb.curselection()[0]
            self.lb.delete(sel)
            del self.workers[sel]
            with open('some_files/workers.pickle', 'wb') as f:
                pickle.dump(self.workers, f)
        except:
            mbox.showerror('Ошибка!', 'Выберите сотрудника для удаления!')
