import os
import pickle
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
        self.lb.bind("<BackSpace>", self.del_func)
        Label(self, text='Добавить сотрудника').place(x=150, y=20)
        self.e_add = Entry(self)
        self.e_add.place(x=150, y=45)
        self.e_add.bind('<Return>', self.addItem)
        Label(self, text='Группа').place(x=150, y=80)
        self.e_group = Entry(self)
        self.e_group.place(x=150, y=100)

        Button(self, text='Сохранить', command=self.save_data).place(x=170, y=130)

    def addItem(self, event):
        if event.widget.get().strip() != '':
            self.lb.insert(END, event.widget.get())
            self.workers.append([event.widget.get(), ''])
            self.save()
            self.e_add.delete(0, END)

    def save(self):
        with open('some_files/workers.pickle', 'wb') as f:
            pickle.dump(self.workers, f)

    def onSelect(self, event):
        sel = event.widget.curselection()[0]
        if len(self.workers) > sel:
            worker = self.workers[sel][1]
            self.e_group.delete(0, END)
            self.e_group.insert(0, worker)
        else:
            self.e_group.delete(0, END)

    def save_data(self):
        index = self.lb.curselection()[0]
        group = self.e_group.get()
        if group.strip() != '':
            if len(self.workers) > index:
                self.workers[index][1] = group.strip().lower()
            else:
                self.workers.append([self.lb.get(index), group.strip().lower()])
            self.save()
        else:
            mbox.showerror('Ошибка!', 'Некорректные данные!')

    def del_func(self, event):
        try:
            index = event.widget.curselection()[0]
            self.lb.delete(index)
            del self.workers[index]
            self.save()
        except:
            mbox.showerror('Ошибка!', 'Выберите сотрудника для удаления!')
