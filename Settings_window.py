import os
import pickle
from tkinter import Toplevel, Listbox, END
from tkinter.ttk import Button, Entry


class Worker_set(Toplevel):
    def __init__(self, parent):
        super().__init__(parent)

        self.lb = Listbox(self)
        if os.path.exists('some_files/workers.pickle'):
            with open('some_files/workers.pickle', 'rb') as f:
                self.workers = pickle.load(f)
                for i in self.workers:
                    self.lb.insert(END, i.split(':')[0])

        self.lb.place(x=0, y=0)
        self.lb.bind("<<ListboxSelect>>", self.onSelect)
        self.e_add = Entry(self)
        self.e_add.place(x=0, y=180)
        Button(self, text="Добавить сотрудника", command=self.addItem).place(x=0, y=200)

    def addItem(self):
        self.lb.insert(END, self.e_add.get())
        self.workers.append(f'{self.e_add.get()}:')


    def onSelect(self):
        pass

