import datetime
import os
import pickle
import shutil
import threading
import tkinter as tk
import openpyxl
from tkinter import Tk, BOTH, filedialog, PhotoImage, Menu, messagebox as mbox
from tkinter.ttk import Frame, Button, Style, Entry, Label
from Settings_window import Worker_set
from parse_time_table import Parser
from support_file import *


def thread(fn):
    def execute(*args, **kwargs):
        threading.Thread(target=fn, args=args, kwargs=kwargs).start()

    return execute


class Example(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent)
        self.parent = parent
        self.parsed = False
        self.wb = openpyxl.load_workbook('some_files/template.xlsx')
        self.initUI()

    def initUI(self):
        self.parent.title("Шото с чем-то")
        self.parent.iconphoto(False, PhotoImage(file='some_files/magic-wand.png'))
        self.style = Style()
        self.style.theme_use("default")
        self.pack(fill=BOTH, expand=1)

        menubar = Menu(self)
        menubar.add_cascade(label="Сотрудники", command=self.open_worker_settings)
        self.parent.config(menu=menubar)

        bias = 100

        file_label = Label(self, text='Расположение файла с расписания занятий')
        file_label.place(x=20 + bias, y=70)

        self.file_path = Entry(self)
        self.file_path.place(x=20 + bias, y=90, width=267, height=30)

        file_picker = Button(self, text="Выбрать файл", command=self.open_file)
        file_picker.place(x=20 + bias, y=130)

        self.parse_button = Button(self, text="Спарсить", command=self.parse_file)
        self.parse_button.place(x=200 + bias, y=130)

        just_button = Button(self, text='Работать!', command=self.make_table)
        just_button.place(x=120 + bias, y=450)

    def open_file(self):
        ftypes = [('Excel files', '*.xlsx')]
        dlg = filedialog.Open(self, filetypes=ftypes)

        fl = dlg.show()
        self.file_path.insert(0, fl)

    def open_worker_settings(self):
        window = Worker_set(self)
        window.geometry("300x256")
        window.grab_set()

    @thread
    def parse_file(self):
        self.parse_button.config(state=tk.DISABLED)
        if os.path.exists('some_files/data.pickle'):
            shutil.copyfile('some_files/data.pickle', 'some_files/old_data.pickle')
        if self.file_path.get() == '':
            mbox.showerror("Ошибка!", 'Выберите сначала файл!')
        else:
            self.data = Parser(self.file_path.get())
            with open('some_files/data.pickle', 'wb') as output:
                pickle.dump(self.data, output)
            self.parsed = True
        self.parse_button.config(state=tk.NORMAL)

    def make_table(self):
        if not self.parsed:
            if os.path.exists('some_files/data.pickle'):
                with open('some_files/data.pickle', 'rb') as f:
                    self.data = pickle.load(f)
            else:
                mbox.showerror('Ошибка!', 'Не найден файл data.pickle')
        wb = openpyxl.load_workbook('some_files/template.xlsx')

        groups = ['Илья Г.:09-933 (1)',
                  'Данил:09-933 (1)',
                  'Ахад:09-063 (1)',
                  'Максим:09-012 (1)',
                  'Илья К.:09-012 (1)', ]
        with open('some_files/workers.pickle', 'rb') as f:
            workers = pickle.load(f)
        # Create area to put lessons ( sry about this ;) )
        magic_var = list(range(6, 77, 14))

        # Getting lessons for all groups
        all_lessons = {}

        for worker in workers:
            all_lessons[':'.join(worker)] = self.data.get_lessons_by_group(worker[1]+worker[2])

        # region Create table for workers
        group_coord = 'C'
        for worker, lessons in all_lessons.items():
            worker=worker.split(':')
            wb[wb.sheetnames[1]][group_coord + "1"] = worker[0]
            wb[wb.sheetnames[1]][group_coord + "2"] = worker[1]+worker[2]
            paint_cells(wb[wb.sheetnames[1]], lessons, bad_color, magic_var, group_coord, True)

            group_coord = chr(ord(group_coord) + 1)
        # endregion

        # region Create table for classrooms
        all_lessons_by_classroom = self.data.get_lessons_by_classrooms(our_classrooms)

        for classroom, lessons in all_lessons_by_classroom.items():
            classroom_index = our_classrooms.index(classroom)
            column = classrooms_columns[classroom_index]
            paint_cells(wb[wb.sheetnames[0]], lessons, good_color, magic_var, column)
        # endregion

        directory = filedialog.askdirectory()
        if directory:
            now = datetime.datetime.now()
            date = f'{now.day}.{now.month}.{now.year}'
            wb.save(f'{directory}/schedule_{date}.xlsx')
        else:
            mbox.showwarning('Тест', 'Файл не сохранён')
        print('DONE!!!')


def main():
    root = Tk()
    root.geometry("512x512")
    Example(root)
    root.mainloop()


if __name__ == '__main__':
    main()
