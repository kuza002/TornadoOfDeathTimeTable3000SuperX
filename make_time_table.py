import os
import pickle
import shutil
import time
import tkinter.filedialog

import numpy as np
import openpyxl
from openpyxl.styles import PatternFill

from parse_time_table import Parser, write_in_file

days_of_week = ['пн', 'вт', 'ср', 'чт', 'пт', 'сб']
times = ['8.30-10.00', '10.10-11.40', '11.50-13.20', '14.00-15.30', '15.40-17.10', '17.50-19.20']
our_classrooms = ['802', '804', '808', '809', '810', '811', '910', '1009', '1111', '1112', '1206']
classrooms_columns = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
good_color = 'ccffcc'
bad_color = 'ffadd6'


def paint_cells(sheet, lessons, color, magic_var, column):
    for lesson in lessons:

        dw_index = days_of_week.index(lesson.day_of_week)

        time_index = times.index(lesson.duration)

        lesson_row = magic_var[dw_index] + time_index * 2

        if 'н/н' in lesson.cell.value:
            sheet[column + str(lesson_row)].fill = PatternFill("solid", start_color=color)
        elif 'ч/н' in lesson.cell.value:
            sheet[column + str(lesson_row + 1)].fill = PatternFill("solid", start_color=color)
        else:
            sheet[column + str(lesson_row)].fill = PatternFill("solid", start_color=color)
            sheet[column + str(lesson_row + 1)].fill = PatternFill("solid", start_color=color)


start_time = time.time()
# Get data from timetable

selection = input('Parse new file?(y/n):\n')
if selection == 'y':
    data = Parser()
    with open('filename.pickle', 'wb') as output:
        pickle.dump(data, output)
elif selection == 'n':
    with open('filename.pickle', 'rb') as handle:
        data = pickle.load(handle)
else:
    print('\nError\nIncorrect input!')
# groups = input("Введите нужные группы через запятую с пробелом, точно как в расписании: ") \
#     .strip().lower().split(", ")

groups = ['Илья Г.:09-933 (1)',
          'Данил:09-933 (1)',
          'Досхан:09-932 (1)',
          'Ахад:09-063 (1)',
          'Максим:09-012 (1)',
          'Илья К.:09-012 (1)', ]


# Copy template

out_file = 'output.xlsx'
shutil.copyfile('template.xlsx', out_file)

# Create area to put lessons ( sry about this ;) )

magic_var = np.array([i for i in range(6, 77, 14)])

group_coord = 'C'

wb = openpyxl.load_workbook(out_file)
sheet = wb[wb.sheetnames[1]]


# Getting lessons for all groups

all_lessons = {}

for group in groups:
    lessons_for_group = data.get_lessons_by_group(group.split(':')[1])
    all_lessons[group] = lessons_for_group


# Create table for workers
print("Start generate table for workers")
# print(all_lessons.keys())
for group, lessons in all_lessons.items():
    sheet[group_coord + "1"] = group.split(':')[0]
    sheet[group_coord + "2"] = group.split(':')[1]
    for lesson in lessons:
        dw_index = days_of_week.index(lesson.day_of_week)
        try:
            time_index = times.index(lesson.duration)

            if lesson.day_of_week == 'сб' and times.index(lesson.duration) > 3:
                continue
            lesson_row = magic_var[dw_index] + time_index * 2

            if 'н/н' in lesson.cell.value:
                sheet[group_coord + str(lesson_row)].fill = PatternFill("solid", start_color=bad_color)
            elif 'ч/н' in lesson.cell.value:
                sheet[group_coord + str(lesson_row + 1)].fill = PatternFill("solid", start_color=bad_color)
            else:
                sheet[group_coord + str(lesson_row)].fill = PatternFill("solid", start_color=bad_color)
                sheet[group_coord + str(lesson_row + 1)].fill = PatternFill("solid", start_color=bad_color)
        except:
            write_in_file('log.txt', f'{lesson.cell.coordinate, lesson.cell.value} is a lesson after 19:30')
            continue

    group_coord = chr(ord(group_coord) + 1)

print("Table for worker is done")
print("Start create table by classroom")

all_lessons_by_classroom = data.get_lessons_by_classrooms(our_classrooms)

# for i in all_lessons_by_classroom['802']:
#     print(i.cell.value)
# magic_var = np.array([i for i in range(4, 76, 14)])

for classroom, lessons in all_lessons_by_classroom.items():
    classroom_index = our_classrooms.index(classroom)
    column = classrooms_columns[classroom_index]
    paint_cells(wb[wb.sheetnames[0]], lessons, good_color, magic_var, column)
print('Select a folder to save the file')

while True:
    directory = tkinter.filedialog.askdirectory()
    if directory:
        wb.save(f'{directory}/{out_file}')
        break
    else:
        print('Error!')

print(f'Worked in {time.time() - start_time}')
