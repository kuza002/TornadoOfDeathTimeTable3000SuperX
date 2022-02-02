from openpyxl.styles import PatternFill

from parse_time_table import Parser
import xlsxwriter
import shutil
import pickle
import openpyxl
import numpy as np

days_of_week = ['пн', 'вт', 'ср', 'чт', 'пт', 'сб']
times = ['8.30-10.00', '10.10-11.40', '11.50-13.20', '14.00-15.30', '15.40-17.10', '17.50-19.20']
good_color = 'ccffcc'
bad_color = 'ffadd6'

# Get data from timetable
# data = Parser()

# with open('filename.pickle', 'wb') as output:
#     pickle.dump(data, output)

with open('filename.pickle', 'rb') as handle:
    data = pickle.load(handle)

# groups = input("Введите нужные группы через запятую с пробелом, точно как в расписании: ") \
#     .strip().lower().split(", ")

groups = ['09-933 (1)', '09-833 (1)']

# Getting lessons for all groups

all_lessons = {}

for group in groups:
    lessons_for_group = data.get_lessons_by_group(group)
    all_lessons[group] = lessons_for_group

# Copy template


out_file = 'output.xlsx'
shutil.copyfile('template.xlsx', out_file)

# Create area to put lessons ( sry about this ;) )

magic_var = np.array([i for i in range(6, 85, 14)])

group_coord = 'C'

wb = openpyxl.load_workbook(out_file)
sheet = wb['Расписание лаборантов 21-22 1се']

for group, lessons in all_lessons.items():
    sheet[group_coord + "2"] = group

    for lesson in lessons:
        dw_index = days_of_week.index(lesson.day_of_week)

        time_index = times.index(lesson.duration)

        lesson_row = magic_var[dw_index] + time_index * 2

        if 'н/н' in lesson.cell.value:
            sheet[group_coord + str(lesson_row)].fill = PatternFill("solid", start_color=bad_color)
        elif 'ч/н' in lesson.cell.value:
            sheet[group_coord + str(lesson_row + 1)].fill = PatternFill("solid", start_color=bad_color)
        else:
            sheet[group_coord + str(lesson_row)].fill = PatternFill("solid", start_color=bad_color)
            sheet[group_coord + str(lesson_row + 1)].fill = PatternFill("solid", start_color=bad_color)

    group_coord = chr(ord(group_coord) + 1)

wb.save(out_file)
