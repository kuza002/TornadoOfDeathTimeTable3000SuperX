import re
import time

import openpyxl.cell
from openpyxl import load_workbook
from support_file import write_in_file, Lesson

# CAN BE CHANGE IN NEW VERSION TIMETABLE!!!!!!!!!!!!!
pattern_to_finde_class_number = r'ауд.\s*(\d+)'
short_list_day_of_week = ['пн', 'вт', 'ср', 'чт', 'пт', 'сб']
list_day_of_week = ["понедельник", "вторник", "среда", "четверг", "пятница", "суббота"]
patter_to_define_time = r'^(([0-1]?[0-9]|2[0-3])(\.|:)[0-5][0-9])-(([0-1]?[0-9]|2[0-3])(\.|:)[0-5][0-9])$'


class Parser:
    def __init__(self, file_path):
        print("Start parsing")
        start_time = time.time()
        sheet = load_workbook(file_path).active
        data, dow_rows, self.groups_columns, groups_row, time_rows = self.preprocessing(sheet)

        print("Getting lessons")
        lessons = self.get_lessons(data, groups_row, sheet, time_rows)
        self.lessons = self.MergedCell_processing(dow_rows, lessons, sheet, time_rows)
        print(f'Worked in {time.time() - start_time}')

    def MergedCell_processing(self, dow_rows, lessons, sheet, time_rows):
        tmp = []
        for lesson in lessons:
            value = lesson.value
            if 'курсы по выбору:' in value.lower() or 'курс по выбору:' in value.lower():
                try:
                    for item in value.split(';'):
                        if 'н/н' in value and 'ч/н' not in value:
                            if 'н/н' not in item:
                                item = 'н/н ' + item
                        elif 'н/н' not in value and 'ч/н' in value:
                            if 'ч/н' not in item:
                                item = 'ч/н ' + item

                        new_cell = openpyxl.cell.Cell(sheet, row=lesson.row, column=lesson.column, value=item)

                        tmp.append(Lesson(new_cell,
                                          self.get_duration(new_cell, time_rows),
                                          self.get_day_of_week(new_cell.row, dow_rows),
                                          self.get_group_number(new_cell, self.groups_columns)))
                except:
                    write_in_file('some_files/log.txt', f'"{lesson.coordinate, lesson.value}" is not choice lesson\n')
            else:
                try:
                    tmp.append(Lesson(lesson,
                                      self.get_duration(lesson, time_rows),
                                      self.get_day_of_week(lesson.row, dow_rows),
                                      self.get_group_number(lesson, self.groups_columns)))
                except:
                    write_in_file('some_files/log.txt', f'"{lesson.coordinate, lesson.value}" is not lesson\n')

        return tmp

    def get_lessons(self, data, groups_row, sheet, time_rows):
        lessons = []
        faculty_row = groups_row - 1
        data_len = len(data)
        for work_time, cell in enumerate(data):
            if self.get_duration(cell, time_rows) is not None and \
                    self.get_group_number(cell, self.groups_columns) is not None:
                if cell.row == faculty_row:
                    continue
                elif str(cell.value).lower().strip() in list_day_of_week:
                    continue
                elif type(cell).__name__ == 'MergedCell':
                    index = self.get_index_for_mergedcell(cell, sheet.merged_cell_ranges)
                    new_cell = openpyxl.cell.Cell(sheet, cell.row, cell.column,
                                                  sheet[index].value)
                    lessons.append(new_cell)
                else:
                    lessons.append(cell)

                if work_time % 100 == 0:
                    print(f'{work_time}/{data_len}')
        print(f'{data_len}/{data_len}')
        return lessons

    def preprocessing(self, sheet):
        data = []
        time_rows = {}
        groups_columns = {}
        dow_rows = {}
        groups_row = -1
        for cellObj in sheet:
            for cell in cellObj:
                value = str(cell.value).lower().strip()
                if value in short_list_day_of_week:
                    if value not in dow_rows.keys():
                        dow_rows[value] = cell.row
                elif '09-' in value:
                    if groups_row == -1:
                        groups_row = cell.row
                    column = self.get_column_from_coordinate(cell.coordinate)
                    groups_columns[cell.value.lower().strip()] = column
                elif re.search(patter_to_define_time, value):
                    time_rows[cell.row] = value
                elif type(cell) == openpyxl.cell.cell.MergedCell:
                    data.append(cell)
                elif value != 'none':
                    data.append(cell)
                else:
                    continue
        dow_rows["вс"] = 100000
        return data, dow_rows, groups_columns, groups_row, time_rows

    def get_lessons_by_group(self, group):
        lessons_of_group = []
        if group in self.groups_columns.keys():
            for lesson in self.lessons:
                if lesson.group_number == group:
                    lessons_of_group.append(lesson)
            return lessons_of_group
        else:
            print(f"\nГруппа {group} не найдена")
            return None

    def get_lessons_by_classrooms(self, class_numbers):
        classrooms = {}
        for lesson in self.lessons:
            classroom = self.get_class_number_from_cell(lesson)
            if classroom in class_numbers:
                if classroom in classrooms.keys():
                    classrooms[classroom].append(lesson)
                else:
                    classrooms[classroom] = []
                    classrooms[classroom].append(lesson)
        return classrooms

    # region staticmethods
    @staticmethod
    def get_class_number_from_cell(lesson):
        try:
            return re.search(pattern_to_finde_class_number, lesson.cell.value.lower()).group(1)
        except:
            return None

    @staticmethod
    def get_index_for_mergedcell(cell, merged_ranges):
        for rng in merged_ranges:
            if cell.coordinate in rng:
                index = rng.coord.split(':')[0]
                return index

    @staticmethod
    def get_column_from_coordinate(coord):
        return ''.join(filter(lambda x: x if x.isalpha() else '', coord))

    @staticmethod
    def get_day_of_week(row, dw_rows):
        day_of_week = ""
        for k, v in dw_rows.items():
            v = v
            if v > row:
                break
            else:
                day_of_week = k

        if day_of_week == "":
            return "пн"

        return day_of_week

    @staticmethod
    def get_duration(cell, time_rows):
        for key, value in time_rows.items():
            if key == cell.row or key == (cell.row - 1):
                return value
        return None

    def get_group_number(self, cell, groups_columns):
        for key, value in groups_columns.items():
            if self.get_column_from_coordinate(cell.coordinate) == value:
                return key
        return None
    # endregion


if __name__ == '__main__':
    print('parser class file!')
