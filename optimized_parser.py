import pickle
import re
import threading
import time
import openpyxl.cell
from openpyxl import load_workbook

from support_file import write_in_file, Lesson

# CAN BE CHANGE IN NEW VERSION TIMETABLE!!!!!!!!!!!!!
pattern_to_finde_class_number = r'ауд.\s*(\d+)'
short_list_day_of_week = ['пн', 'вт', 'ср', 'чт', 'пт', 'сб']
list_day_of_week = ["понедельник", "вторник", "среда", "четверг", "пятница", "суббота"]
patter_to_define_time = r'^(([0-1]?[0-9]|2[0-3])(\.|:)[0-5][0-9])-(([0-1]?[0-9]|2[0-3])(\.|:)[0-5][0-9])$'


def thread(fn):
    def execute(*args, **kwargs):
        threading.Thread(target=fn, args=args, kwargs=kwargs).start()

    return execute


class Parser:
    def __init__(self, file_path):
        # Get workbook
        start_time = time.time()
        wb = load_workbook(file_path)

        print("Start parsing")
        # Get active page
        self.sheet = wb.active
        # wb.close()

        # Get data from sheet
        data = []
        days_of_week = []
        groups = []
        time_rows = {}
        groups_columns = {}
        dow_rows = {}
        groups_row=-1
        # CAN BE CHANGE IN NEW VERSION TIMETABLE!!!!!!!!!!!!!

        for cellObj in self.sheet:
            for cell in cellObj:
                value = str(cell.value).lower().strip()
                if value in short_list_day_of_week:
                    if value not in dow_rows.keys():
                        dow_rows[value] = cell.row
                # CAN BE CHANGE IN NEW VERSION TIMETABLE!!!!!!!!!!!!!
                elif '09-' in value:
                    if groups_row==-1:
                        groups_row=cell.row
                    column = self.get_column_from_coordinate(cell.coordinate)
                    groups_columns[cell.value.lower().strip()] = column
                elif re.search(patter_to_define_time, value):
                    time_rows[cell.row] = value
                elif type(cell) == openpyxl.cell.cell.MergedCell:
                    data.append(cell)
                    # print("Merged cell", cell, "is added")
                elif value != 'none':
                    data.append(cell)
                else:
                    continue
        dow_rows["вс"] = 100000
        print("Getting lessons")

        lessons = []
        faculty_row = groups_row - 1
        data_len = len(data)
        for work_time, cell in enumerate(data):
            if self.get_duration(cell, time_rows) != None and self.get_group_number(cell, groups_columns) != None:
                if cell.row == faculty_row:
                    continue
                elif str(cell.value).lower().strip() in list_day_of_week:
                    continue
                elif type(cell).__name__ == 'MergedCell':
                    index = self.get_value_for_mergedcell(cell)
                    new_cell = openpyxl.cell.Cell(self.sheet, cell.row, cell.column,
                                                  self.sheet[index].value)
                    lessons.append(new_cell)
                else:
                    lessons.append(cell)

                if work_time % 100 == 0:
                    print(f'{work_time}/{data_len}')

        print(f'{data_len}/{data_len}')

        self.lessons = []
        # print(len(lessons))
        # print(len(set(lessons)))
        # return
        for lesson in lessons:
            if 'курсы по выбору:' or 'курс по выбору:' in lesson.value.strip().lower():
                try:
                    for item in lesson.value.split(';'):
                        if 'н/н' in lesson.value and 'ч/н' not in lesson.value:
                            if 'н/н' not in item:
                                item = 'н/н ' + item
                        elif 'н/н' not in lesson.value and 'ч/н' in lesson.value:
                            if 'ч/н' not in item:
                                item = 'ч/н ' + item

                        new_cell = openpyxl.cell.Cell(self.sheet, row=lesson.row, column=lesson.column, value=item)

                        self.lessons.append(Lesson(new_cell,
                                                   self.get_duration(new_cell, time_rows),
                                                   self.get_day_of_week(new_cell.row, dow_rows),
                                                   self.get_group_number(new_cell, groups_columns)))
                except:
                    write_in_file('some_files/log.txt', f'"{lesson.coordinate, lesson.value}" is not choice lesson\n')
            else:
                try:
                    self.lessons.append(Lesson(lesson,
                                               self.get_duration(lesson, time_rows),
                                               self.get_day_of_week(lesson.row, dow_rows),
                                               self.get_group_number(lesson, groups_columns)))
                except:
                    write_in_file('some_files/log.txt', f'"{lesson.coordinate, lesson.value}" is not lesson\n')
        self.groups_columns = groups_columns
        print(f'Worked in {time.time() - start_time}')

    def get_value_for_mergedcell(self, cell):
        for rng in self.sheet.merged_cells.ranges:
            if cell.coordinate in rng:
                index = rng.coord.split(':')[0]
                return index

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
            # if lesson.cell.coordinate == 'DQ82':
            # print(lesson.cell.value)
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
    def get_column_from_coordinate(coord):
        return ''.join(filter(lambda x: x if x.isalpha() else '', coord))

    def get_day_of_week(self, row, dw_rows):
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

    def get_duration(self, cell, time_rows):
        for key, value in time_rows.items():
            if key == cell.row or key == (cell.row - 1):
                return value
        write_in_file('some_files/log.txt', f'{cell.coordinate} dont have a time\n')
        return None

    def get_group_number(self, cell, groups_columns):
        for key, value in groups_columns.items():
            if self.get_column_from_coordinate(cell.coordinate) == value:
                return key
        return None
    # endregion


@thread
def test(savename, name):
    a = Parser(name)
    with open(savename, 'wb') as output:
        pickle.dump(a, output)
    return


if __name__ == '__main__':
    print('parser class file!')
