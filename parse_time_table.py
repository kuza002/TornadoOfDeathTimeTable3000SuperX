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
        # CAN BE CHANGE IN NEW VERSION TIMETABLE!!!!!!!!!!!!!

        for work_time, cellObj in enumerate(self.sheet):
            for cell in cellObj:
                value = str(cell.value).lower().strip()

                if value in short_list_day_of_week:
                    days_of_week.append(cell)

                # CAN BE CHANGE IN NEW VERSION TIMETABLE!!!!!!!!!!!!!
                elif '09-' in value:
                    groups.append(cell)

                elif re.search(patter_to_define_time, value):
                    time_rows[self.get_row_from_coordinate(cell.coordinate)] = value

                elif type(cell) == openpyxl.cell.cell.MergedCell:
                    data.append(cell)
                    # print("Merged cell", cell, "is added")
                elif value != 'none':
                    data.append(cell)
                    # print("Not null cell", cell, "is added")

        # Get the row of the days of the week
        dw_rows = {}
        for item in days_of_week:
            prev = item.value.lower()
            dw_rows[prev] = self.get_row_from_coordinate(item.coordinate)

        # Get columns of the groups

        groups_columns = {}

        for item in groups:
            column = self.get_column_from_coordinate(item.coordinate)
            groups_columns[item.value.lower().strip()] = column

        # Delete no important from data to create table with lessons

        print("Getting lessons")

        lessons = []
        faculty_row = str(int(self.get_row_from_coordinate(groups[0].coordinate)) - 1)
        data_len = len(data)
        for work_time, cell in enumerate(data):
            value = str(cell.value).lower().strip()

            if self.get_row_from_coordinate(cell.coordinate) == faculty_row:
                continue
            elif value in list_day_of_week:
                continue

            # add all merged group
            has_merged_cells = False
            for rng in self.sheet.merged_cells.ranges:
                if cell.coordinate in rng:

                    ab = str(rng).split(':')

                    a = ab[0]
                    b = ab[1]

                    a_row = self.get_row_from_coordinate(a)
                    b_row = self.get_row_from_coordinate(b)

                    if a_row == b_row or self.get_row_from_coordinate(cell.coordinate) == a_row:
                        new_cell = openpyxl.cell.Cell(self.sheet, cell.row, cell.column,
                                                      self.sheet[str(rng)][0][0].value)
                        lessons.append(new_cell)

                    has_merged_cells = True

            if not has_merged_cells:
                lessons.append(cell)

            if work_time % 100 == 0:
                print(f'{work_time}/{data_len}')

        print(f'{data_len}/{data_len}')

        self.lessons = []

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
                                                   self.get_day_of_week(new_cell.coordinate, dw_rows),
                                                   self.get_group_number(new_cell, groups_columns)))
                except:
                    write_in_file('some_files/log.txt', f'"{lesson.coordinate, lesson.value}" is not choice lesson\n')
            else:
                try:
                    self.lessons.append(Lesson(lesson,
                                               self.get_duration(lesson, time_rows),
                                               self.get_day_of_week(lesson.coordinate, dw_rows),
                                               self.get_group_number(lesson, groups_columns)))
                except:
                    write_in_file('some_files/log.txt', f'"{lesson.coordinate, lesson.value}" is not lesson\n')
        self.groups_columns = groups_columns
        print(f'Worked in {time.time() - start_time}')

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
    def get_row_from_coordinate(coord):
        row = ""
        for i in coord:
            if i.isdigit():
                row += i
        return row

    @staticmethod
    def get_column_from_coordinate(coord):
        column = ""
        for i in coord:
            if i.isalpha():
                column += i
        return column

    def get_day_of_week(self, coord, dw_rows):
        row = int(self.get_row_from_coordinate(coord))
        day_of_week = ""

        dw_rows["вс"] = '100000'

        for k, v in dw_rows.items():
            v = int(v)

            if v > row:
                break
            else:
                day_of_week = k

        if day_of_week == "":
            return "пн"

        return day_of_week

    def get_duration(self, coordinate, time_rows):
        row = self.get_row_from_coordinate(coordinate.coordinate)
        for key, value in time_rows.items():
            if key == row or key == str(int(row) - 1):
                return value
        write_in_file('some_files/log.txt', f'{coordinate} dont have a time\n')
        return None

    def get_group_number(self, cell, groups_columns):
        for key, value in groups_columns.items():
            if self.get_column_from_coordinate(cell.coordinate) == value:
                return key
    # endregion


@thread
def test(savename, name):
    a = Parser(name)
    with open(savename, 'wb') as output:
        pickle.dump(a, output)
    return


if __name__ == '__main__':
    print('parser class file!')
