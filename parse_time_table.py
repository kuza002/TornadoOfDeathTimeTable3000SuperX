import openpyxl.cell
from openpyxl import load_workbook
from openpyxl import cell
import re
from Entitys import Lesson

# CAN BE CHANGE IN NEW VERSION TIMETABLE!!!!!!!!!!!!!
pattern_to_finde_class_number = r'ауд. (\d+)'
short_list_day_of_week = ['пн', 'вт', 'ср', 'чт', 'пт', 'сб']
list_day_of_week = ["понедельник", "вторник", "среда", "четверг", "пятница", "суббота"]
patter_to_define_time = r'^(([0-1]?[0-9]|2[0-3])(\.|:)[0-5][0-9])-(([0-1]?[0-9]|2[0-3])(\.|:)[0-5][0-9])$'
UP_LEFT_POINT = 'A15'
DOWN_RIGHT_POINT = 'FR103'
INSTITUTE_PREFIX = '09'
pattern_to_detect_group_number = INSTITUTE_PREFIX+r'-\d{3}( \(\d\))?'


def write_in_file(path, var):
    with open(path, 'a') as file:
        file.write(var)


def is_day_of_week(text):
    text.lower()
    return text in short_list_day_of_week


def get_row_from_coordinate(coord):
    row = ""
    for i in coord:
        if i.isdigit():
            row += i
    return row


def get_column_from_coordinate(coord):
    column = ""
    for i in coord:
        if i.isalpha():
            column += i
    return column


def get_class_number_from_cell(text):
    try:
        return re.search(pattern_to_finde_class_number, text).group(1).strip().lower()
    except:
        return None


def get_day_of_week(coord, dw_rows):
    row = int(get_row_from_coordinate(coord))
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


def get_duration(coordinate, time_rows):
    row = get_row_from_coordinate(coordinate.coordinate)
    for key, value in time_rows.items():
        if key == row or key == str(int(row) - 1):
            return value
    print("Failed to find row with time")
    return None


class Parser:

    def get_group_number(self, cell, groups_columns):

        for key, value in groups_columns.items():
            if get_column_from_coordinate(cell.coordinate) == value:
                return key

    def __init__(self):
        print("Start parsing")
        # Get workbook
        wb = load_workbook('file.xlsx')

        # Get active page
        self.sheet = wb.active

        # Get data from sheet
        data = []
        days_of_week = []
        groups = []
        time_rows = {}

        for work_time, cellObj in enumerate(self.sheet[UP_LEFT_POINT:DOWN_RIGHT_POINT]):
            for cell in cellObj:
                value = str(cell.value).lower().strip()

                if is_day_of_week(value):
                    days_of_week.append(cell)

                # CAN BE CHANGE IN NEW VERSION TIMETABLE!!!!!!!!!!!!!
                elif value[0:2] == INSTITUTE_PREFIX:
                    groups.append(cell)

                elif re.search(patter_to_define_time, value):
                    time_rows[get_row_from_coordinate(cell.coordinate)] = value

                elif type(cell) == openpyxl.cell.cell.MergedCell:
                    data.append(cell)
                    # print("Merged cell", cell, "is added")
                elif value != 'none':
                    data.append(cell)
                    # print("Not null cell", cell, "is added")

            if work_time % 10 == 0:
                print(work_time)

        # Get the row of the days of the week
        prev = days_of_week[0].value
        dw_rows = {prev: get_row_from_coordinate(days_of_week[0].coordinate)}

        for i in days_of_week:
            if i.value != prev:
                prev = i.value.lower()
                dw_rows[prev] = get_row_from_coordinate(i.coordinate)

        # Get columns of the groups

        groups_columns = {}

        for i in groups:
            column = get_column_from_coordinate(i.coordinate)

            # Validate the key which is the group number
            key = i.value.lower().strip()
            key = re.search(pattern_to_detect_group_number, key).group(0)

            groups_columns[key] = column

        # Delete no important from data to create table with lessons

        print("Getting lessons")

        lessons = []
        faculty_row = str(int(get_row_from_coordinate(groups[0].coordinate)) - 1)
        for work_time, cell in enumerate(data):
            value = str(cell.value).lower().strip()

            if get_row_from_coordinate(cell.coordinate) == faculty_row:
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

                    a_row = get_row_from_coordinate(a)
                    b_row = get_row_from_coordinate(b)

                    if a_row == b_row or get_row_from_coordinate(cell.coordinate) == a_row:
                        new_cell = openpyxl.cell.Cell(self.sheet, cell.row, cell.column,
                                                      self.sheet[str(rng)][0][0].value)
                        lessons.append(new_cell)

                    # if new_cell.coordinate == "W56":
                    #     print(new_cell.value)
                    # write_in_file('tmp.txt', new_cell.coordinate+" "+new_cell.value+"\n")
                    has_merged_cells = True
            if not has_merged_cells:
                lessons.append(cell)

            if work_time % 100 == 0:
                print(work_time)

        self.lessons = []
        for lesson in lessons:
            self.lessons.append(Lesson(lesson,
                                       get_duration(lesson, time_rows),
                                       get_day_of_week(lesson.coordinate, dw_rows),
                                       self.get_group_number(lesson, groups_columns)))

        self.groups_columns = groups_columns

    # sry about this (

    def get_lessons_by_group(self, group):
        lessons_of_group = []
        if group in self.groups_columns.keys():
            for lesson in self.lessons:
                if lesson.group_number == group:
                    lessons_of_group.append(lesson)
            return lessons_of_group
        else:
            print(f"\nГруппа {group} не найдена")
            exit(400)

    def get_lessons_by_classrooms(self, class_numbers):
        classrooms = {}
        for lesson in self.lessons:
            classroom = get_class_number_from_cell(lesson.cell.value)
            if classroom in class_numbers:
                if classroom in classrooms.keys():
                    classrooms[classroom].append(lesson)
                else:
                    classrooms[classroom] = []
        return classrooms
