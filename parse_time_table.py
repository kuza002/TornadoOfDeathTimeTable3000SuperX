from openpyxl import load_workbook
import re

# CAN BE CHANGE IN NEW VERSION TIMETABLE!!!!!!!!!!!!!
pattern_to_finde_class_number = r'ауд. (\d+)'
short_list_day_of_week = ['пн', 'вт', 'ср', 'чт', 'пт', 'сб']
list_day_of_week = ["понедельник", "вторник", "среда", "четверг", "пятница", "суббота"]
patter_to_define_time = r'^(([0-1]?[0-9]|2[0-3])(\.|:)[0-5][0-9])-(([0-1]?[0-9]|2[0-3])(\.|:)[0-5][0-9])$'


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
        return re.search(pattern_to_finde_class_number, text).group(1)
    except:
        return None


def day_of_week_by_coordinate(coord, dw_rows):
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


def get_tabletime_for_classrums(class_numbers, data):
    classrums = []
    for i in data:
        tmp = get_class_number_from_cell(i.value)
        if tmp in class_numbers:
            classrums.append(i)


class Parser:
    def __init__(self):

        # Get workbook
        wb = load_workbook('file.xlsx')

        # Get active page
        sheet = wb.active

        # Get data from sheet
        data = []
        days_of_week = []
        groups = []
        time_rows = {}

        # CAN BE CHANGE IN NEW VERSION TIMETABLE!!!!!!!!!!!!!
        for cellObj in sheet['A20':'FV108']:
            for cell in cellObj:
                value = str(cell.value).lower().strip()

                if is_day_of_week(value):
                    days_of_week.append(cell)

                # CAN BE CHANGE IN NEW VERSION TIMETABLE!!!!!!!!!!!!!
                elif value[0:2] == '09':
                    groups.append(cell)

                elif re.search(patter_to_define_time, value):
                    time_rows[get_row_from_coordinate(cell.coordinate)] = value

                elif value != None and value != '' and value != 'none':
                    data.append(cell)

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
            groups_columns[i.value.lower().strip()] = column

        # Delete no important from data to create table with lessons

        lessons = []
        faculty_row = str(int(get_row_from_coordinate(groups[0].coordinate)) - 1)

        for cell in data:
            value = str(cell.value).lower().strip()

            if get_row_from_coordinate(cell.coordinate) == faculty_row:
                continue
            elif value in list_day_of_week:
                continue
            lessons.append(cell)

        self.lessons = lessons
        self.groups_columns = groups_columns
        self.time_rows = time_rows

    def get_lessons_by_group(self, group):
        lessons_of_group = []
        if group in self.groups_columns.keys():
            for i in self.lessons:
                if get_column_from_coordinate(i.coordinate) == self.groups_columns[group]:
                    lessons_of_group.append(i)

        else:
            print(f"\nГруппа {group} не найдена")
            exit(400)

        return lessons_of_group
