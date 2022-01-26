from parse_time_table import Parser


# Get data from timetable
data = Parser()

groups = input("Введите нужные группы через запятую с пробелом, точно как в расписании: ")\
    .strip().lower().split(", ")


all_groups = data.groups_columns.keys()

time_table = None

for group in groups:
    time_table = data.get_lessons_by_group(group)

for i in time_table:
    print(i.value)
