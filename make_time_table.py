from parse_time_table import Parser


# Get data from timetable
data = Parser()

groups = input("Введите нужные группы через запятую с пробелом, точно как в расписании: ")\
    .strip().lower().split(", ")


# all_groups = data.groups_columns.keys()

lessons = []

for group in groups:
    lessons_for_group = data.get_lessons_by_group(group)
    lessons.append(lessons_for_group)

for lessons_for_group in lessons:
    for lesson in lessons_for_group:
        print(lesson.cell.value)
