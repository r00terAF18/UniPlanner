classes = [i for i in range(0, 25)]

sat = []
sun = []
mon = []
tue = []
wed = []

times = ['8-10', '10-12', '12-14', '14-16', '16-18', '18-20']
cells = ['E', 'D', 'C', 'B', 'A']
cell_num = [i for i in range(2, 8)]


def fillClass(weekday: list, day: str):
    for time in times:
        c = input(f"{day} at {time} -> ")
        weekday.append(c)

fillClass(sat, "Saturday")
fillClass(sun, "Sunday")
fillClass(mon, "Monday")
fillClass(tue, "Tuesday")
fillClass(wed, "Wednesday")

weekdays = [sat, sun, mon, tue, wed]

from openpyxl import load_workbook

workbook = load_workbook("template.xlsx")
sheet = workbook.active

def writeXl(weekday: list, cell: str):
    for i, time in enumerate(cell_num):
        sheet[f'{cell}{time}'] = weekday[i]
        

writeXl(sat, 'E')
writeXl(sun, 'D')
writeXl(mon, 'C')
writeXl(tue, 'B')
writeXl(wed, 'A')

fname = input("Student Name: ")
workbook.save(filename=f"{fname}.xlsx")
print("\n\tBye Bye Ma'fucker\n")