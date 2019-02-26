import xlrd
from datetime import time

def floatToTime(t):
    (t) = int((t) * 24 * 3600)
    hour = (t) // 3600
    if hour > 12:
        hour -= 12
        meridiem = "PM"
    else:
        meridiem = "AM"
    minute = ((t)%3600)//60
    if minute == 0:
        minute = "00"

    timeString = str(hour) + ":" + str(minute)
    return (timeString, meridiem)

# ------- import main details --------- #
wb = xlrd.open_workbook('test.xlsx')
main = wb.sheet_by_index(0)

staticData = main.row_values(1)
date = staticData[8]

year = xlrd.xldate_as_tuple(date,0)[0]
month = xlrd.xldate_as_tuple(date,0)[1]
day = xlrd.xldate_as_tuple(date,0)[2]

name = staticData[0]
number = staticData[1]
event = staticData[2]
start = staticData[3]
end = staticData[4]
fau = staticData[5]
cc = staticData[6]
pc = staticData[7]

t = floatToTime(start)
start = t[0]
startXM = t[1]

t = floatToTime(end)
end = t[0]
endXM = t[1]

acti = fau[0:6]
fund = fau[7:12]
func = fau[-2:]



print("Month: " + str(month))
print("Day: " + str(day))
print("Year: " + str(year) + "\n")
print("Name: " + name)
print("Start: " + start + startXM)
print("End: " + end + endXM)
print("fau: " + acti + "-" + fund + "-" + func)
print("cc: " + cc)
print("pc: " + pc)


# ----------- import equipment ------------- #

equipment = wb.sheet_by_index(1)

stuff = []

for item in equipment.col_values(0):
    stuff.append(item)

#delete header
del stuff[0]
print "Equipment:"
for x in range(0, len(stuff)):
    print stuff[x]

# ----------- import Building/Floors/Room ------------- #

locations = wb.sheet_by_index(2)
buildings = []
floors = []
rooms = []

for x in range(1, locations.nrows):
    buildings.append(locations.cell(x, 0))
    floors.append(locations.cell(x, 1))
    rooms.append(locations.cell(x,2))




for x in range(0, len(buildings)):
    print("Building: " + buildings[x].value + "\n" + "floors: " + floors[x].value + "\n" + "rooms: " + (str(rooms[x].value))[0:-2] + "\n")

# pythonic wow ;)
# print (str(rooms[x].value))[0:-2]
